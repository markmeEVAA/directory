// EVAA Finance — Treasurer Console
// Tabs:
//   - Requests  : table + detail drawer + status workflow + Export CSV
//   - Options   : edit FinanceFormOptions (dropdown source of truth + account codes)

(async () => {
  const $ = (id) => document.getElementById(id);
  const show = (id) =>
    ["view-loading", "view-signin", "view-no-access", "view-console"].forEach((v) =>
      $(v).classList.toggle("hidden", v !== id)
    );

  // Display labels for option types (used by the Options-tab renderer).
  // Declared up here so it's accessible from the first renderOptionsEditor()
  // call (which fires before the bottom of the IIFE finishes — temporal-dead-zone trap).
  const OPTION_TYPE_LABELS = {
    Sport: "Sports",
    ExpenseCategory: "Expense / income categories",
    ProgramType: "Program types",
    TravelingSubtype: "Traveling subtypes",
    Season: "Seasons",
    VendorCardinality: "Vendor cardinality",
    RequestType: "Request types",
  };

  // Status workflow
  const STATUS_ORDER = ["Submitted", "Under Review", "Approved", "Paid/Deposited", "Exported", "Denied"];
  const STATUS_CLASS = {
    "Submitted":      "status-Submitted",
    "Under Review":   "status-UnderReview",
    "Approved":       "status-Approved",
    "Paid/Deposited": "status-PaidDeposited",
    "Exported":       "status-Exported",
    "Denied":         "status-Denied",
  };
  const NEXT = {
    "Submitted":      ["Under Review", "Approved", "Denied"],
    "Under Review":   ["Approved", "Denied"],
    "Approved":       ["Paid/Deposited", "Denied"],
    "Paid/Deposited": ["Exported"],
    "Exported":       [],
    "Denied":         ["Under Review"],
  };

  // ─── Boot ──────────────────────────────────────────────────────────────────
  show("view-loading");

  let account;
  try { account = await AUTH.init({ consoleMode: true }); }
  catch (e) { $("loading-text").textContent = "Auth init failed: " + e.message; return; }

  if (!account) {
    const btn = $("sign-in-btn");
    btn.disabled = false;
    btn.addEventListener("click", async () => {
      // Redirect flow: navigates away; page re-boots on return and init() picks up the account.
      try { await AUTH.signIn(); }
      catch (e) { $("signin-error").textContent = e.message; $("signin-error").classList.remove("hidden"); }
    });
    show("view-signin");
    return;
  }

  let me;
  try { me = await GRAPH.getMe(); }
  catch (e) { me = { displayName: account.name || account.username, mail: account.username }; }
  $("user-name").textContent = me.displayName || me.userPrincipalName || "";
  $("user-area").classList.remove("hidden");
  $("sign-out-btn").addEventListener("click", () => AUTH.signOut());
  $("signout-noaccess-btn").addEventListener("click", () => AUTH.signOut());

  $("loading-text").textContent = "Checking access…";
  let canAccess;
  try { canAccess = await GRAPH.isTreasurerOrAdmin(); }
  catch (e) { canAccess = false; }
  if (!canAccess) { show("view-no-access"); return; }

  // Load list + options in parallel
  $("loading-text").textContent = "Loading requests…";
  let items = [], options = [];
  try {
    [items, options] = await Promise.all([
      GRAPH.listFinanceRequests(),
      GRAPH.getFinanceFormOptions({ force: true }),
    ]);
  } catch (e) {
    $("loading-text").innerHTML = `<span class="alert alert-error">Failed to load: ${escapeHtml(e.message)}</span>`;
    return;
  }

  // Populate dynamic filters from options
  const typeSel = $("filter-type");
  options.filter((o) => o.type === "RequestType" && o.active)
    .sort((a, b) => a.order - b.order)
    .forEach((o) => { const opt = document.createElement("option"); opt.value = o.title; opt.textContent = o.title; typeSel.appendChild(opt); });
  const sportSel = $("filter-sport");
  options.filter((o) => o.type === "Sport" && o.active)
    .sort((a, b) => a.order - b.order)
    .forEach((o) => { const opt = document.createElement("option"); opt.value = o.title; opt.textContent = o.title; sportSel.appendChild(opt); });

  // Map: category title -> account code (for quick lookup in render + export)
  const categoryToCode = {};
  options.filter((o) => o.type === "ExpenseCategory")
    .forEach((o) => { categoryToCode[o.title] = o.code || ""; });

  ["filter-type", "filter-status", "filter-sport", "filter-search"].forEach((id) =>
    $(id).addEventListener("input", renderRows)
  );

  $("refresh-btn").addEventListener("click", refresh);
  $("export-btn").addEventListener("click", () => openExportDialog());

  // Tab switching
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

  show("view-console");
  renderRows();
  renderOptionsEditor(); // pre-render so tab switch is instant

  // Deep-link: ?id=NN
  const params = new URLSearchParams(window.location.search);
  const deepLinkId = params.get("id");
  if (deepLinkId) {
    const target = items.find((it) => String(it.id) === String(deepLinkId));
    if (target) openDetail(target);
  }

  async function refresh() {
    $("loading-text").textContent = "Refreshing…";
    show("view-loading");
    try {
      [items, options] = await Promise.all([
        GRAPH.listFinanceRequests(),
        GRAPH.getFinanceFormOptions({ force: true }),
      ]);
      Object.keys(categoryToCode).forEach((k) => delete categoryToCode[k]);
      options.filter((o) => o.type === "ExpenseCategory").forEach((o) => { categoryToCode[o.title] = o.code || ""; });
      show("view-console");
      renderRows();
      renderOptionsEditor();
    } catch (e) {
      $("loading-text").innerHTML = `<span class="alert alert-error">${escapeHtml(e.message)}</span>`;
    }
  }

  function switchTab(name) {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("hidden", p.dataset.panel !== name));
  }

  // ─── Requests tab ──────────────────────────────────────────────────────────
  function applyFilters() {
    const t = $("filter-type").value;
    const s = $("filter-status").value;
    const sp = $("filter-sport").value;
    const q = $("filter-search").value.trim().toLowerCase();
    return items.filter((it) => {
      const f = it.fields || {};
      if (t && f.RequestType !== t) return false;
      if (s === "__open") {
        if (!["Submitted", "Under Review", "Approved"].includes(f.Status)) return false;
      } else if (s === "__unexported") {
        if (f.Status !== "Paid/Deposited") return false;
      } else if (s && f.Status !== s) return false;
      if (sp && f.Sport !== sp) return false;
      if (q) {
        const hay = [
          f.SubmittedByName, f.SubmittedByEmail, f.VendorName, f.PayerName,
          f.AltPayeeName, f.Notes, f.TreasurerNotes, f.Title, f.Category,
        ].filter(Boolean).join(" ").toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function renderRows() {
    const filtered = applyFilters();
    filtered.sort((a, b) => (b.createdDateTime || "").localeCompare(a.createdDateTime || ""));
    $("result-summary").textContent =
      `${filtered.length} of ${items.length} requests` +
      (filtered.length !== items.length ? " (filtered)" : "");

    const html = filtered.map((it) => {
      const f = it.fields || {};
      const payee = f.VendorName || f.PayerName || f.AltPayeeName || (f.VendorCardinality === "Multiple Vendors" ? "Various" : "—");
      const statusClass = STATUS_CLASS[f.Status] || "status-Submitted";
      const code = categoryToCode[f.Category] || "";
      const catCell = f.Category ? `${escapeHtml(f.Category)}${code ? `<br><span style="color:#888;font-size:0.85em;font-variant-numeric:tabular-nums;">${escapeHtml(code)}</span>` : ""}` : "";
      return `<tr data-id="${escapeAttr(it.id)}">
        <td>${formatDate(f.DateRequested || it.createdDateTime)}</td>
        <td>${escapeHtml(f.SubmittedByName || "")}<br><span style="color:#888; font-size:0.9em;">${escapeHtml(f.SubmittedByEmail || "")}</span></td>
        <td>${escapeHtml(f.RequestType || "")}</td>
        <td>${escapeHtml(f.Sport || "")}</td>
        <td>${catCell}</td>
        <td>${escapeHtml(payee)}</td>
        <td style="text-align:right; font-variant-numeric: tabular-nums;">${formatMoney(f.Amount)}</td>
        <td><span class="status-pill ${statusClass}">${escapeHtml(f.Status || "Submitted")}</span></td>
      </tr>`;
    }).join("") || `<tr><td colspan="8" style="text-align:center; padding:30px; color:#888;">No requests match these filters.</td></tr>`;

    $("rows").innerHTML = html;
    $("rows").querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", () => {
        const it = items.find((x) => String(x.id) === tr.dataset.id);
        if (it) openDetail(it);
      });
    });
  }

  // ─── Detail drawer ─────────────────────────────────────────────────────────
  $("detail-close").addEventListener("click", closeDetail);
  $("detail-overlay").addEventListener("click", (e) => {
    if (e.target === $("detail-overlay")) closeDetail();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !$("detail-overlay").classList.contains("hidden")) closeDetail();
  });

  function closeDetail() {
    $("detail-overlay").classList.add("hidden");
    if (window.location.search.includes("id=")) {
      const u = new URL(window.location);
      u.searchParams.delete("id");
      window.history.replaceState({}, "", u);
    }
  }

  function openDetail(item) {
    const f = item.fields || {};
    const status = f.Status || "Submitted";
    const statusClass = STATUS_CLASS[status] || "status-Submitted";
    const nextStates = NEXT[status] || [];
    const accountCode = categoryToCode[f.Category] || "";

    $("detail-title").textContent = `${f.RequestType || "Request"} #${item.id}`;

    const rows = [
      ["Status",            `<span class="status-pill ${statusClass}">${escapeHtml(status)}</span>`],
      ["Request type",      f.RequestType],
      ["Date requested",    formatDate(f.DateRequested)],
      ["Sport",             f.Sport],
      ["Program type",      f.ProgramType],
      ["Traveling subtype", f.TravelingSubtype],
      ["Season",            f.Season],
      ["Category",          f.Category + (accountCode ? ` <span style="color:#888; font-size:0.85em;">(account ${escapeHtml(accountCode)})</span>` : "")],
      ["Amount",            formatMoney(f.Amount)],
      ["", "<hr style='border:none; border-top:1px solid #eee; margin:8px 0;'/>"],
      ["Requester",         `${escapeHtml(f.SubmittedByName || "")}<br>${escapeHtml(f.SubmittedByEmail || "")}` + (f.SubmitterPhone ? `<br>${escapeHtml(f.SubmitterPhone)}` : "")],
      ["Vendor cardinality",f.VendorCardinality],
      ["Vendor name",       f.VendorName],
      ["Payer name",        f.PayerName],
      ["Multi-vendor note", f.VendorNote_Multi],
      ["Payee address",     [f.PayeeAddress, f.PayeeCity, f.PayeeState, f.PayeeZip].filter(Boolean).join(", ")],
      ["Make check payable to", f.AltPayeeName],
      ["Mail check to",     [f.AltPayeeAddress, f.AltPayeeCity, f.AltPayeeState, f.AltPayeeZip].filter(Boolean).join(", ")],
      ["", "<hr style='border:none; border-top:1px solid #eee; margin:8px 0;'/>"],
      ["Notes",             f.Notes ? `<div style="white-space:pre-wrap;">${escapeHtml(f.Notes)}</div>` : ""],
      ["Receipt",           `<div id="receipt-cell">(looking up…)</div>`],
      ["Source",            f.SourceSystem || "Portal"],
      ["", "<hr style='border:none; border-top:1px solid #eee; margin:8px 0;'/>"],
      ["Treasurer notes",   `<textarea id="treasurer-notes" rows="3" style="width:100%;">${escapeHtml(f.TreasurerNotes || "")}</textarea>`],
      ["Last decision",     f.DecisionBy ? `${escapeHtml(f.DecisionBy)} on ${formatDate(f.DecisionAt)}` : "(no decisions yet)"],
    ].filter(([k, v]) => k === "" || (v !== null && v !== undefined && v !== ""));

    const workflow = nextStates.length
      ? `<div class="status-workflow">
           <strong style="align-self:center;">Move to:</strong>
           ${nextStates.map((s) => `<button class="btn-primary" data-transition="${escapeAttr(s)}">${escapeHtml(s)}</button>`).join("")}
         </div>`
      : `<div class="alert alert-info">Final state — no further transitions.</div>`;

    $("detail-body").innerHTML =
      rows.map(([k, v]) => k === "" ? v : `<div class="row"><div class="k">${escapeHtml(k)}</div><div class="v">${v ?? ""}</div></div>`).join("") +
      workflow +
      `<p id="detail-alert" class="hidden"></p>`;

    $("detail-body").querySelectorAll("button[data-transition]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const newStatus = btn.dataset.transition;
        const notes = ($("treasurer-notes") && $("treasurer-notes").value) || "";
        if (!confirm(`Move request #${item.id} to "${newStatus}"?`)) return;
        $("detail-body").querySelectorAll("button[data-transition]").forEach((b) => (b.disabled = true));
        try {
          await GRAPH.updateRequestStatus(item.id, newStatus, notes);
          const refreshed = await GRAPH.getFinanceRequest(item.id);
          const idx = items.findIndex((x) => String(x.id) === String(item.id));
          if (idx >= 0) items[idx] = refreshed;
          renderRows();
          openDetail(refreshed);
        } catch (e) {
          const a = $("detail-alert");
          a.className = "alert alert-error";
          a.textContent = "Update failed: " + e.message;
          a.classList.remove("hidden");
          $("detail-body").querySelectorAll("button[data-transition]").forEach((b) => (b.disabled = false));
        }
      });
    });

    $("detail-overlay").classList.remove("hidden");

    // Async receipt lookup + inline preview
    GRAPH.getReceiptForRequest(item.id).then((rcpt) => {
      const cell = document.getElementById("receipt-cell");
      if (!cell) return;
      if (!rcpt) { cell.textContent = "(none uploaded)"; return; }
      const isImage = (rcpt.mimeType || "").startsWith("image/");
      const isPdf = (rcpt.mimeType || "") === "application/pdf";
      if (isImage && rcpt.downloadUrl) {
        cell.innerHTML = `<div class="receipt-preview">
          <img src="${escapeAttr(rcpt.downloadUrl)}" alt="${escapeAttr(rcpt.name)}" />
          <div class="filename">${escapeHtml(rcpt.name)}
            ${rcpt.webUrl ? ` · <a href="${escapeAttr(rcpt.webUrl)}" target="_blank" rel="noopener">open in SharePoint</a>` : ""}
          </div>
        </div>`;
      } else if (isPdf && rcpt.webUrl) {
        cell.innerHTML = `<div class="receipt-preview">
          <p style="margin:0 0 8px;">${escapeHtml(rcpt.name)}</p>
          <a class="download-link" href="${escapeAttr(rcpt.webUrl)}" target="_blank" rel="noopener">Open PDF</a>
        </div>`;
      } else if (rcpt.webUrl) {
        cell.innerHTML = `<a href="${escapeAttr(rcpt.webUrl)}" target="_blank" rel="noopener">${escapeHtml(rcpt.name)}</a>`;
      } else {
        cell.textContent = rcpt.name;
      }
    }).catch(() => {
      const cell = document.getElementById("receipt-cell");
      if (cell) cell.textContent = "(lookup failed)";
    });
  }

  // ─── Export to CSV ─────────────────────────────────────────────────────────
  function openExportDialog() {
    const filtered = applyFilters();
    const paidCount = filtered.filter((it) => (it.fields || {}).Status === "Paid/Deposited").length;
    const dialog = document.createElement("div");
    dialog.className = "export-dialog";
    dialog.innerHTML = `
      <div class="card">
        <h3>Export ${filtered.length} rows to CSV</h3>
        <p style="color:#6c7682;">All fields + account codes will be included. Filename includes today's date.</p>
        <div class="checkbox-row">
          <input type="checkbox" id="flip-exported" ${paidCount > 0 ? "checked" : ""} ${paidCount === 0 ? "disabled" : ""}/>
          <label for="flip-exported">
            <strong>Mark exported as <code>Exported</code></strong> — flips the ${paidCount} <code>Paid/Deposited</code>
            ${paidCount === 1 ? "row" : "rows"} in this export to status <code>Exported</code> after the file downloads.
            ${paidCount === 0 ? "<br><em style='color:#888;'>No Paid/Deposited rows in current filter — nothing to flip.</em>" : ""}
          </label>
        </div>
        <div class="actions">
          <button class="btn-secondary" id="export-cancel">Cancel</button>
          <button class="btn-primary" id="export-go">⬇ Download CSV</button>
        </div>
        <p id="export-status" style="color:#6c7682;font-size:0.9em;margin-top:10px;"></p>
      </div>`;
    document.body.appendChild(dialog);
    dialog.querySelector("#export-cancel").addEventListener("click", () => dialog.remove());
    dialog.querySelector("#export-go").addEventListener("click", async () => {
      const flip = dialog.querySelector("#flip-exported").checked;
      const statusEl = dialog.querySelector("#export-status");
      const goBtn = dialog.querySelector("#export-go");
      goBtn.disabled = true;
      statusEl.textContent = "Building CSV…";
      try {
        downloadCsv(filtered);
        statusEl.textContent = "✓ CSV downloaded.";
        if (flip) {
          const toFlip = filtered.filter((it) => (it.fields || {}).Status === "Paid/Deposited");
          for (let i = 0; i < toFlip.length; i++) {
            statusEl.textContent = `Flipping to Exported: ${i + 1}/${toFlip.length}…`;
            try {
              await GRAPH.updateRequestStatus(toFlip[i].id, "Exported", "(auto: marked exported via CSV export)");
            } catch (e) { console.warn("Flip failed for #" + toFlip[i].id, e); }
          }
          statusEl.textContent = `✓ Exported CSV and flipped ${toFlip.length} rows to Exported. Refreshing…`;
          setTimeout(() => { dialog.remove(); refresh(); }, 800);
        } else {
          setTimeout(() => dialog.remove(), 1500);
        }
      } catch (e) {
        statusEl.textContent = "Export failed: " + e.message;
        goBtn.disabled = false;
      }
    });
  }

  function downloadCsv(rows) {
    const headers = [
      "Id", "Status", "RequestType", "DateRequested", "SubmittedByEmail", "SubmittedByName",
      "Sport", "Category", "AccountCode", "ProgramType", "TravelingSubtype", "Season",
      "Amount", "VendorCardinality", "VendorName", "PayerName", "PayeeAddress", "PayeeCity",
      "PayeeState", "PayeeZip", "AltPayeeName", "AltPayeeAddress", "AltPayeeCity",
      "AltPayeeState", "AltPayeeZip", "Notes", "TreasurerNotes", "DecisionAt", "DecisionBy",
      "CreatedAt",
    ];
    const csvRows = [headers.join(",")];
    rows.forEach((it) => {
      const f = it.fields || {};
      const cells = [
        it.id, f.Status, f.RequestType, f.DateRequested, f.SubmittedByEmail, f.SubmittedByName,
        f.Sport, f.Category, categoryToCode[f.Category] || "", f.ProgramType, f.TravelingSubtype, f.Season,
        f.Amount, f.VendorCardinality, f.VendorName, f.PayerName, f.PayeeAddress, f.PayeeCity,
        f.PayeeState, f.PayeeZip, f.AltPayeeName, f.AltPayeeAddress, f.AltPayeeCity,
        f.AltPayeeState, f.AltPayeeZip, f.Notes, f.TreasurerNotes, f.DecisionAt, f.DecisionBy,
        it.createdDateTime,
      ];
      csvRows.push(cells.map(csvEscape).join(","));
    });
    const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const today = new Date().toISOString().slice(0, 10);
    const a = document.createElement("a");
    a.href = url;
    a.download = `evaa-finance-${today}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function csvEscape(v) {
    if (v === null || v === undefined) return "";
    const s = String(v);
    if (s.includes('"') || s.includes(",") || s.includes("\n") || s.includes("\r")) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  // ─── Options editor ────────────────────────────────────────────────────────
  // OPTION_TYPE_LABELS declared at the top of the IIFE so the first render()
  // call (which fires before this section is reached) can see it.

  function renderOptionsEditor() {
    const container = $("options-editor");
    if (!container) { console.error("[finance] options-editor div missing"); return; }
    if (!Array.isArray(options)) {
      container.innerHTML = `<div class="alert alert-error">Options array isn't loaded yet (type=${typeof options}). Refresh the page.</div>`;
      console.error("[finance] options is not an array:", options);
      return;
    }
    console.log(`[finance] renderOptionsEditor: ${options.length} options loaded`);
    const typeOrder = ["Sport", "ExpenseCategory", "ProgramType", "TravelingSubtype", "Season", "VendorCardinality", "RequestType"];
    try {
    container.innerHTML = typeOrder.map((type) => {
      const rows = options
        .filter((o) => o.type === type)
        .sort((a, b) => (a.order - b.order) || a.title.localeCompare(b.title));
      const showCodeCol = type === "ExpenseCategory";
      return `<div class="options-section" data-type="${escapeAttr(type)}">
        <h3>${escapeHtml(OPTION_TYPE_LABELS[type] || type)} <span class="count">${rows.length}</span><span class="save-indicator" data-save="${escapeAttr(type)}"></span></h3>
        <table>
          <thead><tr>
            <th>Option text</th>
            ${showCodeCol ? "<th>Account code</th>" : ""}
            <th class="order-col">Order</th>
            <th class="active-col">Active</th>
            <th class="actions-col"></th>
          </tr></thead>
          <tbody>
            ${rows.map((o) => optionRowHtml(o, showCodeCol)).join("")}
          </tbody>
        </table>
        <button class="btn-secondary btn-add" data-add="${escapeAttr(type)}">+ Add ${escapeHtml((OPTION_TYPE_LABELS[type] || type).toLowerCase().replace(/s$/, ""))}</button>
      </div>`;
    }).join("");
    wireOptionsHandlers();
    } catch (err) {
      console.error("[finance] renderOptionsEditor threw:", err);
      container.innerHTML = `<div class="alert alert-error">Options editor crashed: ${escapeHtml(err.message)}<br><br><pre style="font-size:11px;overflow:auto;">${escapeHtml(String(err.stack || ""))}</pre></div>`;
    }
  }

  function optionRowHtml(opt, showCodeCol) {
    return `<tr data-opt-id="${escapeAttr(opt.id)}">
      <td><input type="text" data-field="title" value="${escapeAttr(opt.title)}" /></td>
      ${showCodeCol ? `<td class="code-col"><input type="text" data-field="code" value="${escapeAttr(opt.code)}" placeholder="e.g. 5100" /></td>` : ""}
      <td class="order-col"><input type="number" data-field="order" value="${opt.order || 0}" /></td>
      <td class="active-col"><input type="checkbox" data-field="active" ${opt.active ? "checked" : ""}/></td>
      <td class="actions-col"><button class="btn-del" data-action="delete" title="Delete">✕</button></td>
    </tr>`;
  }

  function wireOptionsHandlers() {
    document.querySelectorAll(".options-section").forEach((sec) => {
      const type = sec.dataset.type;
      const showCodeCol = type === "ExpenseCategory";

      // Inline edit
      sec.querySelectorAll("tr[data-opt-id] input").forEach((el) => {
        el.addEventListener("change", async () => {
          const tr = el.closest("tr");
          const id = tr.dataset.optId;
          const fields = {};
          tr.querySelectorAll("input[data-field]").forEach((input) => {
            const f = input.dataset.field;
            if (f === "title") fields.Title = input.value;
            else if (f === "code") fields.AccountCode = input.value;
            else if (f === "order") fields.DisplayOrder = Number(input.value) || 0;
            else if (f === "active") fields.Active = input.checked;
          });
          await persistOption(type, () => GRAPH.updateFormOption(id, fields));
        });
      });

      // Delete
      sec.querySelectorAll("button[data-action='delete']").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const tr = btn.closest("tr");
          const id = tr.dataset.optId;
          const title = tr.querySelector("input[data-field='title']").value;
          if (!confirm(`Delete "${title}"? This removes it from the dropdown immediately.`)) return;
          await persistOption(type, () => GRAPH.deleteFormOption(id));
          // Remove row + refresh local cache
          options = options.filter((o) => o.id !== id);
          renderOptionsEditor();
        });
      });

      // Add new
      const addBtn = sec.querySelector(`button[data-add="${type}"]`);
      addBtn.addEventListener("click", async () => {
        const title = prompt(`New ${OPTION_TYPE_LABELS[type] || type} option text:`);
        if (!title) return;
        const order = (options.filter((o) => o.type === type).reduce((m, o) => Math.max(m, o.order), 0)) + 10;
        const code = showCodeCol ? (prompt(`Account code for "${title}" (optional):`, "") || "") : "";
        await persistOption(type, async () => {
          const created = await GRAPH.createFormOption({ title, type, code, order, active: true });
          // Add to local cache
          options.push({
            id: created.id,
            title,
            type,
            code,
            order,
            active: true,
          });
          // Rebuild categoryToCode for instant refresh of requests-tab UI
          Object.keys(categoryToCode).forEach((k) => delete categoryToCode[k]);
          options.filter((o) => o.type === "ExpenseCategory").forEach((o) => { categoryToCode[o.title] = o.code || ""; });
          renderOptionsEditor();
          // If sport/type/category, also refresh requests filter dropdowns
          if (type === "Sport") {
            const sp = $("filter-sport");
            const opt = document.createElement("option"); opt.value = title; opt.textContent = title; sp.appendChild(opt);
          } else if (type === "RequestType") {
            const tp = $("filter-type");
            const opt = document.createElement("option"); opt.value = title; opt.textContent = title; tp.appendChild(opt);
          }
        });
      });
    });
  }

  async function persistOption(type, fn) {
    const indicator = document.querySelector(`[data-save="${type}"]`);
    if (indicator) { indicator.textContent = "saving…"; indicator.className = "save-indicator"; }
    try {
      await fn();
      if (indicator) { indicator.textContent = "✓ saved"; indicator.className = "save-indicator ok"; }
      setTimeout(() => { if (indicator) indicator.textContent = ""; }, 1500);
    } catch (e) {
      if (indicator) { indicator.textContent = "save failed: " + e.message.slice(0, 60); indicator.className = "save-indicator err"; }
    }
  }

  // ─── Helpers ───────────────────────────────────────────────────────────────
  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
  }
  function escapeAttr(s) { return escapeHtml(s); }
  function formatDate(s) {
    if (!s) return "";
    const d = new Date(s.length === 10 ? s + "T00:00:00" : s);
    if (isNaN(d.getTime())) return s;
    return d.toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
  }
  function formatMoney(n) {
    if (n === null || n === undefined || n === "") return "—";
    const v = Number(n);
    if (!Number.isFinite(v)) return String(n);
    return v.toLocaleString("en-US", { style: "currency", currency: "USD" });
  }
})();
