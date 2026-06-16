// EVAA Finance — Treasurer Console behavior
// Loads FinanceRequests via Graph, renders a filterable table,
// opens a detail drawer with the status workflow, writes back updates + audit log.
// Deep-link: ?id=<itemId> opens that item's drawer directly (used by treasurer
// notification email's CTA button).

(async () => {
  const $ = (id) => document.getElementById(id);
  const show = (id) =>
    ["view-loading", "view-signin", "view-no-access", "view-console"].forEach((v) =>
      $(v).classList.toggle("hidden", v !== id)
    );

  // ─── Status workflow (next-state options per current state) ─────────────────
  const STATUS_ORDER = ["Submitted", "Under Review", "Approved", "Paid/Deposited", "Denied"];
  const STATUS_CLASS = {
    "Submitted":      "status-Submitted",
    "Under Review":   "status-UnderReview",
    "Approved":       "status-Approved",
    "Paid/Deposited": "status-PaidDeposited",
    "Denied":         "status-Denied",
  };
  // Allowed transitions from each state (excludes "no change" and reversing back to Submitted).
  const NEXT = {
    "Submitted":      ["Under Review", "Approved", "Denied"],
    "Under Review":   ["Approved", "Denied"],
    "Approved":       ["Paid/Deposited", "Denied"],
    "Paid/Deposited": [],
    "Denied":         ["Under Review"],
  };

  // ─── Boot ──────────────────────────────────────────────────────────────────
  show("view-loading");

  let account;
  try {
    account = await AUTH.init({ consoleMode: true });
  } catch (e) {
    $("loading-text").textContent = "Auth init failed: " + e.message;
    return;
  }

  if (!account) {
    const btn = $("sign-in-btn");
    btn.disabled = false;
    btn.addEventListener("click", async () => {
      try {
        const acct = await AUTH.signIn();
        if (acct) window.location.reload();
      } catch (e) {
        $("signin-error").textContent = e.message;
        $("signin-error").classList.remove("hidden");
      }
    });
    show("view-signin");
    return;
  }

  // Identity + admin/leadership gate
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

  // ─── Load list ─────────────────────────────────────────────────────────────
  $("loading-text").textContent = "Loading requests…";
  let items = [];
  try { items = await GRAPH.listFinanceRequests(); }
  catch (e) { $("loading-text").innerHTML = `<span class="alert alert-error">Failed to load: ${escapeHtml(e.message)}</span>`; return; }

  // Populate filter dropdowns from schema + observed values
  const typeSel = $("filter-type");
  ["Check Request", "Deposit Request", "Credit Card Use"].forEach((t) => {
    const o = document.createElement("option"); o.value = t; o.textContent = t; typeSel.appendChild(o);
  });
  const sportSel = $("filter-sport");
  SCHEMA.enums.SPORTS.forEach((s) => {
    const o = document.createElement("option"); o.value = s; o.textContent = s; sportSel.appendChild(o);
  });

  // Wire toolbar
  ["filter-type", "filter-status", "filter-sport", "filter-search"].forEach((id) =>
    $(id).addEventListener("input", renderRows)
  );
  $("refresh-btn").addEventListener("click", async () => {
    $("loading-text").textContent = "Refreshing…";
    show("view-loading");
    try { items = await GRAPH.listFinanceRequests(); show("view-console"); renderRows(); }
    catch (e) { $("loading-text").innerHTML = `<span class="alert alert-error">${escapeHtml(e.message)}</span>`; }
  });

  show("view-console");
  renderRows();

  // Deep-link: ?id=NN opens that row's detail drawer
  const params = new URLSearchParams(window.location.search);
  const deepLinkId = params.get("id");
  if (deepLinkId) {
    const target = items.find((it) => String(it.id) === String(deepLinkId));
    if (target) openDetail(target);
  }

  // ─── Render ────────────────────────────────────────────────────────────────
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
    // Newest first
    filtered.sort((a, b) => (b.createdDateTime || "").localeCompare(a.createdDateTime || ""));
    $("result-summary").textContent =
      `${filtered.length} of ${items.length} requests` +
      (filtered.length !== items.length ? " (filtered)" : "");

    const html = filtered.map((it) => {
      const f = it.fields || {};
      const payee = f.VendorName || f.PayerName || f.AltPayeeName || (f.VendorCardinality === "Multiple Vendors" ? "Various" : "—");
      const statusClass = STATUS_CLASS[f.Status] || "status-Submitted";
      return `<tr data-id="${escapeAttr(it.id)}">
        <td>${formatDate(f.DateRequested || it.createdDateTime)}</td>
        <td>${escapeHtml(f.SubmittedByName || "")}<br><span style="color:#888; font-size:0.9em;">${escapeHtml(f.SubmittedByEmail || "")}</span></td>
        <td>${escapeHtml(f.RequestType || "")}</td>
        <td>${escapeHtml(f.Sport || "")}</td>
        <td>${escapeHtml(payee)}</td>
        <td style="text-align:right; font-variant-numeric: tabular-nums;">${formatMoney(f.Amount)}</td>
        <td><span class="status-pill ${statusClass}">${escapeHtml(f.Status || "Submitted")}</span></td>
      </tr>`;
    }).join("") || `<tr><td colspan="7" style="text-align:center; padding:30px; color:#888;">No requests match these filters.</td></tr>`;

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
    // Remove the ?id= from the URL so a refresh doesn't reopen it
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

    $("detail-title").textContent = `${f.RequestType || "Request"} #${item.id}`;

    // Render KV rows in a sensible order
    const rows = [
      ["Status",            `<span class="status-pill ${statusClass}">${escapeHtml(status)}</span>`],
      ["Request type",      f.RequestType],
      ["Date requested",    formatDate(f.DateRequested)],
      ["Sport",             f.Sport],
      ["Program type",      f.ProgramType],
      ["Traveling subtype", f.TravelingSubtype],
      ["Season",            f.Season],
      ["Category",          f.Category],
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
      ["Receipt",           `<span id="receipt-cell">${f.ReceiptUrl ? `<a href="${escapeAttr(f.ReceiptUrl)}" target="_blank" rel="noopener">${escapeHtml(f.ReceiptFileName || "Open receipt")}</a>` : "(looking up…)"}</span>`],
      ["Source",            f.SourceSystem || "Portal"],
      ["", "<hr style='border:none; border-top:1px solid #eee; margin:8px 0;'/>"],
      ["Treasurer notes",   `<textarea id="treasurer-notes" rows="3" style="width:100%;">${escapeHtml(f.TreasurerNotes || "")}</textarea>`],
      ["Last decision",     f.DecisionBy ? `${escapeHtml(f.DecisionBy)} on ${formatDate(f.DecisionAt)}` : "(no decisions yet)"],
    ].filter(([k, v]) => k === "" || (v !== null && v !== undefined && v !== "")); // hide empty rows but keep separator rules

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

    // Wire transition buttons
    $("detail-body").querySelectorAll("button[data-transition]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const newStatus = btn.dataset.transition;
        const notes = ($("treasurer-notes") && $("treasurer-notes").value) || "";
        if (!confirm(`Move request #${item.id} to "${newStatus}"?`)) return;
        $("detail-body").querySelectorAll("button[data-transition]").forEach((b) => (b.disabled = true));
        try {
          await GRAPH.updateRequestStatus(item.id, newStatus, notes);
          // Refresh in place
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

    // Async: look up the receipt by item-id prefix in FinanceReceipts library.
    // The submission flow uploads files named "{itemId}__{filename}" but does NOT patch
    // the request's ReceiptUrl field, so we discover it here.
    if (!f.ReceiptUrl) {
      GRAPH.getReceiptForRequest(item.id).then((rcpt) => {
        const cell = document.getElementById("receipt-cell");
        if (!cell) return;
        if (rcpt && rcpt.webUrl) {
          cell.innerHTML = `<a href="${escapeAttr(rcpt.webUrl)}" target="_blank" rel="noopener">${escapeHtml(rcpt.name)}</a>`;
        } else {
          cell.textContent = "(none uploaded)";
        }
      }).catch(() => {
        const cell = document.getElementById("receipt-cell");
        if (cell) cell.textContent = "(lookup failed)";
      });
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
    // Accept date-only ("YYYY-MM-DD") or ISO datetime
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
