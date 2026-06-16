// EVAA Finance Portal — submission form behavior
// IIFE module pattern matches /admin/js/app.js.
// Renders the form from SCHEMA, evaluates conditional visibility live,
// validates required fields, and submits via GRAPH.submitFinanceRequest().

(async () => {
  const $ = (id) => document.getElementById(id);
  const show = (id) => {
    ["view-loading", "view-signin", "view-form", "view-success"].forEach((v) =>
      $(v).classList.toggle("hidden", v !== id)
    );
  };
  const setError = (msg) => {
    const el = $("form-alert");
    el.className = "alert alert-error";
    el.textContent = msg;
    el.classList.remove("hidden");
    el.scrollIntoView({ behavior: "smooth", block: "center" });
  };
  const clearError = () => $("form-alert").classList.add("hidden");

  // ─── Boot ──────────────────────────────────────────────────────────────────
  show("view-loading");

  let account;
  try {
    account = await AUTH.init({ consoleMode: false });
  } catch (err) {
    $("loading-text").textContent = "Auth init failed: " + err.message;
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

  // Signed in — populate header and prefill from /me
  let me;
  try {
    me = await GRAPH.getMe();
  } catch (e) {
    console.warn("getMe failed (non-fatal, will still let user fill manually):", e);
    me = { displayName: account.name || account.username, mail: account.username };
  }
  $("user-name").textContent = me.displayName || me.userPrincipalName || "";
  $("user-area").classList.remove("hidden");
  $("sign-out-btn").addEventListener("click", () => AUTH.signOut());

  // ─── Form state ────────────────────────────────────────────────────────────
  const values = {};
  let receiptFile = null;

  // Load dropdown options from FinanceFormOptions list before first render
  try {
    await SCHEMA.load();
  } catch (e) {
    console.warn("Schema option load failed — using defaults:", e);
  }

  // Prefill from /me where possible
  values.FirstName = me.givenName || (me.displayName || "").split(" ")[0] || "";
  values.LastName = me.surname || (me.displayName || "").split(" ").slice(1).join(" ") || "";
  values.ContactEmail = me.mail || me.userPrincipalName || "";
  values.Phone1 = (me.businessPhones && me.businessPhones[0]) || "";
  values.DateRequested = new Date().toISOString().slice(0, 10);

  // ─── Render ────────────────────────────────────────────────────────────────
  const container = $("form-sections");

  function fieldInput(field) {
    const id = `f-${field.id}`;
    const cur = values[field.id] ?? "";
    const required = SCHEMA.isRequired(field, values);
    const reqMark = required ? `<span class="req">*</span>` : "";
    const help = field.help ? `<div class="help">${escapeHtml(field.help)}</div>` : "";
    const modeled = field._modeled ? `<span class="modeled-badge" title="Inferred — no real example to validate against">modeled</span>` : "";
    // Label flips per RequestType where the schema declares labelByType.
    const labelText = SCHEMA.labelFor(field, values);

    let control;
    const reqAttr = required ? "required" : "";
    switch (field.type) {
      case "choice":
        control = `<select id="${id}" data-field="${field.id}" ${reqAttr}>
          <option value="">— select —</option>
          ${field.options.map((o) =>
            `<option value="${escapeAttr(o)}" ${cur === o ? "selected" : ""}>${escapeHtml(o)}</option>`).join("")}
        </select>`;
        break;
      case "textarea":
        control = `<textarea id="${id}" data-field="${field.id}" rows="3" ${reqAttr}>${escapeHtml(cur)}</textarea>`;
        break;
      case "currency":
        control = `<input id="${id}" data-field="${field.id}" type="number" step="0.01" min="0" inputmode="decimal" value="${escapeAttr(cur)}" ${reqAttr}>`;
        break;
      case "number":
        control = `<input id="${id}" data-field="${field.id}" type="number" value="${escapeAttr(cur)}" ${reqAttr}>`;
        break;
      case "date":
        control = `<input id="${id}" data-field="${field.id}" type="date" value="${escapeAttr(cur)}" ${reqAttr}>`;
        break;
      case "email":
        control = `<input id="${id}" data-field="${field.id}" type="email" value="${escapeAttr(cur)}" autocomplete="email" ${reqAttr}>`;
        break;
      case "tel":
        control = `<input id="${id}" data-field="${field.id}" type="tel" value="${escapeAttr(cur)}" autocomplete="tel" ${reqAttr}>`;
        break;
      case "file":
        control = `<div class="file-control">
            <input id="${id}" data-field="${field.id}" type="file" accept="image/*,application/pdf" capture="environment">
            <div class="help" id="${id}-name" style="display:none;"></div>
          </div>`;
        break;
      case "boolean":
        // Render as a checkbox with the label INSIDE the field block (not above as a separate <label>).
        return `<div class="field" data-field-wrap="${field.id}" style="display:flex;align-items:flex-start;gap:8px;">
          <input id="${id}" data-field="${field.id}" type="checkbox" ${cur ? "checked" : ""} style="margin-top:4px;">
          <label for="${id}" style="font-weight:normal;cursor:pointer;">${escapeHtml(labelText)}${reqMark}${modeled}
            ${field.help ? `<div class="help" style="margin-top:2px;">${escapeHtml(field.help)}</div>` : ""}
          </label>
        </div>`;
      default: // text
        control = `<input id="${id}" data-field="${field.id}" type="text" value="${escapeAttr(cur)}" ${reqAttr}>`;
    }
    return `<div class="field" data-field-wrap="${field.id}">
      <label for="${id}">${escapeHtml(labelText)}${reqMark}${modeled}</label>
      ${control}
      ${help}
    </div>`;
  }

  function renderForm() {
    const html = SCHEMA.sections.map((sec) => {
      const visible = SCHEMA.isVisible(sec, values);
      const modeledFlag = sec._modeled ? `<span class="modeled-badge">modeled</span>` : "";
      const help = sec.help ? `<p class="section-help">${escapeHtml(sec.help)}</p>` : "";
      return `<div class="section" data-section="${sec.id}" style="${visible ? "" : "display:none;"}">
        <h2>${escapeHtml(sec.title)}${modeledFlag}</h2>
        ${help}
        ${sec.fields.map((f) =>
          SCHEMA.isVisible(f, values)
            ? fieldInput(f)
            : `<div class="field" data-field-wrap="${f.id}" style="display:none;">${fieldInput(f)}</div>`
        ).join("")}
      </div>`;
    }).join("");
    container.innerHTML = html;
    wireInputs();
  }

  function wireInputs() {
    container.querySelectorAll("[data-field]").forEach((el) => {
      el.addEventListener("input", onFieldChange);
      el.addEventListener("change", onFieldChange);
    });
  }

  // Fields whose value affects other fields' visibility or required-ness.
  // Re-render is ONLY triggered for these — typing in the Amount/Notes/Phone/etc.
  // no longer reflows the whole form, which fixes the cursor-jump bug.
  const GATING_FIELDS = SCHEMA.gatingFieldIds();

  function onFieldChange(e) {
    const fid = e.target.dataset.field;
    if (!fid) return;
    if (e.target.type === "file") {
      const f = e.target.files && e.target.files[0];
      receiptFile = f || null;
      const nameEl = document.getElementById(e.target.id + "-name");
      if (f && nameEl) {
        nameEl.style.display = "block";
        nameEl.textContent = `Selected: ${f.name} (${formatBytes(f.size)})`;
        // For oversized images, preview the compressed size so the user sees it'll work.
        if (f.type && f.type.startsWith("image/") && f.size > 2 * 1024 * 1024) {
          nameEl.textContent += " — compressing for upload…";
          GRAPH.readFileAsBase64(f).then((res) => {
            if (res) {
              nameEl.textContent =
                `Selected: ${res.fileName} (${formatBytes(res.sentBytes)}, resized from ${formatBytes(res.originalBytes)})`;
            }
          }).catch(() => {/* leave the original message */});
        }
      } else if (nameEl) {
        nameEl.style.display = "none";
        nameEl.textContent = "";
      }
      return;
    }
    // Checkbox -> store as bool; text inputs -> store the value string
    values[fid] = e.target.type === "checkbox" ? e.target.checked : e.target.value;
    // Only re-render when the changed field gates downstream visibility or required-ness.
    // For all other fields (Amount, Notes, addresses, phone, etc.) we leave the DOM alone,
    // which preserves cursor position and avoids the typing-reverses-input bug.
    if (GATING_FIELDS.has(fid)) {
      rerenderPreservingFocus(fid);
    }
  }

  function rerenderPreservingFocus(changedFieldId) {
    const active = document.activeElement;
    const activeFieldId = active && active.dataset && active.dataset.field;
    const selStart = active && "selectionStart" in active ? active.selectionStart : null;
    renderForm();
    if (activeFieldId) {
      const el = container.querySelector(`[data-field="${activeFieldId}"]`);
      if (el) {
        el.focus();
        if (selStart != null && "setSelectionRange" in el) {
          try { el.setSelectionRange(selStart, selStart); } catch {}
        }
      }
    }
  }

  // ─── Validate + submit ─────────────────────────────────────────────────────
  function validate() {
    const visible = SCHEMA.visibleFields(values);
    const missing = visible.filter((f) => {
      if (!SCHEMA.isRequired(f, values)) return false;
      const v = values[f.id];
      if (f.type === "boolean") return false; // booleans default to false, never "missing"
      return v === undefined || v === null || String(v).trim() === "";
    });
    if (missing.length) {
      return `Please fill in: ${missing.map((f) => f.label).join(", ")}`;
    }
    // Spot-check amount
    if (values.Amount !== undefined && values.Amount !== "") {
      const n = Number(values.Amount);
      if (!Number.isFinite(n) || n <= 0) return "Amount must be a positive number.";
    }
    return null;
  }

  $("finance-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    clearError();
    const err = validate();
    if (err) return setError(err);

    const submitBtn = $("submit-btn");
    submitBtn.disabled = true;
    submitBtn.textContent = "Submitting…";

    try {
      // Build a payload containing ONLY currently-visible fields (so we don't
      // smuggle stale data from sections the user toggled away).
      const visible = SCHEMA.visibleFields(values);
      const payload = {};
      visible.forEach((f) => { payload[f.id] = values[f.id] ?? null; });

      const result = await GRAPH.submitFinanceRequest(payload, receiptFile);
      const refId = result && result.id ? `Reference #${result.id}` : "";
      $("success-ref").textContent = refId;
      show("view-success");
    } catch (e) {
      setError("Submit failed: " + e.message);
      submitBtn.disabled = false;
      submitBtn.textContent = "Submit request";
    }
  });

  $("reset-btn").addEventListener("click", () => {
    if (!confirm("Clear the form?")) return;
    Object.keys(values).forEach((k) => delete values[k]);
    receiptFile = null;
    values.DateRequested = new Date().toISOString().slice(0, 10);
    values.FirstName = me.givenName || "";
    values.LastName = me.surname || "";
    values.ContactEmail = me.mail || me.userPrincipalName || "";
    renderForm();
    clearError();
  });

  $("another-btn").addEventListener("click", () => window.location.reload());

  // ─── Helpers ───────────────────────────────────────────────────────────────
  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
  }
  function escapeAttr(s) { return escapeHtml(s); }
  function formatBytes(n) {
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / (1024 * 1024)).toFixed(1)} MB`;
  }

  renderForm();
  show("view-form");
})();
