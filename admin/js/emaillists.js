// EVAA Admin Portal — Email Lists tab
// Director view of the auto-built SportsEngine family distribution lists.
// Reads EmailListRegistry (display) + the live DL members (Graph), and writes
// add/remove requests to EmailListOverrides (applied by the nightly sync job).
// Self-contained module; reuses AUTH.getToken + GRAPH.getMe/logAuditEntry.

const EMAILLISTS = (() => {
  const BASE = "https://graph.microsoft.com/v1.0";
  const SITE = "evaasports.sharepoint.com,5c93dacd-279c-41bd-a4b0-64288b689f69,3c4714c8-a098-4f4b-bdd9-ad7a69c13740";
  const REGISTRY = "adbb503e-16ce-4c33-915e-fe46798ad8ec";   // EmailListRegistry
  const OVERRIDES = "12d30059-845a-47df-9c32-3a8501eb4ae9";  // EmailListOverrides

  async function _g(path, options = {}) {
    const token = await AUTH.getToken();
    const url = path.startsWith("http") ? path : BASE + path;
    const resp = await fetch(url, {
      ...options,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) },
    });
    if (!resp.ok) { const b = await resp.text().catch(() => ""); throw new Error(`Graph ${resp.status}: ${b.slice(0, 300)}`); }
    if (resp.status === 204 || resp.status === 202) return null;
    const t = await resp.text(); return t ? JSON.parse(t) : null;
  }
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  function toast(msg) {
    const b = document.getElementById("toast-banner");
    if (!b) return;
    document.getElementById("toast-text").textContent = msg;
    b.className = "toast-banner toast-success"; b.classList.remove("hidden");
    setTimeout(() => b.classList.add("hidden"), 4000);
  }

  async function getRegistry() {
    const r = await _g(`/sites/${SITE}/lists/${REGISTRY}/items?$expand=fields&$top=500`);
    return (r.value || []).map((i) => i.fields).filter((f) => (f.Status || "") !== "Deleted");
  }
  async function getMembers(dlMail) {
    const g = await _g(`/groups?$filter=${encodeURIComponent("mail eq '" + dlMail + "'")}&$select=id`);
    if (!g.value || !g.value.length) return [];
    const id = g.value[0].id;
    const out = []; let next = `/groups/${id}/members?$select=id,displayName,mail&$top=200`;
    while (next) { const d = await _g(next); out.push(...(d.value || [])); next = d["@odata.nextLink"] || null; }
    return out.filter((m) => m.mail);
  }
  async function addOverride(reg, action, email, name) {
    const me = await GRAPH.getMe();
    const fields = { Title: `${action} ${email}`, RegistrationId: String(reg), ActionType: action, Email: email, RecipientName: name || "", Actor: me?.userPrincipalName || me?.mail || "" };
    await _g(`/sites/${SITE}/lists/${OVERRIDES}/items`, { method: "POST", body: JSON.stringify({ fields }) });
    GRAPH.logAuditEntry({ actor: fields.Actor, action: action === "Add" ? "emaillist add request" : "emaillist remove request", targetName: email, targetGroup: reg, result: "Success", notes: JSON.stringify({ list: reg, action }) });
  }

  function root() { return document.getElementById("emaillists-view"); }

  async function load() {
    root().innerHTML = `<div class="card"><p class="loading">Loading email lists…</p></div>`;
    try { renderList(await getRegistry()); }
    catch (e) { root().innerHTML = `<div class="card error-card"><h2>Couldn't load</h2><p>${esc(e.message)}</p></div>`; }
  }

  function renderList(regs) {
    if (!regs.length) {
      root().innerHTML = `<div class="card"><h2>Family Email Lists</h2><p class="muted">No family lists yet — they're auto-built nightly from SportsEngine registrations.</p></div>`;
      return;
    }
    root().innerHTML = `<div class="card">
      <h2>Family Email Lists</h2>
      <p class="muted">Auto-built from SportsEngine registrations (Guardian 1 + 2 emails). Click a list to view recipients and add/remove people. Only the sport's <strong>board leaders</strong> can email a list; lists auto-retire 18 months after the registration's create date.</p>
      <table class="data-table"><thead><tr><th>List address</th><th>Sport</th><th>Recipients</th><th>Expires</th></tr></thead>
      <tbody>${regs.map((f) => `<tr data-reg="${esc(f.RegistrationId)}" data-mail="${esc(f.Title)}" style="cursor:pointer"><td>${esc(f.Title)}</td><td>${esc(f.Sport)}</td><td>${esc(f.RecipientCount)}</td><td>${esc((f.ExpiresOn || "").slice(0, 10))}</td></tr>`).join("")}</tbody></table></div>`;
    root().querySelectorAll("tr[data-reg]").forEach((r) => r.addEventListener("click", () => openDetail(r.dataset.reg, r.dataset.mail)));
  }

  async function openDetail(reg, mail) {
    root().innerHTML = `<div class="card"><button class="btn-link back-link" id="el-back">← Back to lists</button><h2>${esc(mail)}</h2><p class="loading">Loading recipients…</p></div>`;
    document.getElementById("el-back").addEventListener("click", load);
    try { renderDetail(reg, mail, await getMembers(mail)); }
    catch (e) { root().querySelector(".loading").textContent = "Couldn't load recipients: " + e.message; }
  }

  function renderDetail(reg, mail, members) {
    root().innerHTML = `<div class="card">
      <button class="btn-link back-link" id="el-back">← Back to lists</button>
      <h2>${esc(mail)}</h2>
      <p class="muted"><strong>${members.length}</strong> recipients. Add or remove below — changes are <em>queued</em> and applied at the next sync. The list itself is rebuilt from registrations automatically; your removals are remembered so they won't be re-added.</p>
      <div class="toolbar">
        <input type="email" id="el-add-email" placeholder="add an email…" style="min-width:240px" />
        <input type="text" id="el-add-name" placeholder="name (optional)" />
        <button class="btn-secondary" id="el-add-btn">+ Add</button>
      </div>
      <input type="search" id="el-filter" placeholder="filter recipients…" style="margin:10px 0;width:100%;padding:6px" />
      <table class="data-table"><thead><tr><th>Email</th><th>Name</th><th class="col-actions"></th></tr></thead>
      <tbody id="el-tbody">${members.map((m) => `<tr data-email="${esc(m.mail)}"><td>${esc(m.mail)}</td><td>${esc(m.displayName || "")}</td><td><button class="btn-link el-remove" style="color:#b00020">remove</button></td></tr>`).join("")}</tbody></table></div>`;
    document.getElementById("el-back").addEventListener("click", load);

    document.getElementById("el-add-btn").addEventListener("click", async () => {
      const email = document.getElementById("el-add-email").value.trim();
      const name = document.getElementById("el-add-name").value.trim();
      if (!email || !email.includes("@")) { alert("Enter a valid email address."); return; }
      try {
        await addOverride(reg, "Add", email, name);
        toast(`Saved: ${email} will be added at the next sync`);
        // optimistic: show it immediately as pending so the change is visible right away
        const tb = document.getElementById("el-tbody");
        const tr = document.createElement("tr");
        tr.style.background = "#fffbe6";
        tr.innerHTML = `<td>${esc(email)}</td><td>${esc(name)} <span class="muted">(pending — applies at next sync)</span></td><td></td>`;
        tb.prepend(tr);
        document.getElementById("el-add-email").value = ""; document.getElementById("el-add-name").value = "";
      }
      catch (e) { alert("Failed to queue add: " + e.message); }
    });

    document.getElementById("el-tbody").addEventListener("click", async (ev) => {
      const btn = ev.target.closest(".el-remove"); if (!btn) return;
      const tr = btn.closest("tr"); const email = tr.dataset.email;
      if (!confirm(`Queue removal of ${email} from this list? It will be removed at the next sync and stay removed.`)) return;
      try { await addOverride(reg, "Exclude", email, ""); tr.style.opacity = 0.4; btn.textContent = "queued"; btn.disabled = true; toast(`Queued: remove ${email}`); }
      catch (e) { alert("Failed to queue removal: " + e.message); }
    });

    document.getElementById("el-filter").addEventListener("input", (e) => {
      const f = e.target.value.toLowerCase();
      document.querySelectorAll("#el-tbody tr").forEach((tr) => { tr.style.display = tr.textContent.toLowerCase().includes(f) ? "" : "none"; });
    });
  }

  return { load };
})();
