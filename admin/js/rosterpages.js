// EVAA Admin Portal — Roster Pages tab (admin-only)
// Manages which EVAA/Fusion groups get published to the public roster pages
// (markmeevaa.github.io/directory/*.html, embedded via iframe into evaasports.org) by the
// "EVAA - SportsEngine Directory Sync" Power Automate flow. The flow reads this SharePoint
// list (RosterPublishGroups) instead of a hardcoded array, and self-heals: if a group is
// deleted, the flow auto-flips that row's Status to Deleted and stamps LastError, rather
// than failing silently forever (see EMAIL-LISTS.md-style history for why this exists —
// a deleted sub-committee group's stale GUID caused days of unnoticed failures).
// Self-contained module; reuses AUTH.getToken + GRAPH helpers, same pattern as emaillists.js.

const ROSTERPAGES = (() => {
  const BASE = "https://graph.microsoft.com/v1.0";
  const SITE = "evaasports.sharepoint.com,5c93dacd-279c-41bd-a4b0-64288b689f69,3c4714c8-a098-4f4b-bdd9-ad7a69c13740";
  const LIST = "df6493dd-4b85-4011-8a85-6a3b85f9a37e"; // RosterPublishGroups
  // SharePoint hex-encoded "H1" as an internal column name (_x0048_1) -- a quirk of how it
  // got created, harmless, just has to be referenced by this internal name in every read/write.
  const H1_FIELD = "_x0048_1";
  const GITHUB_PAGES_BASE = "https://markmeevaa.github.io/directory/";

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
  function toast(msg, kind = "success") {
    const b = document.getElementById("toast-banner");
    if (!b) return;
    document.getElementById("toast-text").textContent = msg;
    b.className = "toast-banner toast-" + kind; b.classList.remove("hidden");
    setTimeout(() => b.classList.add("hidden"), 4000);
  }
  const slug = (s) => String(s || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);

  async function getRows() {
    const r = await _g(`/sites/${SITE}/lists/${LIST}/items?$expand=fields&$top=500`);
    return (r.value || []).map((i) => ({ id: i.id, f: i.fields }));
  }
  async function createRow(fields) {
    await _g(`/sites/${SITE}/lists/${LIST}/items`, { method: "POST", body: JSON.stringify({ fields }) });
  }
  async function patchRow(itemId, fields) {
    await _g(`/sites/${SITE}/lists/${LIST}/items/${itemId}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
  }
  async function deleteRow(itemId) {
    await _g(`/sites/${SITE}/lists/${LIST}/items/${itemId}`, { method: "DELETE" });
  }

  function root() { return document.getElementById("rosterpages-view"); }

  function embedSnippet(filename) {
    return `<iframe src="${GITHUB_PAGES_BASE}${filename}" style="width:100%; border:0; display:block;" scrolling="no"></iframe>\n<script src="${GITHUB_PAGES_BASE}iframe-resize-host.js"></script>`;
  }
  function copyEmbed(filename, btn) {
    const snippet = embedSnippet(filename);
    navigator.clipboard.writeText(snippet).then(
      () => { toast(`Embed code for ${filename} copied.`); },
      () => { toast("Couldn't copy automatically — select and copy the code shown.", "error"); showEmbedFallback(snippet); }
    );
  }
  function showEmbedFallback(snippet) {
    const modal = document.getElementById("confirm-modal");
    if (!modal) { alert(snippet); return; }
    document.getElementById("confirm-modal-title").textContent = "Embed code";
    document.getElementById("confirm-modal-body").innerHTML = `<textarea readonly style="width:100%;height:100px;font-family:monospace;font-size:12px">${esc(snippet)}</textarea><p class="muted">The iframe-resize-host.js script only needs to appear once per host page, even if you embed multiple roster pages there.</p>`;
    const ok = document.getElementById("confirm-modal-ok"), cancel = document.getElementById("confirm-modal-cancel");
    ok.textContent = "Close"; ok.className = "btn-primary"; modal.classList.remove("hidden");
    const cleanup = () => { modal.classList.add("hidden"); ok.removeEventListener("click", cleanup); cancel.removeEventListener("click", cleanup); };
    ok.addEventListener("click", cleanup); cancel.addEventListener("click", cleanup);
  }

  async function load() {
    root().innerHTML = `<div class="card"><p class="loading">Loading roster pages…</p></div>`;
    try {
      const rows = await getRows();
      render(rows);
    } catch (e) {
      root().innerHTML = `<div class="card error-card"><h2>Couldn't load</h2><p>${esc(e.message)}</p></div>`;
    }
  }

  function render(rows) {
    const active = rows.filter((r) => (r.f.Status || "") === "Active").sort((a, b) => String(a.f.Title || "").localeCompare(String(b.f.Title || "")));
    const orphaned = rows.filter((r) => (r.f.Status || "") === "Deleted" && r.f.LastError).sort((a, b) => String(b.f.LastValidated || "").localeCompare(String(a.f.LastValidated || "")));

    const rowHtml = (r) => {
      const f = r.f;
      return `<tr data-id="${esc(r.id)}">
        <td>${esc(f.Title)}</td>
        <td>${esc(f.Filename)}</td>
        <td>${f.IncludeInAll === "Yes" ? "Yes" : "No"}</td>
        <td class="muted">${esc((f.LastValidated || "").slice(0, 16).replace("T", " ")) || "—"}</td>
        <td style="white-space:nowrap">
          <button class="btn-link rp-embed" data-filename="${esc(f.Filename)}">Copy embed</button>
          · <button class="btn-link rp-edit" data-id="${esc(r.id)}">Edit</button>
          · <button class="btn-link rp-remove" data-id="${esc(r.id)}" style="color:#b00020">Remove</button>
        </td>
      </tr>`;
    };
    const orphanRowHtml = (r) => {
      const f = r.f;
      return `<tr data-id="${esc(r.id)}">
        <td>${esc(f.Title)}</td>
        <td>${esc(f.Filename)}</td>
        <td class="muted" style="max-width:280px">${esc(f.LastError)}</td>
        <td style="white-space:nowrap">
          <button class="btn-link rp-reactivate" data-id="${esc(r.id)}">Reactivate</button>
          · <button class="btn-link rp-delete-forever" data-id="${esc(r.id)}" style="color:#b00020">Delete row</button>
        </td>
      </tr>`;
    };

    root().innerHTML = `<div class="card">
      <div class="section-header" style="align-items:center">
        <h2>Roster Pages</h2>
        <button class="btn-secondary" id="rp-add-btn">+ Add group</button>
      </div>
      <p class="muted">Controls which groups get published to the public roster pages (embedded on evaasports.org). Sports rarely change; sub-committees do — remove a group here when it's retired instead of just deleting it from Entra, or the sync flow will catch it automatically and flag it below.</p>
      <div id="rp-add-form" class="hidden" style="margin:0 0 14px;padding:12px;background:#f4f6f9;border-radius:8px"></div>
      <table class="data-table"><thead><tr><th>Group</th><th>Filename</th><th>In "All"</th><th>Last validated</th><th class="col-actions"></th></tr></thead>
      <tbody>${active.length ? active.map(rowHtml).join("") : `<tr><td colspan="5" class="muted">No active roster pages.</td></tr>`}</tbody></table>
      ${orphaned.length ? `
        <h3 style="margin-top:22px;color:#b00020">Orphaned (${orphaned.length})</h3>
        <p class="muted">These groups no longer exist — the sync flow found them missing and stopped publishing their pages automatically. Reactivate if this was a mistake (e.g. the group was recreated with the same purpose), or delete the row to clean up.</p>
        <table class="data-table"><thead><tr><th>Group</th><th>Filename</th><th>Error</th><th class="col-actions"></th></tr></thead>
        <tbody>${orphaned.map(orphanRowHtml).join("")}</tbody></table>` : ""}
    </div>`;

    document.getElementById("rp-add-btn").addEventListener("click", openAddForm);
    root().querySelectorAll(".rp-embed").forEach((b) => b.addEventListener("click", () => copyEmbed(b.dataset.filename, b)));
    root().querySelectorAll(".rp-edit").forEach((b) => b.addEventListener("click", () => openEditForm(rows.find((r) => r.id === b.dataset.id))));
    root().querySelectorAll(".rp-remove").forEach((b) => b.addEventListener("click", () => removeRow(b.dataset.id)));
    root().querySelectorAll(".rp-reactivate").forEach((b) => b.addEventListener("click", () => reactivateRow(b.dataset.id)));
    root().querySelectorAll(".rp-delete-forever").forEach((b) => b.addEventListener("click", () => deleteForever(b.dataset.id)));
  }

  async function removeRow(itemId) {
    if (!confirm("Remove this group from the roster pages? Its page will no longer be updated (the existing published page is left as-is until manually deleted from GitHub).")) return;
    try {
      await patchRow(itemId, { Status: "Deleted", LastError: "" });
      GRAPH.logAuditEntry({ actor: (await GRAPH.getMe())?.userPrincipalName, action: "rosterpages remove group", targetGroup: itemId, result: "Success" });
      toast("Removed from roster pages.");
      load();
    } catch (e) { alert("Failed to remove: " + e.message); }
  }
  async function reactivateRow(itemId) {
    try {
      await patchRow(itemId, { Status: "Active", LastError: "" });
      toast("Reactivated — will be validated on the next sync.");
      load();
    } catch (e) { alert("Failed to reactivate: " + e.message); }
  }
  async function deleteForever(itemId) {
    if (!confirm("Permanently delete this row? This can't be undone (though you can always add the group back later).")) return;
    try {
      await deleteRow(itemId);
      toast("Row deleted.");
      load();
    } catch (e) { alert("Failed to delete: " + e.message); }
  }

  function openEditForm(row) {
    if (!row) return;
    const f = row.f;
    const form = document.getElementById("rp-add-form");
    form.innerHTML = `
      <div class="toolbar" style="flex-wrap:wrap;gap:8px">
        <input type="text" id="rp-e-filename" value="${esc(f.Filename)}" placeholder="Filename" style="min-width:180px" />
        <input type="text" id="rp-e-h1" value="${esc(f[H1_FIELD])}" placeholder="Page heading (H1)" style="min-width:260px" />
        <label style="display:flex;align-items:center;gap:6px"><input type="checkbox" id="rp-e-all" ${f.IncludeInAll === "Yes" ? "checked" : ""}/> Include in all-rosters.html</label>
        <button class="btn-primary" id="rp-e-save">Save</button>
        <button class="btn-link" id="rp-e-cancel">Cancel</button>
      </div>`;
    form.classList.remove("hidden");
    document.getElementById("rp-e-cancel").addEventListener("click", () => form.classList.add("hidden"));
    document.getElementById("rp-e-save").addEventListener("click", async () => {
      const filename = document.getElementById("rp-e-filename").value.trim();
      const h1 = document.getElementById("rp-e-h1").value.trim();
      const includeInAll = document.getElementById("rp-e-all").checked;
      if (!filename || !h1) { alert("Filename and heading are required."); return; }
      try {
        await patchRow(row.id, { Filename: filename, [H1_FIELD]: h1, IncludeInAll: includeInAll ? "Yes" : "No" });
        toast("Saved.");
        load();
      } catch (e) { alert("Failed to save: " + e.message); }
    });
  }

  async function openAddForm() {
    const form = document.getElementById("rp-add-form");
    form.innerHTML = `<p class="loading">Loading groups…</p>`;
    form.classList.remove("hidden");
    let groups;
    try { groups = await GRAPH.listAllEvaaFusionGroups(); }
    catch (e) { form.innerHTML = `<p>Couldn't load groups: ${esc(e.message)}</p>`; return; }
    groups.sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));

    const rowsHtml = groups.map((g) => `<button class="user-result rp-pick" data-id="${esc(g.id)}" data-name="${esc(g.displayName)}">${esc(g.displayName)}</button>`).join("");
    form.innerHTML = `
      <input type="search" id="rp-add-filter" placeholder="Filter groups…" style="width:100%;padding:6px;margin-bottom:8px" />
      <div id="rp-add-results" style="max-height:220px;overflow:auto">${rowsHtml}</div>
      <div id="rp-add-details" class="hidden" style="margin-top:12px;padding-top:12px;border-top:1px solid #e3e8ef">
        <div class="toolbar" style="flex-wrap:wrap;gap:8px">
          <input type="text" id="rp-a-filename" placeholder="Filename (e.g. baseball.html)" style="min-width:180px" />
          <input type="text" id="rp-a-h1" placeholder="Page heading (H1)" style="min-width:260px" />
          <label style="display:flex;align-items:center;gap:6px"><input type="checkbox" id="rp-a-all" checked /> Include in all-rosters.html</label>
          <button class="btn-primary" id="rp-a-save">Add</button>
          <button class="btn-link" id="rp-a-cancel">Cancel</button>
        </div>
      </div>`;
    document.getElementById("rp-add-filter").addEventListener("input", (e) => {
      const q = e.target.value.toLowerCase();
      form.querySelectorAll(".rp-pick").forEach((b) => { b.style.display = b.dataset.name.toLowerCase().includes(q) ? "" : "none"; });
    });
    document.getElementById("rp-a-cancel").addEventListener("click", () => form.classList.add("hidden"));
    let picked = null;
    form.querySelectorAll(".rp-pick").forEach((b) => b.addEventListener("click", () => {
      picked = { id: b.dataset.id, displayName: b.dataset.name };
      document.getElementById("rp-add-details").classList.remove("hidden");
      document.getElementById("rp-a-filename").value = slug(b.dataset.name) + ".html";
      document.getElementById("rp-a-h1").value = `EVAA ${b.dataset.name.replace(/^EVAA\s*-\s*/i, "")} — Board Members`;
    }));
    document.getElementById("rp-a-save").addEventListener("click", async () => {
      if (!picked) return;
      const filename = document.getElementById("rp-a-filename").value.trim();
      const h1 = document.getElementById("rp-a-h1").value.trim();
      const includeInAll = document.getElementById("rp-a-all").checked;
      if (!filename || !h1) { alert("Filename and heading are required."); return; }
      try {
        await createRow({
          Title: picked.displayName,
          GroupId: picked.id,
          Filename: filename,
          [H1_FIELD]: h1,
          IncludeInAll: includeInAll ? "Yes" : "No",
          Status: "Active",
        });
        GRAPH.logAuditEntry({ actor: (await GRAPH.getMe())?.userPrincipalName, action: "rosterpages add group", targetGroup: picked.id, targetName: picked.displayName, result: "Success" });
        toast(`Added ${picked.displayName} — will build on the next sync.`);
        load();
      } catch (e) { alert("Failed to add: " + e.message); }
    });
  }

  return { load };
})();
