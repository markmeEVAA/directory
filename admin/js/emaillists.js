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
  const TEAMCAT = "6ee018fd-c2b6-42a4-add0-3910452b690d";    // TeamCatalog (SE Programs teams + counts)
  const REGCAT  = "3e3ed978-37e0-4d72-b1e1-bcced6892ce2";    // RegistrationCatalog (available registrations + counts)
  let _lastRegs = [];   // scoped lists from the last load(), reused by the composer
  // Power Automate flow SAS URL that fires the GitHub sync on edits (true ~5-min apply).
  // Empty string = edits apply at the nightly sync instead. Set this once the flow exists.
  const TRIGGER_URL = "";
  async function triggerSync() {
    if (!TRIGGER_URL) return;
    try { await fetch(TRIGGER_URL, { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" }); } catch (e) { /* non-fatal */ }
  }

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
    return (r.value || []).map((i) => ({ id: i.id, f: i.fields })).filter((x) => (x.f.Status || "") !== "Deleted");
  }
  async function patchRegistry(itemId, fields) {
    await _g(`/sites/${SITE}/lists/${REGISTRY}/items/${itemId}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
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
    triggerSync();
  }

  function root() { return document.getElementById("emaillists-view"); }

  // --- SportsEngine-sourced list creation (opt-in): per-registration, per-team, coaches ---
  // Promise<boolean> confirm dialog reusing the shared #confirm-modal DOM (app.js owns the markup).
  function confirmModal({ title = "Email Admin Portal asks…", body, okLabel = "OK", okClass = "btn-primary" }) {
    return new Promise((resolve) => {
      const modal = document.getElementById("confirm-modal");
      if (!modal) { resolve(window.confirm(String(body).replace(/<[^>]+>/g, ""))); return; }
      document.getElementById("confirm-modal-title").textContent = title;
      document.getElementById("confirm-modal-body").innerHTML = body;
      const ok = document.getElementById("confirm-modal-ok"), cancel = document.getElementById("confirm-modal-cancel");
      ok.textContent = okLabel; ok.className = okClass; modal.classList.remove("hidden");
      const cleanup = (r) => { modal.classList.add("hidden"); ok.removeEventListener("click", onOk); cancel.removeEventListener("click", onCancel); document.removeEventListener("keydown", onKey); resolve(r); };
      const onOk = () => cleanup(true), onCancel = () => cleanup(false), onKey = (e) => { if (e.key === "Escape") cleanup(false); else if (e.key === "Enter") cleanup(true); };
      ok.addEventListener("click", onOk); cancel.addEventListener("click", onCancel); document.addEventListener("keydown", onKey);
    });
  }

  async function getCatalogRows(listId) {
    const r = await _g(`/sites/${SITE}/lists/${listId}/items?$expand=fields&$top=999`);
    return (r.value || []).map((i) => i.fields);
  }
  // Catalogs scoped to the director's sport (admins see all), matched on BoardGroup like the registry.
  async function getScopedCatalogs() {
    const isAdmin = await GRAPH.isPortalAdmin();
    let owned = null;
    if (!isAdmin) {
      const me = await GRAPH.getMe();
      const gs = await GRAPH.getUserOwnedGroups(me.id);
      owned = new Set(gs.map((g) => (g.mail || "").toLowerCase()).filter(Boolean));
    }
    const inScope = (f) => isAdmin || owned.has((f.BoardGroup || "").toLowerCase());
    // Exclude items that already have a (non-deleted) list so directors can't create duplicates.
    const existing = new Set((await getRegistry()).map((x) => String(x.f.RegistrationId || "")));
    const teams = (await getCatalogRows(TEAMCAT)).filter(inScope).filter((t) => !existing.has("team-fam-" + t.TeamId));
    const regs = (await getCatalogRows(REGCAT)).filter(inScope).filter((r) => !existing.has(String(r.RegistrationId)));
    return { teams, regs, existing };
  }

  // Naming mirrors Sync-EmailLists.ps1 Get-Naming for registration lists; team/coach list
  // addresses are authoritative here (the sync builds team rows at exactly this Title).
  function regFamAddress(sport, regName) {
    const s = String(sport || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const year = (String(regName).match(/\b(20\d\d)\b/) || [])[1];
    const season = (String(regName).match(/spring|summer|fall|winter/i) || [""])[0].toLowerCase();
    const domain = /softball|fusion/i.test(regName) ? "avfusion.org" : "evaasports.org";
    return (s && year && season) ? `${s}-${year}-${season}-families@${domain}` : null;
  }
  const teamFamAddress = (sport, teamName) => slug(`${sport}-${teamName}`) + "-fam@evaasports.org";
  const coachesAddress = (sport) => slug(sport) + "-coaches@evaasports.org";

  async function createListRow(fields) {
    await _g(`/sites/${SITE}/lists/${REGISTRY}/items`, { method: "POST", body: JSON.stringify({ fields }) });
  }

  async function openSeCreate() {
    root().innerHTML = `<div class="card"><button class="btn-link back-link" id="sc-back">← Back to lists</button><p class="loading">Loading SportsEngine teams & registrations…</p></div>`;
    document.getElementById("sc-back").addEventListener("click", load);
    try { renderSeCreate(await getScopedCatalogs()); }
    catch (e) { root().querySelector(".loading").textContent = "Couldn't load: " + e.message; }
  }

  function renderSeCreate(cat) {
    root().innerHTML = `<div class="card">
      <button class="btn-link back-link" id="sc-back">← Back to lists</button>
      <h2>Create lists from SportsEngine</h2>
      <p class="muted">Pick a source, choose what you want, and check the counts against your registration report. New lists are hidden from the address book and send-locked to your board; they build at the next sync.</p>
      <div class="toolbar" style="gap:8px;margin-bottom:12px">
        <button class="btn-secondary" id="sc-src-reg">By registration</button>
        <button class="btn-secondary" id="sc-src-team">By team &amp; season</button>
      </div>
      <div id="sc-body"></div>
      <div style="margin-top:14px"><button class="btn-primary" id="sc-create">Create selected…</button></div>
    </div>`;
    document.getElementById("sc-back").addEventListener("click", load);
    const body = document.getElementById("sc-body");
    let mode = "reg";

    const renderReg = () => {
      if (!cat.regs.length) { body.innerHTML = `<p class="muted">No registrations available for your sport.</p>`; return; }
      const rows = cat.regs.slice().sort((a, b) => String(b.RegCreated || "").localeCompare(String(a.RegCreated || "")));
      body.innerHTML = `<p class="muted">One family list per registration (all guardian emails). "Registrants" should match your registration report. (Team lists build automatically — no action needed here.)</p>
        <table class="data-table"><thead><tr><th style="width:28px"><input type="checkbox" id="sc-all" title="Select all"></th><th>Registration</th><th>Registrants</th></tr></thead><tbody>${rows.map((r) => `
          <tr><td style="width:28px"><input type="checkbox" class="sc-reg" data-id="${esc(r.RegistrationId)}" data-name="${esc(r.RegName)}" data-sport="${esc(r.Sport)}" data-board="${esc(r.BoardGroup)}" data-count="${esc(r.RegistrantCount)}"></td>
          <td>${esc(r.RegName)}</td><td>${esc(r.RegistrantCount)}</td></tr>`).join("")}</tbody></table>`;
      const all = document.getElementById("sc-all");
      if (all) all.addEventListener("change", () => document.querySelectorAll(".sc-reg").forEach((cb) => { cb.checked = all.checked; }));
    };
    const renderTeam = () => {
      if (!cat.teams.length) { body.innerHTML = `<p class="muted">Team lists build <strong>automatically</strong> for sports that use SportsEngine teams/seasons (and retire after the season ends) — nothing to create here.</p>`; return; }
      const t0 = cat.teams[0];
      const byDiv = {};
      cat.teams.forEach((t) => { (byDiv[t.Division || "(teams)"] = byDiv[t.Division || "(teams)"] || []).push(t); });
      const coachExists = cat.existing && cat.existing.has("team-coach-" + t0.ProgramId);
      let html = `<p class="muted">A family list per team, and/or one coaches list for the whole program (${esc(t0.ProgramName)}).</p>`
        + (coachExists
            ? `<p class="muted" style="margin:6px 0 12px">✓ Coaches list already created.</p>`
            : `<label style="display:block;margin:6px 0 12px"><input type="checkbox" id="sc-coaches" data-sport="${esc(t0.Sport)}" data-board="${esc(t0.BoardGroup)}" data-prog="${esc(t0.ProgramId)}"> <strong>Coaches list</strong> — all ${esc(t0.Sport)} coaches</label>`);
      html += `<label style="display:block;margin:6px 0 10px"><input type="checkbox" id="sc-all-team"> <strong>Select all teams</strong></label>`;
      Object.keys(byDiv).sort().forEach((div) => {
        html += `<div style="margin:10px 0 4px;font-weight:600">${esc(div)}</div><table class="data-table"><tbody>${byDiv[div].map((t) => `
          <tr><td style="width:28px"><input type="checkbox" class="sc-team" data-id="${esc(t.TeamId)}" data-name="${esc(t.TeamName)}" data-sport="${esc(t.Sport)}" data-board="${esc(t.BoardGroup)}" data-count="${esc(t.GuardianEmailCount)}"></td>
          <td>${esc(t.TeamName)}</td><td class="muted">${esc(t.PlayerCount)} players · ${esc(t.GuardianEmailCount)} families · ${esc(t.CoachCount)} coaches</td></tr>`).join("")}</tbody></table>`;
      });
      body.innerHTML = html;
      const allT = document.getElementById("sc-all-team");
      if (allT) allT.addEventListener("change", () => document.querySelectorAll(".sc-team").forEach((cb) => { cb.checked = allT.checked; }));
    };
    const setMode = (m) => {
      mode = m;
      document.getElementById("sc-src-reg").className = m === "reg" ? "btn-primary" : "btn-secondary";
      document.getElementById("sc-src-team").className = m === "team" ? "btn-primary" : "btn-secondary";
      (m === "reg" ? renderReg : renderTeam)();
    };
    document.getElementById("sc-src-reg").addEventListener("click", () => setMode("reg"));
    document.getElementById("sc-src-team").addEventListener("click", () => setMode("team"));
    setMode("reg");

    document.getElementById("sc-create").addEventListener("click", async () => {
      const picks = [];
      if (mode === "reg") {
        document.querySelectorAll(".sc-reg:checked").forEach((cb) => {
          const d = cb.dataset; const addr = regFamAddress(d.sport, d.name);
          picks.push({ title: addr || (slug(`${d.sport}-${d.name}`) + "-families@evaasports.org"), regId: d.id, sport: d.sport, board: d.board, source: "Registration", label: d.name, count: +d.count || 0 });
        });
      } else {
        const coach = document.getElementById("sc-coaches");
        if (coach && coach.checked) picks.push({ title: coachesAddress(coach.dataset.sport), regId: "team-coach-" + coach.dataset.prog, sport: coach.dataset.sport, board: coach.dataset.board, source: "Team", label: `All ${coach.dataset.sport} coaches`, count: null });
        document.querySelectorAll(".sc-team:checked").forEach((cb) => {
          const d = cb.dataset;
          picks.push({ title: teamFamAddress(d.sport, d.name), regId: "team-fam-" + d.id, sport: d.sport, board: d.board, source: "Team", label: `${d.name} (families)`, count: +d.count || 0 });
        });
      }
      if (!picks.length) { alert("Select at least one item to create."); return; }
      const total = picks.reduce((s, p) => s + (p.count || 0), 0);
      const lines = picks.map((p) => `<li>${esc(p.label)} → <code>${esc(p.title)}</code>${p.count != null ? ` <span class="muted">(~${p.count})</span>` : ""}</li>`).join("");
      const ok = await confirmModal({
        title: `Create ${picks.length} list${picks.length > 1 ? "s" : ""}?`,
        okLabel: "Create",
        body: `<p>These will be created now and <strong>built at the next nightly sync</strong> (ready by tomorrow morning). Larger lists take longer to build.</p>
               <ul style="max-height:240px;overflow:auto">${lines}</ul>
               ${total ? `<p><strong>~${total}</strong> recipients total — please confirm this lines up with your registration report.</p>` : ""}
               <p class="muted">Each list is hidden from the global address book; only your board can send to it.</p>`
      });
      if (!ok) return;
      try {
        for (const p of picks) await createListRow({ Title: p.title, RegistrationId: p.regId, Sport: p.sport, BoardGroup: p.board, Source: p.source, Status: "Active", RecipientCount: 0 });
        triggerSync();
        toast(`Queued ${picks.length} list${picks.length > 1 ? "s" : ""} — building at the next sync.`);
        load();
      } catch (e) { alert("Failed to create: " + e.message); }
    });
  }

  async function load() {
    root().innerHTML = `<div class="card"><p class="loading">Loading email lists…</p></div>`;
    try {
      let regs = await getRegistry();
      const isAdmin = await GRAPH.isPortalAdmin();
      if (!isAdmin) {
        // board leader (owner): show only lists for groups they own
        const me = await GRAPH.getMe();
        const owned = await GRAPH.getUserOwnedGroups(me.id);
        const ownedMails = new Set(owned.map((g) => (g.mail || "").toLowerCase()).filter(Boolean));
        regs = regs.filter((x) => ownedMails.has((x.f.BoardGroup || "").toLowerCase()));
      }
      _lastRegs = regs;
      renderList(regs);
    }
    catch (e) { root().innerHTML = `<div class="card error-card"><h2>Couldn't load</h2><p>${esc(e.message)}</p></div>`; }
  }

  const slug = (s) => String(s || "").toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 40);
  async function eligibleGroups() {
    const isAdmin = await GRAPH.isPortalAdmin();
    const me = await GRAPH.getMe();
    const groups = isAdmin ? await GRAPH.listManagedGroups() : await GRAPH.getUserOwnedGroups(me.id);
    return (groups || []).filter((g) => g.mail).map((g) => ({ name: g.displayName, mail: g.mail })).sort((a, b) => a.name.localeCompare(b.name));
  }
  async function createCustomList(name, boardMail) {
    const a = slug(name);
    if (!a) throw new Error("Enter a valid list name.");
    if (!boardMail) throw new Error("Pick who's allowed to send to it.");
    const smtp = a + "@evaasports.org";
    const fields = { Title: smtp, RegistrationId: "manual-" + a, Sport: "(custom)", BoardGroup: boardMail, Source: "Manual", Status: "Active", RecipientCount: 0 };
    await _g(`/sites/${SITE}/lists/${REGISTRY}/items`, { method: "POST", body: JSON.stringify({ fields }) });
    return smtp;
  }

  function distinctVals(regs, sel) {
    return [...new Set(regs.map((x) => x.f[sel]).filter((v) => v && v !== "All"))].sort();
  }
  // Shared Type/Gender/Age/text filter bar — used by both the list screen and the send composer.
  function filterControls(regs, prefix) {
    const optList = (label, vals) => `<option value="">${label}</option>` + vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
    const cats = distinctVals(regs, "Category"), genders = distinctVals(regs, "Gender"), ages = distinctVals(regs, "AgeGroup");
    if (!(cats.length || genders.length || ages.length)) return "";
    return `<div class="toolbar" style="gap:8px;margin:4px 0 10px;flex-wrap:wrap">
        <select id="${prefix}-cat" style="padding:5px">${optList("All types", cats)}</select>
        <select id="${prefix}-gender" style="padding:5px">${optList("All genders", genders)}</select>
        <select id="${prefix}-age" style="padding:5px">${optList("All ages", ages)}</select>
        <input type="search" id="${prefix}-text" placeholder="filter address…" style="min-width:180px;padding:5px" />
      </div>`;
  }
  // Wire a filter bar built by filterControls; calls onApply(cat, gender, age, text) on every change.
  // Returns the apply fn (or null if no bar was rendered) so callers can run it once up front.
  function wireFilters(prefix, onApply) {
    const el = (s) => document.getElementById(s);
    if (!el(prefix + "-cat")) return null;
    const run = () => onApply(el(prefix + "-cat").value, el(prefix + "-gender").value, el(prefix + "-age").value, (el(prefix + "-text").value || "").toLowerCase());
    [prefix + "-cat", prefix + "-gender", prefix + "-age"].forEach((id) => el(id).addEventListener("change", run));
    el(prefix + "-text").addEventListener("input", run);
    return run;
  }
  // Which bucket a list belongs to (drives the grouped layout in the composer).
  function listKind(f) {
    const rid = String(f.RegistrationId || "");
    if (rid.startsWith("team-coach-")) return "coaches";
    if (rid.startsWith("team-fam-")) return "teams";
    if (rid.startsWith("manual-")) return "custom";
    return "reg";
  }

  function renderList(regs) {
    const createUI = `
      <div class="section-header" style="align-items:center">
        <h2>Email Lists</h2>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          <button class="btn-secondary" id="el-send-btn">✉ Send to lists</button>
          <button class="btn-secondary" id="el-se-btn">+ Create from SportsEngine</button>
          <button class="btn-secondary" id="el-create-btn">+ Create custom list</button>
        </div>
      </div>
      <div id="el-create-form" class="hidden" style="margin:0 0 14px;padding:12px;background:#f4f6f9;border-radius:8px">
        <div class="toolbar" style="flex-wrap:wrap;gap:8px">
          <input type="text" id="el-new-name" placeholder="List name (e.g. Tennis Boosters)" style="min-width:220px" />
          <select id="el-new-group" style="min-width:220px;padding:6px"><option value="">Loading groups…</option></select>
          <button class="btn-primary" id="el-new-create">Create</button>
          <button class="btn-link" id="el-new-cancel">Cancel</button>
        </div>
        <p class="muted" style="margin:6px 0 0">A custom list isn't tied to any registration — you manage every recipient by hand. The group you pick is who's allowed to send to it. Add people after it's created.</p>
      </div>`;
    const filterBar = filterControls(regs, "el-f");
    const body = regs.length
      ? `<p class="muted">Auto lists are built from registrations (Guardian 1 + 2) and SportsEngine teams; custom lists you manage by hand. Click a list to view recipients, add/remove people, shorten expiry, or delete it.</p>
         ${filterBar}
         <table class="data-table"><thead><tr><th>List address</th><th>Type</th><th>Division</th><th>Recipients</th><th>Expires</th></tr></thead>
         <tbody>${regs.map((x) => { const f = x.f;
            const type = (String(f.RegistrationId || "").startsWith("manual-")) ? "custom" : (f.Category || f.Sport || "");
            const divLabel = [f.Gender, f.AgeGroup].filter((v) => v && v !== "All").join(" ");
            return `<tr data-reg="${esc(f.RegistrationId)}" data-mail="${esc(f.Title)}" data-itemid="${esc(x.id)}" data-expires="${esc((f.ExpiresOn || "").slice(0, 10))}" data-cat="${esc(f.Category || "")}" data-gender="${esc(f.Gender || "")}" data-age="${esc(f.AgeGroup || "")}" data-text="${esc(String(f.Title || "").toLowerCase())}" style="cursor:pointer"><td>${esc(f.Title)}</td><td>${esc(type)}</td><td>${esc(divLabel)}</td><td>${esc(f.RecipientCount)}</td><td>${esc((f.ExpiresOn || "").slice(0, 10) || "—")}</td></tr>`; }).join("")}</tbody></table>`
      : `<p class="muted">No lists yet. Auto lists are built nightly from SportsEngine registrations + teams — or create a custom one above.</p>`;
    root().innerHTML = `<div class="card">${createUI}${body}</div>`;
    root().querySelectorAll("tr[data-reg]").forEach((r) => r.addEventListener("click", () => openDetail(r.dataset.reg, r.dataset.mail, r.dataset.itemid, r.dataset.expires)));

    // Type / Gender / Age / text filters (boys vs girls, age groups, family vs coaches)
    wireFilters("el-f", (c, g, a, t) => {
      root().querySelectorAll("tr[data-reg]").forEach((r) => {
        const ok = (!c || r.dataset.cat === c) && (!g || r.dataset.gender === g) && (!a || r.dataset.age === a) && (!t || (r.dataset.text || "").includes(t));
        r.style.display = ok ? "" : "none";
      });
    });

    document.getElementById("el-send-btn").addEventListener("click", openCompose);
    document.getElementById("el-se-btn").addEventListener("click", openSeCreate);
    document.getElementById("el-create-btn").addEventListener("click", async () => {
      const form = document.getElementById("el-create-form");
      form.classList.toggle("hidden");
      if (!form.classList.contains("hidden")) {
        const sel = document.getElementById("el-new-group");
        try { const gs = await eligibleGroups(); sel.innerHTML = `<option value="">Who can send to it…</option>` + gs.map((g) => `<option value="${esc(g.mail)}">${esc(g.name)}</option>`).join(""); }
        catch (e) { sel.innerHTML = `<option value="">(couldn't load groups)</option>`; }
      }
    });
    document.getElementById("el-new-cancel").addEventListener("click", () => document.getElementById("el-create-form").classList.add("hidden"));
    document.getElementById("el-new-create").addEventListener("click", async () => {
      const name = document.getElementById("el-new-name").value;
      const group = document.getElementById("el-new-group").value;
      try { const smtp = await createCustomList(name, group); triggerSync(); toast(`Created ${smtp} — add recipients; it builds at the next sync.`); load(); }
      catch (e) { alert(e.message); }
    });
  }

  async function openDetail(reg, mail, itemId, expires) {
    root().innerHTML = `<div class="card"><button class="btn-link back-link" id="el-back">← Back to lists</button><h2>${esc(mail)}</h2><p class="loading">Loading recipients…</p></div>`;
    document.getElementById("el-back").addEventListener("click", load);
    try { renderDetail(reg, mail, await getMembers(mail), itemId, expires); }
    catch (e) { root().querySelector(".loading").textContent = "Couldn't load recipients: " + e.message; }
  }

  function renderDetail(reg, mail, members, itemId, expires) {
    root().innerHTML = `<div class="card">
      <button class="btn-link back-link" id="el-back">← Back to lists</button>
      <h2>${esc(mail)}</h2>
      <div class="section-header" style="align-items:center">
        <div class="muted">Auto-retires <strong id="el-expires">${esc(expires || "—")}</strong> · <button class="btn-link" id="el-edit-expiry">retire earlier…</button></div>
        <button class="btn-danger" id="el-delete-list">Delete entire list</button>
      </div>
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

    // --- registry control plane: shorten expiry / delete whole list ---
    document.getElementById("el-edit-expiry").addEventListener("click", async () => {
      if (!itemId) { alert("This list isn't in the registry yet — available after the first sync."); return; }
      const v = prompt("Retire this list on (YYYY-MM-DD) — must be earlier than the current date:", expires || "");
      if (!v) return;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) { alert("Use format YYYY-MM-DD."); return; }
      if (expires && v >= expires) { alert("New date must be earlier than the current expiry (" + expires + ")."); return; }
      try { await patchRegistry(itemId, { ExpiresOn: v + "T00:00:00Z" }); document.getElementById("el-expires").textContent = v; expires = v; toast("Updated — list will retire on " + v); }
      catch (e) { alert("Failed to update expiry: " + e.message); }
    });
    document.getElementById("el-delete-list").addEventListener("click", async () => {
      if (!itemId) { alert("This list isn't in the registry yet — available after the first sync."); return; }
      if (!confirm(`Delete the ENTIRE "${mail}" list at the next sync? Every recipient is removed and the address is deleted.`)) return;
      try {
        await patchRegistry(itemId, { Status: "Deleted" });
        triggerSync();
        GRAPH.logAuditEntry({ actor: (await GRAPH.getMe())?.userPrincipalName, action: "emaillist delete list", targetGroup: mail, result: "Success", notes: "queued list deletion" });
        toast("List queued for deletion at next sync.");
        load();
      } catch (e) { alert("Failed to queue deletion: " + e.message); }
    });
  }

  // --- compose & send to one or more of the director's lists (BCC; sends as the director) ---
  const EXO_DAILY = 10000;   // EXO recipient-rate limit per mailbox per day (approx)
  function wrapBranded(bodyText) {
    const safe = esc(bodyText).replace(/\n/g, "<br>");
    return `<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:640px;line-height:1.6">
      <div style="background:#1B4F8C;color:#fff;padding:12px 18px;font-weight:600;border-radius:6px 6px 0 0">Eastview Athletic Association</div>
      <div style="border:1px solid #e3e8ef;border-top:none;padding:18px;border-radius:0 0 6px 6px">${safe}</div>
    </div>`;
  }

  function openCompose() {
    const lists = _lastRegs.filter((x) => (x.f.Status || "") !== "Deleted" && x.f.Title);
    if (!lists.length) { alert("You have no lists to send to yet."); return; }

    // Bucket the lists so a sport with dozens of team lists stays scannable.
    const KIND_ORDER = ["reg", "teams", "coaches", "custom"];
    const KIND_LABEL = { reg: "Registration family lists", teams: "Team family lists", coaches: "Coaches lists", custom: "Custom lists" };
    const groups = {};
    lists.forEach((x) => { const k = listKind(x.f); (groups[k] = groups[k] || []).push(x); });
    const groupHtml = KIND_ORDER.filter((k) => groups[k] && groups[k].length).map((k) => {
      const items = groups[k].slice().sort((a, b) => String(a.f.Title).localeCompare(String(b.f.Title)));
      const rows = items.map((x) => {
        const f = x.f;
        const div = [f.Gender, f.AgeGroup].filter((v) => v && v !== "All").join(" ");
        return `<label class="cm-row" style="display:block;margin:3px 0" data-cat="${esc(f.Category || "")}" data-gender="${esc(f.Gender || "")}" data-age="${esc(f.AgeGroup || "")}" data-text="${esc(String(f.Title || "").toLowerCase())}" data-group="${k}">
            <input type="checkbox" class="cm-list" data-mail="${esc(f.Title)}" data-count="${esc(f.RecipientCount || 0)}" data-group="${k}"> ${esc(f.Title)}${div ? ` <span class="muted">· ${esc(div)}</span>` : ""} <span class="muted">· ${esc(f.RecipientCount || 0)} recipients</span>
          </label>`;
      }).join("");
      return `<div class="cm-group" data-group="${k}" style="margin:10px 0">
          <label class="cm-group-head" style="display:block;font-weight:600;border-bottom:1px solid #e3e8ef;padding-bottom:3px"><input type="checkbox" class="cm-group-all" data-group="${k}"> ${esc(KIND_LABEL[k])} <span class="muted">(${items.length})</span></label>
          <div class="cm-group-items" style="margin:4px 0 0 6px">${rows}</div>
        </div>`;
    }).join("");

    root().innerHTML = `<div class="card">
      <button class="btn-link back-link" id="cm-back">← Back to lists</button>
      <h2>Send to lists</h2>
      <p class="muted">Pick one or more of your lists and write your message. Recipients go in <strong>BCC</strong> — they can't see each other or reply-all. The email is sent from your address and saved to your Sent Items.</p>
      ${filterControls(lists, "cm-f")}
      <label style="display:block;margin:8px 0 4px;font-weight:600"><input type="checkbox" id="cm-all"> Select all shown</label>
      <div id="cm-groups" style="margin:4px 0 10px">${groupHtml}</div>
      <div class="muted" id="cm-total" style="margin:6px 0">Selected: 0 lists · ~0 recipients</div>
      <div class="toolbar" style="flex-direction:column;align-items:stretch;gap:8px;max-width:640px">
        <input type="text" id="cm-subject" placeholder="Subject" style="padding:8px" />
        <textarea id="cm-body" placeholder="Write your message…" rows="10" style="padding:8px;font-family:inherit"></textarea>
      </div>
      <div style="margin-top:12px"><button class="btn-primary" id="cm-send">Send…</button></div>
    </div>`;
    document.getElementById("cm-back").addEventListener("click", load);

    const visibleRows = () => [...document.querySelectorAll(".cm-row")].filter((r) => r.style.display !== "none");
    const updateTotal = () => {
      const sel = [...document.querySelectorAll(".cm-list:checked")];
      const total = sel.reduce((s, cb) => s + (+cb.dataset.count || 0), 0);
      document.getElementById("cm-total").textContent = `Selected: ${sel.length} list${sel.length === 1 ? "" : "s"} · ~${total} recipients`;
      // reflect group / select-all checkbox state from their currently-visible children
      const vis = visibleRows();
      document.querySelectorAll(".cm-group-all").forEach((ga) => {
        const kids = vis.filter((r) => r.dataset.group === ga.dataset.group).map((r) => r.querySelector(".cm-list"));
        ga.checked = kids.length > 0 && kids.every((cb) => cb.checked);
        ga.indeterminate = !ga.checked && kids.some((cb) => cb.checked);
      });
      const allCb = document.getElementById("cm-all");
      const visCbs = vis.map((r) => r.querySelector(".cm-list"));
      allCb.checked = visCbs.length > 0 && visCbs.every((cb) => cb.checked);
      allCb.indeterminate = !allCb.checked && visCbs.some((cb) => cb.checked);
    };

    document.querySelectorAll(".cm-list").forEach((cb) => cb.addEventListener("change", updateTotal));
    // group "select all" toggles only the currently-visible rows in that group
    document.querySelectorAll(".cm-group-all").forEach((ga) => ga.addEventListener("change", () => {
      visibleRows().filter((r) => r.dataset.group === ga.dataset.group).forEach((r) => { r.querySelector(".cm-list").checked = ga.checked; });
      updateTotal();
    }));
    document.getElementById("cm-all").addEventListener("change", (e) => {
      visibleRows().forEach((r) => { r.querySelector(".cm-list").checked = e.target.checked; });
      updateTotal();
    });

    // Type / Gender / Age / text filters hide rows; empty groups collapse their header too.
    wireFilters("cm-f", (c, g, a, t) => {
      document.querySelectorAll(".cm-row").forEach((r) => {
        const ok = (!c || r.dataset.cat === c) && (!g || r.dataset.gender === g) && (!a || r.dataset.age === a) && (!t || (r.dataset.text || "").includes(t));
        r.style.display = ok ? "block" : "none";   // .cm-row is a <label> (default inline) — must restore block, not ""
      });
      document.querySelectorAll(".cm-group").forEach((grp) => {
        const any = [...grp.querySelectorAll(".cm-row")].some((r) => r.style.display !== "none");
        grp.style.display = any ? "" : "none";
      });
      updateTotal();
    });

    document.getElementById("cm-send").addEventListener("click", onSend);
  }

  async function onSend() {
    const selected = [...document.querySelectorAll(".cm-list:checked")].map((cb) => ({ mail: cb.dataset.mail, count: +cb.dataset.count || 0 }));
    const subject = (document.getElementById("cm-subject").value || "").trim();
    const bodyText = (document.getElementById("cm-body").value || "").trim();
    if (!selected.length) { alert("Select at least one list to send to."); return; }
    if (!subject) { alert("Enter a subject."); return; }
    if (!bodyText) { alert("Enter a message."); return; }
    const total = selected.reduce((s, x) => s + x.count, 0);
    const lines = selected.map((s) => `<li><code>${esc(s.mail)}</code> <span class="muted">(~${s.count})</span></li>`).join("");
    let warn = "";
    if (total > 9000) warn = `<p style="background:#fdecea;border:1px solid #f5c2c0;padding:8px;border-radius:6px">⚠ This single send (~${total}) would consume nearly your whole ~${EXO_DAILY}/day Microsoft sending limit. Consider sending to fewer lists at a time.</p>`;
    else if (total >= 5000) warn = `<p style="background:#fff7e6;border:1px solid #ffe1a8;padding:8px;border-radius:6px">This expands to ~${total} individual deliveries, which count toward your ~${EXO_DAILY}/day Microsoft limit. If you've sent large batches today, consider splitting across days.</p>`;
    const ok = await confirmModal({
      title: `Send to ${selected.length} list${selected.length > 1 ? "s" : ""}?`,
      okLabel: "Send now",
      body: `<p>You're about to email <strong>~${total}</strong> recipients (BCC) across:</p>
             <ul style="max-height:220px;overflow:auto">${lines}</ul>
             ${warn}
             <p class="muted">People on more than one list are de-duplicated by Exchange, so the real number may be a little lower. Subject: <em>${esc(subject)}</em></p>`
    });
    if (!ok) return;
    let myAddr = "";
    try {
      const me = await GRAPH.getMe();
      myAddr = me.mail || me.userPrincipalName;
      await GRAPH.sendMail([myAddr], subject, wrapBranded(bodyText), { bcc: selected.map((s) => s.mail), saveToSentItems: true });
      GRAPH.logAuditEntry({ actor: myAddr, action: "emaillist send message", targetGroup: selected.map((s) => s.mail).join(", "), result: "Success", notes: JSON.stringify({ subject, lists: selected.map((s) => s.mail), estRecipients: total, via: "bcc" }) });
      toast(`Sent to ${selected.length} list${selected.length > 1 ? "s" : ""} (~${total} recipients).`);
      load();
    } catch (e) {
      try { GRAPH.logAuditEntry({ actor: myAddr, action: "emaillist send message", targetGroup: selected.map((s) => s.mail).join(", "), result: "Failure", errorDetail: e.message, notes: JSON.stringify({ subject }) }); } catch (_) { /* non-fatal */ }
      alert("Send failed: " + e.message);
    }
  }

  return { load };
})();
