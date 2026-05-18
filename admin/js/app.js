// EVAA Admin Portal — main app
// Phase 1: read-only. Sign in, verify admin status, browse groups, drill into members + owners.
// Phase 2 (later): add/remove members, add/remove owners, audit log.

(async function main() {
  // DOM refs
  const $ = (id) => document.getElementById(id);
  const views = {
    signin: $("signin-view"),
    noAccess: $("no-access-view"),
    loading: $("loading-view"),
    groups: $("groups-view"),
    groupDetail: $("group-detail-view"),
  };
  const userArea = $("user-area");
  const userName = $("user-name");
  const errorBanner = $("error-banner");
  const errorText = $("error-text");
  const envInfo = $("env-info");

  function show(viewKey) {
    Object.entries(views).forEach(([k, el]) => el.classList.toggle("hidden", k !== viewKey));
  }

  function showError(msg) {
    errorText.textContent = msg;
    errorBanner.classList.remove("hidden");
  }
  $("error-close").addEventListener("click", () => errorBanner.classList.add("hidden"));

  envInfo.textContent = `Tenant evaasports.org · ${new Date().getFullYear()}`;

  // Wire sign-out buttons immediately (safe — guarded internally)
  $("sign-out-btn").addEventListener("click", () => AUTH.signOut());
  $("signout-from-noaccess-btn").addEventListener("click", () => AUTH.signOut());

  // Boot: MSAL init + post-redirect handling.
  // Sign-in view stays hidden until init resolves so users can't click before MSAL is ready.
  show("loading");
  let account;
  try {
    account = await AUTH.init();
  } catch (err) {
    showError("Auth init failed: " + err.message);
    // Even on failure, show the signin view so user has something to do.
    show("signin");
    return;
  }

  if (!account) {
    // Init succeeded but no signed-in account → show signin view with the button enabled
    const signInBtn = $("sign-in-btn");
    signInBtn.disabled = false;
    signInBtn.addEventListener("click", () => AUTH.signIn().catch((e) => showError(e.message)));
    show("signin");
    return;
  }

  // Signed in — show user, check admin, load groups
  userName.textContent = account.name || account.username || "";
  userArea.classList.remove("hidden");

  // Stay on loading view while we check admin + fetch groups
  $("loading-text").textContent = "Checking access…";

  let isAdmin;
  try {
    isAdmin = await GRAPH.isPortalAdmin();
  } catch (err) {
    showError("Admin check failed: " + err.message);
    show("signin");
    return;
  }

  if (!isAdmin) {
    show("noAccess");
    return;
  }

  // Load all managed groups
  $("loading-text").textContent = "Loading groups…";
  let groups;
  try {
    groups = await GRAPH.listManagedGroups();
  } catch (err) {
    showError("Failed to load groups: " + err.message);
    show("signin");
    return;
  }

  // State for groups view
  const state = {
    groups,
    ownerCounts: new Map(), // groupId -> count (lazily populated)
    filterText: "",
    sortCol: 0, // 0=name asc default
    sortDir: 1,
  };

  // Sort + filter, returns view of state.groups
  function filteredSorted() {
    const f = state.filterText.toLowerCase();
    const out = state.groups.filter((g) =>
      !f || (g.displayName || "").toLowerCase().includes(f) || (g.mail || "").toLowerCase().includes(f)
    );
    out.sort((a, b) => {
      const dir = state.sortDir;
      const col = state.sortCol;
      const aVal = col === 0 ? (a.displayName || "") : col === 2 ? (a.mail || "") : (state.ownerCounts.get(a.id) ?? -1);
      const bVal = col === 0 ? (b.displayName || "") : col === 2 ? (b.mail || "") : (state.ownerCounts.get(b.id) ?? -1);
      if (aVal < bVal) return -1 * dir;
      if (aVal > bVal) return 1 * dir;
      return 0;
    });
    return out;
  }

  function renderGroupsTable() {
    const tbody = $("groups-tbody");
    const rows = filteredSorted();
    $("groups-count").textContent = `${rows.length} group${rows.length === 1 ? "" : "s"}`;
    tbody.innerHTML = rows.map((g) => {
      const ownerCount = state.ownerCounts.get(g.id);
      const ownerCell = ownerCount === undefined
        ? `<span class="muted">…</span>`
        : ownerCount === 0
          ? `<span class="warning-chip">no director assigned</span>`
          : `${ownerCount}`;
      const mailCell = g.mail
        ? `<a href="mailto:${escapeHtml(g.mail)}" onclick="event.stopPropagation()">${escapeHtml(g.mail)}</a>`
        : `<span class="muted">—</span>`;
      return `<tr data-group-id="${escapeHtml(g.id)}">
        <td>${escapeHtml(g.displayName)}</td>
        <td>${ownerCell}</td>
        <td>${mailCell}</td>
      </tr>`;
    }).join("");
    // Click handler on rows
    tbody.querySelectorAll("tr").forEach((row) => {
      row.addEventListener("click", () => openGroupDetail(row.dataset.groupId));
    });
    // Update sort arrows
    document.querySelectorAll("#groups-table thead th").forEach((th, i) => {
      th.classList.toggle("sorted", i === state.sortCol);
      const arrow = th.querySelector(".arrow");
      if (arrow) arrow.textContent = i === state.sortCol ? (state.sortDir === 1 ? "▲" : "▼") : "▽";
    });
  }

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // Wire filter + sort
  $("group-filter").addEventListener("input", (e) => {
    state.filterText = e.target.value;
    renderGroupsTable();
  });
  document.querySelectorAll("#groups-table thead th[data-col]").forEach((th) => {
    th.addEventListener("click", () => {
      const col = parseInt(th.dataset.col, 10);
      if (state.sortCol === col) state.sortDir *= -1;
      else { state.sortCol = col; state.sortDir = 1; }
      renderGroupsTable();
    });
  });

  // Initial render
  show("groups");
  renderGroupsTable();

  // Lazily hydrate owner counts (in parallel, with light throttling)
  // This avoids blocking the initial render on N flow calls.
  hydrateOwnerCounts(state.groups);

  async function hydrateOwnerCounts(groups) {
    // Process in small batches to avoid 429s
    const BATCH = 4;
    for (let i = 0; i < groups.length; i += BATCH) {
      const batch = groups.slice(i, i + BATCH);
      await Promise.all(batch.map(async (g) => {
        try {
          const owners = await GRAPH.listGroupOwners(g.id);
          state.ownerCounts.set(g.id, owners.length);
        } catch (err) {
          console.warn(`Owner fetch failed for ${g.displayName}:`, err);
          state.ownerCounts.set(g.id, -1); // mark as errored
        }
      }));
      renderGroupsTable();
    }
  }

  // Detail view
  async function openGroupDetail(groupId) {
    const g = state.groups.find((x) => x.id === groupId);
    if (!g) return;
    $("group-detail-name").textContent = g.displayName || "Group";
    $("group-detail-mail").textContent = g.mail || "";
    $("owners-tbody").innerHTML = `<tr><td colspan="3" class="loading">Loading…</td></tr>`;
    $("members-tbody").innerHTML = `<tr><td colspan="3" class="loading">Loading…</td></tr>`;
    $("owners-count").textContent = "";
    $("members-count").textContent = "";
    show("groupDetail");

    try {
      const [owners, members] = await Promise.all([
        GRAPH.listGroupOwners(groupId),
        GRAPH.listGroupMembers(groupId),
      ]);
      $("owners-count").textContent = owners.length;
      $("members-count").textContent = members.length;
      $("owners-tbody").innerHTML = renderPeopleRows(owners, "no directors assigned");
      $("members-tbody").innerHTML = renderPeopleRows(members, "no members");
    } catch (err) {
      showError("Failed to load group detail: " + err.message);
    }
  }

  function renderPeopleRows(people, emptyMsg) {
    if (!people || people.length === 0) {
      return `<tr><td colspan="3" class="muted">${escapeHtml(emptyMsg)}</td></tr>`;
    }
    return people.map((p) => `<tr>
      <td>${escapeHtml(p.displayName)}</td>
      <td>${escapeHtml(p.jobTitle || "")}</td>
      <td>${p.mail ? `<a href="mailto:${escapeHtml(p.mail)}">${escapeHtml(p.mail)}</a>` : `<span class="muted">—</span>`}</td>
    </tr>`).join("");
  }

  $("back-to-groups-btn").addEventListener("click", () => {
    show("groups");
  });
})();
