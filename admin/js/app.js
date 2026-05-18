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

  // Detail view — also caches current group ID for the add/remove handlers
  let currentDetailGroup = null;

  async function openGroupDetail(groupId) {
    const g = state.groups.find((x) => x.id === groupId);
    if (!g) return;
    currentDetailGroup = g;
    $("group-detail-name").textContent = g.displayName || "Group";
    $("group-detail-mail").textContent = g.mail || "";
    $("owners-tbody").innerHTML = `<tr><td colspan="4" class="loading">Loading…</td></tr>`;
    $("members-tbody").innerHTML = `<tr><td colspan="4" class="loading">Loading…</td></tr>`;
    $("owners-count").textContent = "";
    $("members-count").textContent = "";
    show("groupDetail");
    await refreshDetail();
  }

  async function refreshDetail() {
    if (!currentDetailGroup) return;
    try {
      const [owners, members] = await Promise.all([
        GRAPH.listGroupOwners(currentDetailGroup.id),
        GRAPH.listGroupMembers(currentDetailGroup.id),
      ]);
      $("owners-count").textContent = owners.length;
      $("members-count").textContent = members.length;
      $("owners-tbody").innerHTML = renderPeopleRows(owners, "no directors assigned", "owner");
      $("members-tbody").innerHTML = renderPeopleRows(members, "no members", "member");
      // Update the cached owner count in the groups list too
      state.ownerCounts.set(currentDetailGroup.id, owners.length);
    } catch (err) {
      showError("Failed to load group detail: " + err.message);
    }
  }

  function renderPeopleRows(people, emptyMsg, role) {
    if (!people || people.length === 0) {
      return `<tr><td colspan="4" class="muted">${escapeHtml(emptyMsg)}</td></tr>`;
    }
    return people.map((p) => `<tr>
      <td>${escapeHtml(p.displayName)}</td>
      <td>${escapeHtml(p.jobTitle || "")}</td>
      <td>${p.mail ? `<a href="mailto:${escapeHtml(p.mail)}">${escapeHtml(p.mail)}</a>` : `<span class="muted">—</span>`}</td>
      <td class="row-actions"><button class="btn-remove" data-user-id="${escapeHtml(p.id)}" data-name="${escapeHtml(p.displayName)}" data-role="${role}" aria-label="Remove">×</button></td>
    </tr>`).join("");
  }

  $("back-to-groups-btn").addEventListener("click", () => {
    show("groups");
    renderGroupsTable(); // refresh cached owner counts that may have changed
  });

  // ---- Phase 2.1: Add/Remove directors + members ----

  // Smart remove: open the confirm modal with per-group vs full-offboard options.
  let pendingRemove = null; // { userId, userName, role, btn, otherManagedGroups }

  function wireRemoveHandlers() {
    ["owners-tbody", "members-tbody"].forEach((id) => {
      $(id).addEventListener("click", async (e) => {
        const btn = e.target.closest(".btn-remove");
        if (!btn) return;
        const userId = btn.dataset.userId;
        const userName = btn.dataset.name;
        const role = btn.dataset.role; // "owner" or "member"
        if (!currentDetailGroup || !userId) return;

        btn.disabled = true;
        btn.textContent = "…";

        // Query what other managed groups this user is in (excluding this one).
        let otherManagedGroups = [];
        try {
          const allGroups = await GRAPH.getUserMemberOf(userId);
          const managedIds = new Set(state.groups.map((g) => g.id));
          otherManagedGroups = allGroups.filter(
            (g) => managedIds.has(g.id) && g.id !== currentDetailGroup.id
          );
        } catch (err) {
          // Non-fatal — just means we can't show the "also in" list. Continue with empty list.
          console.warn("Could not fetch user's other groups:", err);
        }

        pendingRemove = { userId, userName, role, btn, otherManagedGroups };
        openRemovePanel();
      });
    });
  }
  wireRemoveHandlers();

  function openRemovePanel() {
    const { userName, role, otherManagedGroups } = pendingRemove;
    const label = role === "owner" ? "director (owner)" : "member";
    $("remove-panel-title").textContent = `Remove ${userName}`;
    const groupLine = `${userName} is currently a ${label} of <strong>${escapeHtml(currentDetailGroup.displayName)}</strong>.`;
    let bodyHtml = groupLine;
    if (otherManagedGroups.length > 0) {
      bodyHtml += ` They are also in <strong>${otherManagedGroups.length}</strong> other managed group${otherManagedGroups.length === 1 ? "" : "s"}:`;
      $("remove-other-groups").innerHTML = otherManagedGroups
        .map((g) => `<li>${escapeHtml(g.displayName)}</li>`)
        .join("");
      $("remove-other-groups").classList.remove("hidden");
    } else {
      bodyHtml += ` They are not in any other managed groups.`;
      $("remove-other-groups").innerHTML = "";
      $("remove-other-groups").classList.add("hidden");
    }
    $("remove-panel-body").innerHTML = bodyHtml;
    $("remove-panel").classList.remove("hidden");
  }

  function closeRemovePanel() {
    if (pendingRemove && pendingRemove.btn) {
      pendingRemove.btn.disabled = false;
      pendingRemove.btn.textContent = "×";
    }
    pendingRemove = null;
    $("remove-panel").classList.add("hidden");
  }

  $("remove-panel-close").addEventListener("click", closeRemovePanel);

  $("remove-this-group-btn").addEventListener("click", async () => {
    if (!pendingRemove || !currentDetailGroup) return;
    const { userId, userName, role } = pendingRemove;
    $("remove-this-group-btn").disabled = true;
    $("remove-offboard-btn").disabled = true;
    try {
      if (role === "owner") {
        await GRAPH.removeOwner(currentDetailGroup.id, userId);
      } else {
        await GRAPH.removeMember(currentDetailGroup.id, userId);
      }
      logAction(`removed ${role} (per-group only)`, userName, userId);
      $("remove-panel").classList.add("hidden");
      pendingRemove = null;
      await refreshDetail();
    } catch (err) {
      showError(`Failed to remove ${userName}: ${err.message}`);
    } finally {
      $("remove-this-group-btn").disabled = false;
      $("remove-offboard-btn").disabled = false;
    }
  });

  $("remove-offboard-btn").addEventListener("click", async () => {
    if (!pendingRemove || !currentDetailGroup) return;
    const { userId, userName, otherManagedGroups } = pendingRemove;
    const confirmMsg = `Offboard ${userName} fully?\n\nThis will:\n  • Remove from this group + ${otherManagedGroups.length} other managed group(s)\n  • Remove the EVAA license\n  • Disable the account\n\nThis cannot be undone via this UI (re-enable via Entra admin center if needed).`;
    if (!confirm(confirmMsg)) return;

    $("remove-this-group-btn").disabled = true;
    $("remove-offboard-btn").disabled = true;
    const errors = [];

    try {
      // 1. Remove from THIS group (covers the current view's role)
      try {
        if (pendingRemove.role === "owner") {
          await GRAPH.removeOwner(currentDetailGroup.id, userId);
        } else {
          await GRAPH.removeMember(currentDetailGroup.id, userId);
        }
      } catch (err) {
        errors.push(`current group: ${err.message}`);
      }

      // 2. Remove from all OTHER managed groups (members; owners are role-specific but we
      //    aggressively unlink as member to be thorough).
      for (const g of otherManagedGroups) {
        try {
          await GRAPH.removeMember(g.id, userId);
        } catch (err) {
          // Try owner ref too (in case they were owner of that group, not member)
          try { await GRAPH.removeOwner(g.id, userId); } catch (_) {
            errors.push(`${g.displayName}: ${err.message}`);
          }
        }
      }

      // 3. Remove EVAA license
      try {
        await GRAPH.removeUserLicense(userId);
      } catch (err) {
        errors.push(`license: ${err.message}`);
      }

      // 4. Disable account
      try {
        await GRAPH.disableUserAccount(userId);
      } catch (err) {
        errors.push(`disable account: ${err.message}`);
      }

      logAction("offboarded fully", userName, userId);

      if (errors.length > 0) {
        showError(`Offboard partial: ${errors.length} step(s) failed. ${errors.join("; ")}`);
      }
      $("remove-panel").classList.add("hidden");
      pendingRemove = null;
      await refreshDetail();
    } finally {
      $("remove-this-group-btn").disabled = false;
      $("remove-offboard-btn").disabled = false;
    }
  });

  // Add Director / Add Member — opens inline search panel
  let addTarget = null; // "owner" or "member"

  function openAddPanel(role) {
    if (!currentDetailGroup) return;
    addTarget = role;
    const label = role === "owner" ? "director" : "member";
    $("add-panel-title").textContent = `Add ${label} to "${currentDetailGroup.displayName}"`;
    $("add-search-input").value = "";
    $("add-search-results").innerHTML = "";
    $("add-panel").classList.remove("hidden");
    $("add-search-input").focus();
  }

  function closeAddPanel() {
    addTarget = null;
    $("add-panel").classList.add("hidden");
  }

  $("add-owner-btn").addEventListener("click", () => openAddPanel("owner"));
  $("add-member-btn").addEventListener("click", () => openAddPanel("member"));
  $("add-panel-close").addEventListener("click", closeAddPanel);

  let searchDebounce;
  $("add-search-input").addEventListener("input", (e) => {
    const q = e.target.value.trim();
    clearTimeout(searchDebounce);
    if (q.length < 2) { $("add-search-results").innerHTML = ""; return; }
    searchDebounce = setTimeout(async () => {
      try {
        const users = await GRAPH.searchUsers(q);
        renderUserSearchResults(users);
      } catch (err) {
        showError("Search failed: " + err.message);
      }
    }, 250);
  });

  function renderUserSearchResults(users) {
    const container = $("add-search-results");
    if (!users.length) { container.innerHTML = `<p class="muted">No users found.</p>`; return; }
    container.innerHTML = users.map((u) => `<button class="user-result" data-user-id="${escapeHtml(u.id)}" data-name="${escapeHtml(u.displayName)}">
      <span class="user-name">${escapeHtml(u.displayName)}</span>
      <span class="user-mail muted">${escapeHtml(u.mail || u.userPrincipalName || "")}</span>
    </button>`).join("");
    container.querySelectorAll(".user-result").forEach((btn) => {
      btn.addEventListener("click", () => addPickedUser(btn.dataset.userId, btn.dataset.name));
    });
  }

  async function addPickedUser(userId, userName) {
    if (!currentDetailGroup || !addTarget) return;
    const label = addTarget === "owner" ? "director (owner)" : "member";
    if (!confirm(`Add ${userName} as ${label} of "${currentDetailGroup.displayName}"?`)) return;

    try {
      if (addTarget === "owner") {
        await GRAPH.addOwner(currentDetailGroup.id, userId);
      } else {
        await GRAPH.addMember(currentDetailGroup.id, userId);
      }
      logAction(`added ${addTarget}`, userName, userId);
      closeAddPanel();
      await refreshDetail();
    } catch (err) {
      showError(`Failed to add ${userName}: ${err.message}`);
    }
  }

  // Lightweight audit log — console for now; SharePoint AdminActionLog list is Phase 2.2.
  function logAction(action, targetName, targetId) {
    const who = AUTH.getAccount()?.username || "(unknown)";
    const group = currentDetailGroup ? `${currentDetailGroup.displayName} (${currentDetailGroup.id})` : "(unknown)";
    const entry = { ts: new Date().toISOString(), admin: who, action, targetName, targetId, group };
    console.log("[AUDIT]", JSON.stringify(entry));
  }
})();
