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
    help: $("help-view"),
    groups: $("groups-view"),
    groupDetail: $("group-detail-view"),
    members: $("members-view"),
    userDetail: $("user-detail-view"),
  };
  const tabNav = $("tab-nav");
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

  // Toast for success feedback (green, auto-dismisses after 4s).
  let toastTimer = null;
  function showToast(msg, kind = "success") {
    const banner = $("toast-banner");
    $("toast-text").textContent = msg;
    banner.className = "toast-banner toast-" + kind;
    banner.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => banner.classList.add("hidden"), 4000);
  }

  // Branded replacement for native confirm(). Returns Promise<boolean>.
  // Supports HTML body (so we can preserve formatting), custom OK button label/style.
  function confirmCustom({ body, okLabel = "OK", okClass = "btn-primary", title = "Email Admin Portal asks…" }) {
    return new Promise((resolve) => {
      $("confirm-modal-title").textContent = title;
      $("confirm-modal-body").innerHTML = body;
      const okBtn = $("confirm-modal-ok");
      const cancelBtn = $("confirm-modal-cancel");
      okBtn.textContent = okLabel;
      okBtn.className = okClass;
      const modal = $("confirm-modal");
      modal.classList.remove("hidden");
      function cleanup(result) {
        modal.classList.add("hidden");
        okBtn.removeEventListener("click", onOk);
        cancelBtn.removeEventListener("click", onCancel);
        document.removeEventListener("keydown", onKey);
        resolve(result);
      }
      function onOk() { cleanup(true); }
      function onCancel() { cleanup(false); }
      function onKey(e) {
        if (e.key === "Escape") cleanup(false);
        else if (e.key === "Enter") cleanup(true);
      }
      okBtn.addEventListener("click", onOk);
      cancelBtn.addEventListener("click", onCancel);
      document.addEventListener("keydown", onKey);
    });
  }

  envInfo.textContent = `Tenant evaasports.org · ${new Date().getFullYear()}`;

  // Wire sign-out buttons immediately (safe — guarded internally)
  $("sign-out-btn").addEventListener("click", () => AUTH.signOut());
  $("signout-from-noaccess-btn").addEventListener("click", () => AUTH.signOut());

  // Help button — switches to the Help view (no tab highlighted while on Help).
  $("help-btn").addEventListener("click", () => {
    show("help");
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
  });

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
    signInBtn.addEventListener("click", async () => {
      try {
        const acct = await AUTH.signIn();
        if (acct) window.location.reload();
      } catch (e) {
        showError(e.message);
      }
    });
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

  // Load all managed groups first (we need them either way — admin gets all,
  // owner gets the subset they own).
  $("loading-text").textContent = "Loading groups…";
  let allManagedGroups;
  try {
    allManagedGroups = await GRAPH.listManagedGroups();
  } catch (err) {
    showError("Failed to load groups: " + err.message);
    show("signin");
    return;
  }

  // Tri-state role: admin / owner / none.
  // Admins see everything. Owners see only the groups they own (with approval-routed Add/Remove).
  let role = "none";
  let groups;
  if (isAdmin) {
    role = "admin";
    groups = allManagedGroups;
  } else {
    $("loading-text").textContent = "Checking group ownership…";
    let ownedIds = [];
    try {
      ownedIds = await GRAPH.getOwnedManagedGroupIds(allManagedGroups.map((g) => g.id));
    } catch (err) {
      showError("Owner check failed: " + err.message);
      show("noAccess");
      return;
    }
    if (ownedIds.length === 0) {
      show("noAccess");
      return;
    }
    role = "owner";
    const ownedSet = new Set(ownedIds);
    groups = allManagedGroups.filter((g) => ownedSet.has(g.id));
  }

  // Tag the body so CSS can hide admin-only elements when role=owner.
  document.body.classList.add(`role-${role}`);

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

  // Wire tab nav (now that admin is confirmed and groups are loaded)
  tabNav.classList.remove("hidden");
  let activeTab = null; // null = on Help view; "groups" or "members" otherwise
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      activeTab = btn.dataset.tab;
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b === btn));
      if (activeTab === "groups") show("groups");
      else { show("members"); ensureMembersLoaded(); }
    });
  });

  // Land on Help view by default so admins see the orientation page first.
  // Render the groups table in background so it's ready when they click the tab.
  renderGroupsTable();
  show("help");
  document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));

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

  function renderPeopleRows(people, emptyMsg, rowRole) {
    if (!people || people.length === 0) {
      return `<tr><td colspan="4" class="muted">${escapeHtml(emptyMsg)}</td></tr>`;
    }
    return people.map((p) => `<tr>
      <td><button class="link-button" data-jump-user-id="${escapeHtml(p.id)}" data-user-name="${escapeHtml(p.displayName)}">${escapeHtml(p.displayName)}</button></td>
      <td>${escapeHtml(p.jobTitle || "")}</td>
      <td>${p.mail ? `<a href="mailto:${escapeHtml(p.mail)}" onclick="event.stopPropagation()">${escapeHtml(p.mail)}</a>` : `<span class="muted">—</span>`}</td>
      <td class="row-actions"><button class="btn-remove" data-user-id="${escapeHtml(p.id)}" data-name="${escapeHtml(p.displayName)}" data-email="${escapeHtml(p.mail || "")}" data-role="${rowRole}" aria-label="Remove">×</button></td>
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
        // Cross-nav: clicking the person's name jumps to their user detail page
        const jumpBtn = e.target.closest(".link-button[data-jump-user-id]");
        if (jumpBtn) {
          jumpToUser(jumpBtn.dataset.jumpUserId, { id: jumpBtn.dataset.jumpUserId, displayName: jumpBtn.dataset.userName });
          return;
        }
        const btn = e.target.closest(".btn-remove");
        if (!btn) return;
        const userId = btn.dataset.userId;
        const userName = btn.dataset.name;
        const rowRole = btn.dataset.role; // "owner" or "member" (relative to this group)
        if (!currentDetailGroup || !userId) return;

        btn.disabled = true;
        btn.textContent = "…";

        // OWNER MODE: skip the 3-option modal. File a single Remove request and let
        // the admin approval flow decide whether/how to fully offboard.
        if (role === "owner") {
          try {
            const memberEmail = btn.dataset.email || "";
            const parts = (userName || "").split(/\s+/);
            const firstName = parts[0] || userName || "";
            const lastName = parts.slice(1).join(" ") || "";
            const ok = await confirmCustom({
              body: `<p>Submit a request to remove <strong>${escapeHtml(userName)}</strong> from <strong>${escapeHtml(currentDetailGroup.displayName)}</strong>?</p>
                <p class="muted">An admin will be notified to approve the removal. You'll get a confirmation once it's processed.</p>`,
              okLabel: "Submit removal request",
              okClass: "btn-warning",
            });
            if (!ok) { btn.disabled = false; btn.textContent = "×"; return; }
            await GRAPH.createMemberRequest({
              requestType: "Remove",
              sportDisplayName: currentDetailGroup.displayName,
              firstName,
              lastName,
              memberId: userId,
              memberEmail,
            });
            showToast(`Removal request submitted for ${userName}.`);
            logAction("submitted Remove request", userName, userId, { group: currentDetailGroup.displayName });
          } catch (err) {
            showError(`Could not submit removal request: ${err.message}`);
          } finally {
            btn.disabled = false;
            btn.textContent = "×";
          }
          return;
        }

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

        pendingRemove = { userId, userName, role: rowRole, btn, otherManagedGroups };
        openRemovePanel();
      });
    });
  }
  wireRemoveHandlers();

  function openRemovePanel() {
    const { userName, role, otherManagedGroups } = pendingRemove;
    const label = role === "owner" ? "director (owner)" : "member";
    $("remove-panel-title").textContent = `Remove ${userName}`;
    // This handler always runs in group context (called from group-detail × button).
    // Make sure the "Remove from this group only" button is visible.
    $("remove-this-group-btn").style.display = "";
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

  // "Disable account only (preserve data)" — middle option.
  // Removes from all managed groups, disables account, KEEPS the license so mailbox/OneDrive
  // stay preserved indefinitely. Emails portaladmins@ so the group can decide what to do next.
  // Works in either group context (from group detail x) or user context (from user detail offboard).
  $("remove-preserve-btn").addEventListener("click", async () => {
    if (!pendingRemove) return;
    const { userId, userName, otherManagedGroups, context } = pendingRemove;
    const fromGroupCtx = context !== "user" && currentDetailGroup;
    const totalGroups = (fromGroupCtx ? 1 : 0) + otherManagedGroups.length;
    const ok = await confirmCustom({
      body: `<p>Disable <strong>${escapeHtml(userName)}</strong>'s account and preserve their data?</p>
        <p>This will:</p>
        <ul>
          <li>Remove from ${totalGroups} managed group${totalGroups === 1 ? "" : "s"}</li>
          <li>Disable the account (can't sign in)</li>
          <li><strong>KEEP the EVAA license</strong> (mailbox + OneDrive preserved indefinitely)</li>
          <li>Email portaladmins@evaasports.org with a summary</li>
        </ul>
        <p class="muted">License cost (~$3/mo) continues until you fully offboard.</p>`,
      okLabel: "Disable & preserve",
      okClass: "btn-warning",
    });
    if (!ok) return;

    $("remove-this-group-btn").disabled = true;
    $("remove-preserve-btn").disabled = true;
    $("remove-offboard-btn").disabled = true;
    const errors = [];
    const allRemovedGroups = []; // for the admin email

    try {
      // 1. If in group context, remove from this group first
      if (fromGroupCtx) {
        try {
          if (pendingRemove.role === "owner") await GRAPH.removeOwner(currentDetailGroup.id, userId);
          else await GRAPH.removeMember(currentDetailGroup.id, userId);
          allRemovedGroups.push({ displayName: currentDetailGroup.displayName });
        } catch (err) { errors.push(`${currentDetailGroup.displayName}: ${err.message}`); }
      }

      // 2. Remove from other managed groups (member + owner; whichever applies)
      for (const g of otherManagedGroups) {
        try {
          await GRAPH.removeMember(g.id, userId);
          allRemovedGroups.push({ displayName: g.displayName });
        }
        catch (_) {
          try {
            await GRAPH.removeOwner(g.id, userId);
            allRemovedGroups.push({ displayName: g.displayName });
          }
          catch (e) { errors.push(`${g.displayName}: ${e.message}`); }
        }
      }

      // 3. Disable the account (but do NOT remove license)
      try { await GRAPH.disableUserAccount(userId); }
      catch (err) { errors.push(`disable: ${err.message}`); }

      // 4. Notify portal admins by email
      try {
        const adminMail = AUTH.getAccount()?.username || "(unknown admin)";
        const adminName = AUTH.getAccount()?.name || adminMail;
        const groupContextName = fromGroupCtx ? currentDetailGroup.displayName : "user detail page";
        const removedGroupsHtml = allRemovedGroups.length
          ? allRemovedGroups.map((g) => `<li>${escapeHtml(g.displayName)}</li>`).join("")
          : `<li><em>(no managed groups)</em></li>`;
        const html = buildPreserveDataAdminEmailHtml({
          adminName, adminMail, userName, userId,
          groupContextName, removedGroupsHtml,
        });
        await GRAPH.sendMail(
          ["portaladmins@evaasports.org"],
          `[EVAA Admin] ${userName} disabled — data preserved`,
          html,
        );
      } catch (err) {
        errors.push(`admin notify email: ${err.message}`);
      }

      logAction("disabled user with data preserved", userName, userId, { groupCount: allRemovedGroups.length, context: fromGroupCtx ? "group" : "user" });

      if (errors.length) showError(`Preserve-data action partial: ${errors.join("; ")}`);
      else showToast(`${userName} disabled — data preserved, admins notified`);
      $("remove-panel").classList.add("hidden");
      pendingRemove = null;
      if (fromGroupCtx) await refreshDetail();
      else await refreshUserDetail();
    } finally {
      $("remove-this-group-btn").disabled = false;
      $("remove-preserve-btn").disabled = false;
      $("remove-offboard-btn").disabled = false;
    }
  });

  $("remove-offboard-btn").addEventListener("click", async () => {
    if (!pendingRemove) return;
    const { userId, userName, otherManagedGroups, context } = pendingRemove;
    const fromGroupCtx = context !== "user" && currentDetailGroup;
    const totalGroups = (fromGroupCtx ? 1 : 0) + otherManagedGroups.length;
    const ok = await confirmCustom({
      body: `<p>Offboard <strong>${escapeHtml(userName)}</strong> fully?</p>
        <p>This will:</p>
        <ul>
          <li>Remove from ${totalGroups} managed group${totalGroups === 1 ? "" : "s"}</li>
          <li>Remove the EVAA license</li>
          <li>Disable the account</li>
        </ul>
        <p class="warn-block">⚠ Exchange will start a <strong>30-day countdown</strong> to permanently delete the mailbox and OneDrive.</p>
        <p class="muted">Re-enable via this portal or Entra admin center within 30 days to recover.</p>`,
      okLabel: "Offboard fully",
      okClass: "btn-danger",
    });
    if (!ok) return;

    $("remove-this-group-btn").disabled = true;
    $("remove-preserve-btn").disabled = true;
    $("remove-offboard-btn").disabled = true;
    const errors = [];

    try {
      // 1. If in group context, remove from this group
      if (fromGroupCtx) {
        try {
          if (pendingRemove.role === "owner") await GRAPH.removeOwner(currentDetailGroup.id, userId);
          else await GRAPH.removeMember(currentDetailGroup.id, userId);
        } catch (err) { errors.push(`${currentDetailGroup.displayName}: ${err.message}`); }
      }

      // 2. Remove from all OTHER managed groups
      for (const g of otherManagedGroups) {
        try { await GRAPH.removeMember(g.id, userId); }
        catch (_) {
          try { await GRAPH.removeOwner(g.id, userId); }
          catch (e) { errors.push(`${g.displayName}: ${e.message}`); }
        }
      }

      // 3. Remove EVAA license (idempotent — if already gone, treat as success)
      try { await GRAPH.removeUserLicense(userId); }
      catch (err) {
        if (!/does not have a corresponding license/i.test(err.message)) {
          errors.push(`license: ${err.message}`);
        }
      }

      // 4. Disable account
      try { await GRAPH.disableUserAccount(userId); }
      catch (err) { errors.push(`disable account: ${err.message}`); }

      logAction("offboarded fully", userName, userId, { context: fromGroupCtx ? "group" : "user" });

      if (errors.length > 0) {
        showError(`Offboard partial: ${errors.length} step(s) failed. ${errors.join("; ")}`);
      } else {
        showToast(`${userName} offboarded fully`);
      }
      $("remove-panel").classList.add("hidden");
      pendingRemove = null;
      if (fromGroupCtx) await refreshDetail();
      else await refreshUserDetail();
    } finally {
      $("remove-this-group-btn").disabled = false;
      $("remove-preserve-btn").disabled = false;
      $("remove-offboard-btn").disabled = false;
    }
  });

  // Add Director / Add Member — opens inline search panel
  let addTarget = null; // "owner" or "member"

  function openAddPanel(role) {
    if (!currentDetailGroup) return;
    addTarget = role;
    const label = role === "owner" ? "owner" : "member";
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
  // Owner mode: + Add Member skips the user-search panel and opens the Create User
  // form directly (pre-filled with current group, group selector locked). On submit
  // we file a MemberRequests row instead of provisioning directly — admin approves.
  $("add-member-btn").addEventListener("click", () => {
    if (role === "owner") {
      openCreateUserPanel(currentDetailGroup ? currentDetailGroup.id : null);
      // Lock the group selector — owners can only request additions to the current group.
      const sel = $("cu-group");
      if (sel) sel.disabled = true;
    } else {
      openAddPanel("member");
    }
  });
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
    const groupName = currentDetailGroup.displayName;

    let body;
    if (addTarget === "owner") {
      body = `<p>Are you sure you want to add <strong>${escapeHtml(userName)}</strong> as an owner of <strong>${escapeHtml(groupName)}</strong>?</p>
        <p>Owners (also known as <em>directors</em>) can <strong>add and remove members</strong> from this group and will receive admin notifications for this group.</p>`;
    } else {
      body = `<p>Add <strong>${escapeHtml(userName)}</strong> as a member of <strong>${escapeHtml(groupName)}</strong>?</p>
        <p class="muted">Members can read the group's emails and shared files but cannot manage other members.</p>`;
    }
    const ok = await confirmCustom({
      body,
      okLabel: addTarget === "owner" ? "Add as Owner" : "Add as Member",
      okClass: "btn-primary",
    });
    if (!ok) return;

    try {
      if (addTarget === "owner") {
        await GRAPH.addOwner(currentDetailGroup.id, userId);
      } else {
        await GRAPH.addMember(currentDetailGroup.id, userId);
      }
      logAction(`added ${addTarget}`, userName, userId);
      closeAddPanel();
      showToast(`${userName} added as ${addTarget === "owner" ? "owner" : "member"}`);
      await refreshDetail();
    } catch (err) {
      showError(`Failed to add ${userName}: ${err.message}`);
    }
  }

  // Lightweight audit log — console for now; SharePoint AdminActionLog list is Phase 2.2.
  function logAction(action, targetName, targetId, extra) {
    const who = AUTH.getAccount()?.username || "(unknown)";
    const group = currentDetailGroup ? `${currentDetailGroup.displayName} (${currentDetailGroup.id})` : null;
    const entry = { ts: new Date().toISOString(), admin: who, action, targetName, targetId };
    if (group) entry.group = group;
    if (extra) Object.assign(entry, extra);
    console.log("[AUDIT]", JSON.stringify(entry));
  }

  // =====================================================================
  // MEMBERS TAB — user search → user detail view → group membership mgmt
  // =====================================================================

  let currentDetailUser = null; // the user currently shown in user-detail-view

  // Members tab state: full user cache, current filter, pagination
  const membersState = {
    allUsers: null, // null = not yet loaded
    filterText: "",
    sortCol: 0, // 0=Name, 1=UPN/Email, 2=Role/jobTitle
    sortDir: 1,
    page: 0,
    pageSize: 25,
  };

  async function ensureMembersLoaded() {
    if (membersState.allUsers !== null) return;
    $("members-list-tbody").innerHTML = `<tr><td colspan="3" class="loading">Loading users…</td></tr>`;
    try {
      const users = await GRAPH.listAllManagedUsers();
      membersState.allUsers = users;
      renderMembersPage();
    } catch (err) {
      $("members-list-tbody").innerHTML = `<tr><td colspan="3" class="muted">Failed to load users: ${escapeHtml(err.message)}</td></tr>`;
      showError("Failed to load users: " + err.message);
    }
  }

  function filteredSortedMembers() {
    const f = membersState.filterText.toLowerCase();
    const out = (membersState.allUsers || []).filter((u) =>
      !f ||
      (u.displayName || "").toLowerCase().includes(f) ||
      (u.userPrincipalName || "").toLowerCase().includes(f) ||
      (u.mail || "").toLowerCase().includes(f) ||
      (u.jobTitle || "").toLowerCase().includes(f)
    );
    out.sort((a, b) => {
      const dir = membersState.sortDir;
      const col = membersState.sortCol;
      const pick = (u) => col === 0 ? (u.displayName || "") : col === 1 ? (u.userPrincipalName || u.mail || "") : (u.jobTitle || "");
      const aVal = pick(a).toLowerCase();
      const bVal = pick(b).toLowerCase();
      if (aVal < bVal) return -1 * dir;
      if (aVal > bVal) return 1 * dir;
      return 0;
    });
    return out;
  }

  function renderMembersPage() {
    const rows = filteredSortedMembers();
    const totalPages = Math.max(1, Math.ceil(rows.length / membersState.pageSize));
    if (membersState.page >= totalPages) membersState.page = totalPages - 1;
    if (membersState.page < 0) membersState.page = 0;
    const start = membersState.page * membersState.pageSize;
    const pageRows = rows.slice(start, start + membersState.pageSize);

    $("members-count").textContent = `${rows.length} user${rows.length === 1 ? "" : "s"}${membersState.filterText ? ` matching "${membersState.filterText}"` : ""}`;

    const tbody = $("members-list-tbody");
    if (!pageRows.length) {
      tbody.innerHTML = `<tr><td colspan="3" class="muted">No matches.</td></tr>`;
    } else {
      tbody.innerHTML = pageRows.map((u) => {
        const disabled = u.accountEnabled === false;
        const disabledBadge = disabled ? ` <span class="account-state-badge state-disabled">disabled</span>` : "";
        const rowClass = disabled ? ' class="row-disabled"' : "";
        return `<tr${rowClass} data-user-id="${escapeHtml(u.id)}">
          <td><button class="link-button" data-jump-user-id="${escapeHtml(u.id)}" data-user-name="${escapeHtml(u.displayName)}">${escapeHtml(u.displayName)}</button>${disabledBadge}</td>
          <td>${escapeHtml(u.userPrincipalName || u.mail || "")}</td>
          <td>${escapeHtml(u.jobTitle || "")}</td>
        </tr>`;
      }).join("");
      tbody.querySelectorAll(".link-button[data-jump-user-id]").forEach((b) => {
        b.addEventListener("click", () => openUserDetail(b.dataset.jumpUserId, { id: b.dataset.jumpUserId, displayName: b.dataset.userName }));
      });
    }

    // Sort arrows
    document.querySelectorAll("#members-list-table thead th[data-mcol]").forEach((th, i) => {
      th.classList.toggle("sorted", i === membersState.sortCol);
      const arrow = th.querySelector(".arrow");
      if (arrow) arrow.textContent = i === membersState.sortCol ? (membersState.sortDir === 1 ? "▲" : "▼") : "▽";
    });

    // Pagination controls
    const pag = $("members-pagination");
    if (totalPages <= 1) {
      pag.innerHTML = "";
    } else {
      pag.innerHTML = `
        <button class="btn-link page-prev" ${membersState.page === 0 ? "disabled" : ""}>← Prev</button>
        <span class="page-indicator">Page ${membersState.page + 1} of ${totalPages}</span>
        <button class="btn-link page-next" ${membersState.page >= totalPages - 1 ? "disabled" : ""}>Next →</button>
      `;
      pag.querySelector(".page-prev")?.addEventListener("click", () => { membersState.page--; renderMembersPage(); });
      pag.querySelector(".page-next")?.addEventListener("click", () => { membersState.page++; renderMembersPage(); });
    }
  }

  $("member-search-input").addEventListener("input", (e) => {
    membersState.filterText = e.target.value.trim();
    membersState.page = 0;
    renderMembersPage();
  });

  // Wire sort handlers
  document.querySelectorAll("#members-list-table thead th[data-mcol]").forEach((th) => {
    th.addEventListener("click", () => {
      const col = parseInt(th.dataset.mcol, 10);
      if (membersState.sortCol === col) membersState.sortDir *= -1;
      else { membersState.sortCol = col; membersState.sortDir = 1; }
      renderMembersPage();
    });
  });

  async function openUserDetail(userId, userCache) {
    show("userDetail");
    $("user-detail-name").textContent = (userCache && userCache.displayName) || "Loading…";
    $("user-detail-mail").textContent = "";
    $("user-detail-jobtitle").textContent = "";
    $("user-account-state").textContent = "";
    $("user-account-state").className = "account-state-badge";
    $("user-groups-tbody").innerHTML = `<tr><td colspan="4" class="loading">Loading…</td></tr>`;
    $("user-groups-count").textContent = "";

    try {
      const [fullUser, userGroups, ownedGroups] = await Promise.all([
        fetchUserBasic(userId),
        GRAPH.getUserMemberOf(userId),
        GRAPH.getUserOwnedGroups(userId).catch(() => []),
        ensureSkuCatalog(), // load license name map (cached after first call)
      ]);
      currentDetailUser = fullUser;
      renderUserDetail(fullUser, userGroups, ownedGroups);
    } catch (err) {
      showError("Failed to load user: " + err.message);
    }
  }

  async function fetchUserBasic(userId) {
    const token = await AUTH.getToken();
    const resp = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}?$select=id,displayName,mail,userPrincipalName,jobTitle,accountEnabled,assignedLicenses`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) throw new Error(`Graph ${resp.status}`);
    return resp.json();
  }

  // SKU catalog cache: skuId -> { partNumber, friendly }
  let skuCatalog = null;
  async function ensureSkuCatalog() {
    if (skuCatalog) return skuCatalog;
    try {
      const skus = await GRAPH.getSubscribedSkus();
      skuCatalog = new Map();
      for (const s of skus) {
        skuCatalog.set(s.skuId, {
          partNumber: s.skuPartNumber,
          friendly: prettifySkuName(s.skuPartNumber, s.skuId),
        });
      }
    } catch (err) {
      console.warn("Could not fetch SKU catalog:", err);
      skuCatalog = new Map();
    }
    return skuCatalog;
  }

  // Best-effort SKU name humanizer. M365 SKU part numbers are like SPB, ENTERPRISEPACK, etc.
  // Common ones in our tenant are mapped explicitly; others get a generic prettify.
  function prettifySkuName(partNumber, skuId) {
    const KNOWN = {
      "O365_BUSINESS_ESSENTIALS": "Microsoft 365 Business Basic",
      "O365_BUSINESS_PREMIUM": "Microsoft 365 Business Standard",
      "SPB": "Microsoft 365 Business Premium",
      "ENTERPRISEPACK": "Office 365 E3",
      "ENTERPRISEPREMIUM": "Office 365 E5",
      "EXCHANGESTANDARD": "Exchange Online (Plan 1)",
      "EXCHANGEENTERPRISE": "Exchange Online (Plan 2)",
      "TEAMS_EXPLORATORY": "Teams Exploratory",
      "FLOW_FREE": "Power Automate Free",
      "POWERAUTOMATE_ATTENDED_RPA": "Power Automate Premium",
      "POWER_BI_STANDARD": "Power BI (Free)",
      "POWER_BI_PRO": "Power BI Pro",
      "EMS": "Enterprise Mobility + Security E3",
      "AAD_PREMIUM": "Entra ID P1",
      "AAD_PREMIUM_P2": "Entra ID P2",
      "WIN_DEF_ATP": "Defender for Endpoint",
      "MCOMEETADV": "Audio Conferencing",
      "DESKLESSPACK": "Office 365 F3",
      "SHAREPOINTSTANDARD": "SharePoint (Plan 1)",
      "SHAREPOINTENTERPRISE": "SharePoint (Plan 2)",
    };
    if (KNOWN[partNumber]) return KNOWN[partNumber];
    // Fallback: convert SOMETHING_LIKE_THIS to "Something Like This"
    return (partNumber || skuId || "Unknown SKU")
      .split(/[_\s]+/)
      .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
      .join(" ");
  }

  function renderUserDetail(user, userGroups, ownedGroups) {
    // Display name with inline edit
    const nameEl = $("user-detail-name");
    nameEl.innerHTML = `<span class="editable-text" data-field="displayName">${escapeHtml(user.displayName || "(no name)")}</span> <button class="btn-edit-inline" data-field="displayName" data-current="${escapeHtml(user.displayName || "")}" aria-label="Edit display name">✎</button>`;

    $("user-detail-mail").innerHTML = user.mail ? `<a href="mailto:${escapeHtml(user.mail)}">${escapeHtml(user.mail)}</a> · UPN: ${escapeHtml(user.userPrincipalName || "")}` : escapeHtml(user.userPrincipalName || "");

    // Job title with inline edit
    const jtText = user.jobTitle || "(no jobTitle set)";
    $("user-detail-jobtitle").innerHTML = `<span class="muted-label">Role:</span> <span class="editable-text" data-field="jobTitle">${escapeHtml(jtText)}</span> <button class="btn-edit-inline" data-field="jobTitle" data-current="${escapeHtml(user.jobTitle || "")}" aria-label="Edit job title">✎</button>`;

    // Wire edit buttons (after innerHTML write)
    $("user-detail-view").querySelectorAll(".btn-edit-inline").forEach((btn) => {
      btn.addEventListener("click", () => startInlineEdit(btn));
    });

    const stateBadge = $("user-account-state");
    const isDisabled = user.accountEnabled === false;
    const hasLicense = (user.assignedLicenses || []).length > 0;
    const isFullyOffboarded = isDisabled && !hasLicense;
    const isPreservedData = isDisabled && hasLicense;

    if (isFullyOffboarded) {
      stateBadge.textContent = "Fully offboarded";
      stateBadge.className = "account-state-badge state-offboarded";
    } else if (isPreservedData) {
      stateBadge.textContent = "Disabled · Data preserved";
      stateBadge.className = "account-state-badge state-preserved";
    } else {
      stateBadge.textContent = "Active";
      stateBadge.className = "account-state-badge state-active";
    }

    // Button visibility by state:
    //  Active  → Offboard shown, Re-enable hidden
    //  Preserved (disabled with license) → Offboard + Re-enable both shown
    //  Fully offboarded → Re-enable only (Offboard is already done)
    $("offboard-user-btn").classList.toggle("hidden", isFullyOffboarded);
    $("reenable-user-btn").classList.toggle("hidden", !isDisabled);

    // Add-to-group only meaningful if account is active (otherwise the user can't use the group)
    $("add-user-to-group-btn").classList.toggle("hidden", isDisabled);

    // Show a state-specific info banner so admins immediately understand what's happening
    renderUserStateBanner({ isFullyOffboarded, isPreservedData });

    // Render assigned licenses (best-effort: needs the SKU catalog)
    renderUserLicenses(user.assignedLicenses || []);

    // Build a union of managed groups the user is a member of OR an owner of.
    const managedIds = new Set(state.groups.map((g) => g.id));
    const memberIds = new Set((userGroups || []).filter((g) => managedIds.has(g.id)).map((g) => g.id));
    const ownerIds = new Set((ownedGroups || []).filter((g) => managedIds.has(g.id)).map((g) => g.id));
    const allRelevantIds = new Set([...memberIds, ...ownerIds]);

    // Group object lookup (combine member-set and owner-set entries since group records overlap)
    const groupLookup = new Map();
    (userGroups || []).forEach((g) => groupLookup.set(g.id, g));
    (ownedGroups || []).forEach((g) => { if (!groupLookup.has(g.id)) groupLookup.set(g.id, g); });

    const relevant = Array.from(allRelevantIds)
      .map((id) => groupLookup.get(id))
      .filter(Boolean)
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));

    $("user-groups-count").textContent = relevant.length;

    const tbody = $("user-groups-tbody");
    if (!relevant.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="muted">Not a member or owner of any managed EVAA/Fusion groups.</td></tr>`;
      return;
    }
    tbody.innerHTML = relevant.map((g) => {
      const isMember = memberIds.has(g.id);
      const isOwner = ownerIds.has(g.id);
      let roleCell;
      if (isMember && isOwner) {
        roleCell = `<span class="role-badge role-both">Member + Owner</span>`;
      } else if (isOwner) {
        roleCell = `<span class="role-badge role-owner">Owner</span>`;
      } else {
        roleCell = `<span class="role-badge role-member">Member</span>`;
      }
      return `<tr>
        <td><button class="link-button" data-jump-group-id="${escapeHtml(g.id)}">${escapeHtml(g.displayName)}</button></td>
        <td>${roleCell}</td>
        <td>${g.mail ? `<a href="mailto:${escapeHtml(g.mail)}" onclick="event.stopPropagation()">${escapeHtml(g.mail)}</a>` : `<span class="muted">—</span>`}</td>
        <td class="row-actions"><button class="btn-remove" data-group-id="${escapeHtml(g.id)}" data-group-name="${escapeHtml(g.displayName)}" data-is-member="${isMember}" data-is-owner="${isOwner}" aria-label="Remove from group">×</button></td>
      </tr>`;
    }).join("");

    tbody.querySelectorAll(".btn-remove").forEach((btn) => {
      btn.addEventListener("click", () => removeUserFromGroup(btn.dataset.groupId, btn.dataset.groupName, btn));
    });
    // Cross-nav: click group name on the user detail page jumps to that group's detail
    tbody.querySelectorAll(".link-button[data-jump-group-id]").forEach((b) => {
      b.addEventListener("click", () => jumpToGroup(b.dataset.jumpGroupId));
    });
  }

  // State-specific banner under the user header. Explains what's true + what's possible
  // so admins immediately understand a non-active user's status.
  function renderUserStateBanner({ isFullyOffboarded, isPreservedData }) {
    const banner = $("user-state-banner");
    if (isFullyOffboarded) {
      banner.className = "user-state-banner banner-offboarded";
      banner.innerHTML = `
        <strong>This user has been fully offboarded.</strong>
        Account is disabled, EVAA license is removed, and they are not in any managed groups.
        Exchange will permanently delete the mailbox and OneDrive within <strong>30 days</strong> of when the license was removed (unless retention policies were configured — they aren't on this tenant).
        <br><br>
        <em>To recover within the 30-day window:</em> click <strong>Re-enable &amp; re-license</strong> below.
        After 30 days, the user object remains but their email + files are unrecoverable from this UI.
        <br><br>
        <span class="muted">Exact days-remaining isn't shown here because Graph doesn't expose the license-removal timestamp on the user object. Check Entra audit logs if you need the precise date.</span>
      `;
      banner.classList.remove("hidden");
    } else if (isPreservedData) {
      banner.className = "user-state-banner banner-preserved";
      banner.innerHTML = `
        <strong>This user is disabled but data is preserved.</strong>
        Account can't sign in. License is still assigned, so mailbox + OneDrive are intact (~$3/mo cost continues).
        Portal admins were notified by email at the time this happened.
        <br><br>
        <em>Options:</em> <strong>Re-enable &amp; re-license</strong> to restore access, or <strong>Offboard</strong> to remove the license (starts the 30-day data clock).
      `;
      banner.classList.remove("hidden");
    } else {
      banner.classList.add("hidden");
      banner.innerHTML = "";
    }
  }

  function renderUserLicenses(assignedLicenses) {
    const el = $("user-detail-licenses");
    if (!assignedLicenses.length) {
      el.innerHTML = `<span class="muted-label">Licenses:</span> <span class="license-badge license-none">none assigned</span>`;
      return;
    }
    const items = assignedLicenses.map((lic) => {
      const cat = skuCatalog ? skuCatalog.get(lic.skuId) : null;
      const friendly = cat ? cat.friendly : `SKU ${lic.skuId.slice(0, 8)}…`;
      const isEvaa = lic.skuId === GRAPH.EVAA_LICENSE_SKU_ID;
      const cls = isEvaa ? "license-badge license-evaa" : "license-badge";
      return `<span class="${cls}" title="${escapeHtml(lic.skuId)}">${escapeHtml(friendly)}</span>`;
    }).join(" ");
    el.innerHTML = `<span class="muted-label">Licenses:</span> ${items}`;
  }

  function jumpToGroup(groupId) {
    activeTab = "groups";
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === "groups"));
    openGroupDetail(groupId);
  }

  function jumpToUser(userId, userCache) {
    activeTab = "members";
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === "members"));
    ensureMembersLoaded(); // load in background so back-to-members lands on a populated table
    openUserDetail(userId, userCache);
  }

  async function removeUserFromGroup(groupId, groupName, btn) {
    if (!currentDetailUser) return;
    const isMember = btn.dataset.isMember === "true";
    const isOwner = btn.dataset.isOwner === "true";
    let roleDesc;
    if (isMember && isOwner) roleDesc = "as both member and director";
    else if (isOwner) roleDesc = "as director (owner)";
    else roleDesc = "as member";

    const ok = await confirmCustom({
      body: `<p>Remove <strong>${escapeHtml(currentDetailUser.displayName)}</strong> from <strong>${escapeHtml(groupName)}</strong> ${escapeHtml(roleDesc)}?</p>
        <p class="muted">This removes only this group association &mdash; the user's account, license, and other groups are untouched.</p>`,
      okLabel: "Remove from group",
      okClass: "btn-warning",
    });
    if (!ok) return;

    btn.disabled = true; btn.textContent = "…";
    const errors = [];
    try {
      if (isMember) {
        try { await GRAPH.removeMember(groupId, currentDetailUser.id); }
        catch (err) { errors.push(`member: ${err.message}`); }
      }
      if (isOwner) {
        try { await GRAPH.removeOwner(groupId, currentDetailUser.id); }
        catch (err) { errors.push(`owner: ${err.message}`); }
      }
      logAction("removed user from group (members view)", currentDetailUser.displayName, currentDetailUser.id, { group: groupName, asMember: isMember, asOwner: isOwner });
      if (errors.length) showError(`Partial remove for ${currentDetailUser.displayName}: ${errors.join("; ")}`);
      await refreshUserDetail();
    } catch (err) {
      btn.disabled = false; btn.textContent = "×";
      showError(`Failed: ${err.message}`);
    }
  }

  async function refreshUserDetail() {
    if (!currentDetailUser) return;
    const [groups, ownedGroups, fullUser] = await Promise.all([
      GRAPH.getUserMemberOf(currentDetailUser.id),
      GRAPH.getUserOwnedGroups(currentDetailUser.id).catch(() => []),
      fetchUserBasic(currentDetailUser.id),
    ]);
    currentDetailUser = fullUser;
    renderUserDetail(fullUser, groups, ownedGroups);
  }

  // Inline edit handler — used for displayName and jobTitle on the user detail view.
  // Replaces the editable-text + edit button with an input + Save / Cancel.
  function startInlineEdit(editBtn) {
    if (!currentDetailUser) return;
    const field = editBtn.dataset.field;
    const current = editBtn.dataset.current || "";
    const container = editBtn.parentElement;
    const editableSpan = container.querySelector(`.editable-text[data-field="${field}"]`);
    if (!editableSpan) return;

    const originalHtml = container.innerHTML;
    const inputId = `inline-edit-${field}`;
    container.innerHTML = `
      <input type="text" id="${inputId}" class="inline-edit-input" value="${escapeHtml(current)}" />
      <button class="btn-secondary btn-inline-save">Save</button>
      <button class="btn-link btn-inline-cancel">Cancel</button>
    `;
    const input = document.getElementById(inputId);
    input.focus();
    input.select();

    function cancel() { container.innerHTML = originalHtml; rewireUserDetailEdits(); }
    container.querySelector(".btn-inline-cancel").addEventListener("click", cancel);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Escape") cancel();
      else if (e.key === "Enter") save();
    });

    async function save() {
      const newValue = input.value.trim();
      if (newValue === current) { cancel(); return; }
      container.querySelector(".btn-inline-save").disabled = true;
      container.querySelector(".btn-inline-cancel").disabled = true;
      try {
        await GRAPH.updateUser(currentDetailUser.id, { [field]: newValue });
        logAction(`edited ${field}`, currentDetailUser.displayName, currentDetailUser.id, { field, from: current, to: newValue });
        await refreshUserDetail();
      } catch (err) {
        container.innerHTML = originalHtml;
        rewireUserDetailEdits();
        showError(`Failed to update ${field}: ${err.message}`);
      }
    }
    container.querySelector(".btn-inline-save").addEventListener("click", save);
  }

  // Re-wire after a cancel (rebinds the edit buttons since we replaced innerHTML).
  function rewireUserDetailEdits() {
    $("user-detail-view").querySelectorAll(".btn-edit-inline").forEach((btn) => {
      // remove old listeners by cloning, then attach fresh
      const fresh = btn.cloneNode(true);
      btn.parentNode.replaceChild(fresh, btn);
      fresh.addEventListener("click", () => startInlineEdit(fresh));
    });
  }

  $("back-to-members-btn").addEventListener("click", () => {
    currentDetailUser = null;
    show("members");
    ensureMembersLoaded();
  });

  // Add to group (from user detail)
  $("add-user-to-group-btn").addEventListener("click", () => {
    if (!currentDetailUser) return;
    $("add-to-group-title").textContent = `Add ${currentDetailUser.displayName} to group`;
    $("add-to-group-filter").value = "";
    renderAddToGroupResults("");
    $("add-to-group-panel").classList.remove("hidden");
    $("add-to-group-filter").focus();
  });
  $("add-to-group-close").addEventListener("click", () => $("add-to-group-panel").classList.add("hidden"));
  $("add-to-group-filter").addEventListener("input", (e) => renderAddToGroupResults(e.target.value));

  function renderAddToGroupResults(filter) {
    const f = (filter || "").toLowerCase();
    const filtered = state.groups
      .filter((g) => (g.displayName || "").toLowerCase().includes(f))
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""));
    const container = $("add-to-group-results");
    if (!filtered.length) { container.innerHTML = `<p class="muted">No groups match.</p>`; return; }
    container.innerHTML = filtered.map((g) => `<button class="user-result" data-group-id="${escapeHtml(g.id)}" data-group-name="${escapeHtml(g.displayName)}">
      <span class="user-name">${escapeHtml(g.displayName)}</span>
      <span class="user-mail muted">${escapeHtml(g.mail || "")}</span>
    </button>`).join("");
    container.querySelectorAll(".user-result").forEach((btn) => {
      btn.addEventListener("click", () => addUserToPickedGroup(btn.dataset.groupId, btn.dataset.groupName));
    });
  }

  async function addUserToPickedGroup(groupId, groupName) {
    if (!currentDetailUser) return;
    const ok = await confirmCustom({
      body: `<p>Add <strong>${escapeHtml(currentDetailUser.displayName)}</strong> to <strong>${escapeHtml(groupName)}</strong> as a member?</p>`,
      okLabel: "Add as member",
      okClass: "btn-primary",
    });
    if (!ok) return;
    try {
      await GRAPH.addMember(groupId, currentDetailUser.id);
      logAction("added user to group (members view)", currentDetailUser.displayName, currentDetailUser.id, { group: groupName });
      $("add-to-group-panel").classList.add("hidden");
      await refreshUserDetail();
    } catch (err) {
      showError(`Failed to add to group: ${err.message}`);
    }
  }

  // Re-enable & re-license a disabled user — counterpart to Offboard fully.
  // Re-enables the account, assigns the EVAA license back. Groups are admin-added manually.
  $("reenable-user-btn").addEventListener("click", async () => {
    if (!currentDetailUser) return;
    const u = currentDetailUser;
    const ok = await confirmCustom({
      body: `<p>Re-enable <strong>${escapeHtml(u.displayName)}</strong> and reassign the EVAA license?</p>
        <p>This will:</p>
        <ul>
          <li>Set <code>accountEnabled = true</code> (user can sign in again)</li>
          <li>Reassign the EVAA M365 license (mailbox / OneDrive reactivated)</li>
        </ul>
        <p class="muted">Groups are NOT re-added automatically &mdash; use <strong>+ Add to group</strong> afterward for any groups they should rejoin. If they don't remember their password they'll need to reset via the M365 &ldquo;Forgot password&rdquo; link (or wait for Password Reset to ship).</p>`,
      okLabel: "Re-enable & re-license",
      okClass: "btn-primary",
    });
    if (!ok) return;

    const btn = $("reenable-user-btn");
    btn.disabled = true;
    const origText = btn.textContent;
    btn.textContent = "Working…";
    const errors = [];
    try {
      try { await GRAPH.enableUserAccount(u.id); }
      catch (err) { errors.push(`enable: ${err.message}`); }
      try { await GRAPH.assignUserLicense(u.id); }
      catch (err) { errors.push(`license: ${err.message}`); }
      logAction("re-enabled & re-licensed user", u.displayName, u.id);
      if (errors.length) showError(`Re-enable partial: ${errors.join("; ")}`);
      await refreshUserDetail();
    } finally {
      btn.disabled = false;
      btn.textContent = origText;
    }
  });

  // Offboard user (from user detail view) — opens the same 3-option modal as the
  // group-detail × button, but configured for user context (no "this group only" option).
  $("offboard-user-btn").addEventListener("click", async () => {
    if (!currentDetailUser) return;
    const u = currentDetailUser;
    let managedGroups = [];
    try {
      const all = await GRAPH.getUserMemberOf(u.id);
      const managedIds = new Set(state.groups.map((g) => g.id));
      managedGroups = all.filter((g) => managedIds.has(g.id));
    } catch (err) {
      console.warn("Could not fetch user's groups for offboard modal:", err);
    }

    pendingRemove = {
      context: "user",
      userId: u.id,
      userName: u.displayName,
      role: "member",
      btn: null,
      otherManagedGroups: managedGroups, // ALL their managed groups (no "current group" concept)
    };

    $("remove-panel-title").textContent = `Remove ${u.displayName}`;
    let body = `<strong>${escapeHtml(u.displayName)}</strong> is in <strong>${managedGroups.length}</strong> managed group${managedGroups.length === 1 ? "" : "s"}.`;
    if (managedGroups.length > 0) body += " The actions below apply to all of them:";
    $("remove-panel-body").innerHTML = body;
    if (managedGroups.length > 0) {
      $("remove-other-groups").innerHTML = managedGroups.map((g) => `<li>${escapeHtml(g.displayName)}</li>`).join("");
      $("remove-other-groups").classList.remove("hidden");
    } else {
      $("remove-other-groups").innerHTML = "";
      $("remove-other-groups").classList.add("hidden");
    }
    // Hide "Remove from this group only" — doesn't apply in user context
    $("remove-this-group-btn").style.display = "none";
    $("remove-panel").classList.remove("hidden");
  });

  // =====================================================================
  // CREATE NEW USER — shared modal, invoked from both Add Member modal and Members view
  // =====================================================================

  // Email body sent to portaladmins@ when an admin picks "Disable account only (preserve data)".
  // Gives the admin group context + explicit next-step options.
  function buildPreserveDataAdminEmailHtml({ adminName, adminMail, userName, userId, groupContextName, removedGroupsHtml }) {
    const adminNameSafe = escapeHtml(adminName);
    const adminMailSafe = escapeHtml(adminMail);
    const userNameSafe = escapeHtml(userName);
    const userIdSafe = escapeHtml(userId);
    const groupContextNameSafe = escapeHtml(groupContextName);
    return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f8fc;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#222;">
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f5f8fc;padding:24px 0;">
  <tr><td align="center">
    <table role="presentation" cellpadding="0" cellspacing="0" width="600" style="background:#ffffff;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.05);overflow:hidden;">
      <tr><td style="background:#856404;color:#fff;padding:18px 28px;">
        <h1 style="margin:0;font-size:18px;font-weight:700;">EVAA Admin Notification</h1>
        <p style="margin:4px 0 0;font-size:13px;opacity:0.9;">User disabled — data preserved (license retained)</p>
      </td></tr>
      <tr><td style="padding:22px 28px;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;">Hi portal admins,</p>
        <p style="margin:0 0 16px;font-size:14px;line-height:1.5;"><strong>${adminNameSafe}</strong> (${adminMailSafe}) just disabled <strong>${userNameSafe}</strong>'s account from the <em>${groupContextNameSafe}</em> context. The user has been:</p>
        <ul style="margin:0 0 12px;padding-left:22px;font-size:14px;line-height:1.6;">
          <li>Removed from these managed groups:
            <ul style="margin-top:4px;">${removedGroupsHtml}</ul>
          </li>
          <li>Account disabled in Entra (can't sign in)</li>
          <li><strong>License retained</strong> — mailbox and OneDrive are preserved indefinitely</li>
        </ul>

        <h3 style="color:#1B4F8C;font-size:14px;margin:18px 0 6px;">Why this matters</h3>
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;">Keeping the license assigned (~$3/mo) means we can recover the user's data later. With <em>Offboard fully</em>, Exchange would have started a 30-day countdown to permanently delete the mailbox + OneDrive.</p>

        <h3 style="color:#1B4F8C;font-size:14px;margin:18px 0 6px;">Options for the group to decide</h3>
        <ol style="margin:0 0 12px;padding-left:22px;font-size:14px;line-height:1.6;">
          <li><strong>Leave as-is</strong> — data preserved indefinitely, license continues. Good for departures where their data may be needed (e.g. financial records, ongoing correspondence).</li>
          <li><strong>Re-enable later</strong> — open the admin portal, find this user, re-enable the account, re-add to groups as needed. Their mailbox / OneDrive come back exactly as left.</li>
          <li><strong>Offboard fully later</strong> — open the admin portal, find this user, click × on any remaining group (or just on the user) and pick <em>Offboard fully</em>. License removed, 30-day deletion clock starts.</li>
        </ol>

        <p style="margin:16px 0 0;font-size:14px;line-height:1.5;">Admin portal: <a href="https://markmeevaa.github.io/directory/admin/" style="color:#1B4F8C;">markmeevaa.github.io/directory/admin/</a></p>
        <p style="margin:6px 0 0;font-size:11px;color:#888;">Acted by: ${adminNameSafe} &lt;${adminMailSafe}&gt; · User ID: ${userIdSafe}</p>
      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>`;
  }

  // Welcome email body sent to a newly-created user's PersonalEmail.
  // Plain HTML, EVAA blue accents, mobile-friendly inline styles.
  function buildWelcomeEmailHtml({ first, last, upn, password, groupName, jobTitle, adminMail }) {
    const groupNameSafe = escapeHtml(groupName);
    const upnSafe = escapeHtml(upn);
    const pwdSafe = escapeHtml(password);
    const firstSafe = escapeHtml(first);
    const jobTitleSafe = escapeHtml(jobTitle);
    const adminMailSafe = escapeHtml(adminMail);
    return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f8fc;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#222;">
<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f5f8fc;padding:24px 0;">
  <tr><td align="center">
    <table role="presentation" cellpadding="0" cellspacing="0" width="600" style="background:#ffffff;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.05);overflow:hidden;">
      <tr><td style="background:#1B4F8C;color:#fff;padding:20px 28px;">
        <h1 style="margin:0;font-size:20px;font-weight:700;">Welcome to EVAA</h1>
        <p style="margin:4px 0 0;font-size:14px;color:#cfd8e6;">Eastview Athletic Association</p>
      </td></tr>
      <tr><td style="padding:24px 28px;">
        <p style="margin:0 0 12px;font-size:15px;line-height:1.5;">Hi ${firstSafe},</p>
        <p style="margin:0 0 16px;font-size:15px;line-height:1.5;">You've been added to <strong>${groupNameSafe}</strong> as <strong>${jobTitleSafe}</strong>. A new EVAA Microsoft 365 account has been created for you. Here are your sign-in credentials:</p>

        <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;background:#f5f8fc;border:1px solid #d9e1ec;border-radius:6px;margin:12px 0;">
          <tr><td style="padding:14px 16px;">
            <div style="font-size:14px;margin-bottom:6px;"><strong>Sign-in address:</strong></div>
            <div style="font-family:Menlo,Consolas,monospace;font-size:15px;color:#1B4F8C;margin-bottom:14px;">${upnSafe}</div>
            <div style="font-size:14px;margin-bottom:6px;"><strong>Temporary password:</strong></div>
            <div style="font-family:Menlo,Consolas,monospace;font-size:15px;color:#1B4F8C;">${pwdSafe}</div>
          </td></tr>
        </table>

        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;"><strong>Important:</strong> you'll be prompted to change this password the first time you sign in. Pick something memorable but secure.</p>

        <h3 style="color:#1B4F8C;font-size:15px;margin:20px 0 8px;">How to sign in</h3>
        <ol style="margin:0 0 12px;padding-left:22px;font-size:14px;line-height:1.6;">
          <li>Go to <a href="https://www.office.com" style="color:#1B4F8C;">office.com</a> and click <strong>Sign in</strong>.</li>
          <li>Enter the sign-in address above.</li>
          <li>Enter the temporary password.</li>
          <li>Follow the prompts to set a new password.</li>
          <li>From there you'll have access to Outlook (your new EVAA email), Word, Excel, Teams, etc.</li>
        </ol>

        <h3 style="color:#1B4F8C;font-size:15px;margin:20px 0 8px;">Help</h3>
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;">If you have trouble signing in or need anything changed, reply to this email or contact <a href="mailto:${adminMailSafe}" style="color:#1B4F8C;">${adminMailSafe}</a>.</p>
        <p style="margin:0;font-size:14px;line-height:1.5;">Thanks for being part of EVAA!</p>
      </td></tr>
      <tr><td style="background:#f5f8fc;color:#888;padding:14px 28px;font-size:12px;text-align:center;border-top:1px solid #d9e1ec;">
        Eastview Athletic Association · <a href="https://www.evaasports.org" style="color:#1B4F8C;">evaasports.org</a>
      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>`;
  }

  // Generate a secure-ish random temp password meeting M365 complexity requirements.
  function generateTempPassword() {
    const upper = "ABCDEFGHJKMNPQRSTUVWXYZ";
    const lower = "abcdefghjkmnpqrstuvwxyz";
    const digit = "23456789";
    const sym = "!@#$%&*";
    const all = upper + lower + digit + sym;
    const rand = (s) => s[Math.floor(Math.random() * s.length)];
    const chars = [rand(upper), rand(lower), rand(digit), rand(sym)];
    for (let i = chars.length; i < 14; i++) chars.push(rand(all));
    // shuffle
    for (let i = chars.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [chars[i], chars[j]] = [chars[j], chars[i]];
    }
    return chars.join("");
  }

  // Source-context: which group should the new user be added to by default?
  let createUserContextGroupId = null;

  function openCreateUserPanel(presetGroupId) {
    createUserContextGroupId = presetGroupId || null;
    // Reset form
    $("create-user-form").reset();
    $("cu-domain").value = "evaasports.org";
    $("cu-upn").value = "";
    $("cu-jobtitle").value = "";
    $("cu-role-other-wrap").classList.add("hidden");
    upnDirty = false;
    jobTitleDirty = false;
    // Populate group dropdown
    const sel = $("cu-group");
    sel.innerHTML = state.groups
      .slice()
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || ""))
      .map((g) => `<option value="${escapeHtml(g.id)}">${escapeHtml(g.displayName)}</option>`)
      .join("");
    if (presetGroupId) sel.value = presetGroupId;
    $("create-user-progress").classList.add("hidden");
    $("create-user-result").classList.add("hidden");
    $("create-user-form").classList.remove("hidden");
    $("create-user-panel").classList.remove("hidden");
    $("cu-first").focus();
  }
  function closeCreateUserPanel() {
    createUserContextGroupId = null;
    $("create-user-panel").classList.add("hidden");
  }
  $("create-user-close").addEventListener("click", closeCreateUserPanel);
  $("create-user-cancel").addEventListener("click", closeCreateUserPanel);

  // Hook into both entry points
  $("create-user-from-add-btn").addEventListener("click", () => {
    closeAddPanel();
    openCreateUserPanel(currentDetailGroup ? currentDetailGroup.id : null);
  });
  $("create-user-from-members-btn").addEventListener("click", () => openCreateUserPanel(null));

  // Auto-fill UPN from first/last/domain unless user manually edited
  let upnDirty = false;
  $("cu-upn").addEventListener("input", () => { upnDirty = true; });
  function recomputeUpn() {
    if (upnDirty) return;
    const first = $("cu-first").value.trim().toLowerCase().replace(/[^a-z]/g, "");
    const last = $("cu-last").value.trim().toLowerCase().replace(/[^a-z]/g, "");
    const domain = $("cu-domain").value;
    if (first && last) $("cu-upn").value = `${first.charAt(0)}${last}@${domain}`;
  }
  $("cu-first").addEventListener("input", recomputeUpn);
  $("cu-last").addEventListener("input", recomputeUpn);
  $("cu-domain").addEventListener("change", recomputeUpn);

  // Auto-build jobTitle from selected group + role (unless user manually edited).
  // Uses the production "Sport: Role" convention (e.g. "Baseball: President").
  // Strips common prefixes like "EVAA - " or "Fusion - " from the group name.
  let jobTitleDirty = false;
  $("cu-jobtitle").addEventListener("input", () => { jobTitleDirty = true; });

  function stripGroupPrefix(displayName) {
    if (!displayName) return "";
    return displayName
      .replace(/^EVAA\s*-\s*Financial\s+Aid\s*-\s*/i, "Financial Aid - ")
      .replace(/^EVAA\s*-\s*/i, "")
      .replace(/^Fusion\s*-\s*/i, "Fusion ")
      .trim();
  }

  function recomputeJobTitle() {
    if (jobTitleDirty) return;
    const groupId = $("cu-group").value;
    const group = state.groups.find((g) => g.id === groupId);
    const roleSelect = $("cu-role").value;
    const roleOther = $("cu-role-other").value.trim();
    const role = roleSelect === "Other" ? roleOther : roleSelect;
    if (!group || !role) { $("cu-jobtitle").value = ""; return; }
    const sport = stripGroupPrefix(group.displayName);
    // Leadership group uses bare-role convention (no prefix) per project state
    if (group.displayName === "EVAA - Leadership") {
      $("cu-jobtitle").value = role;
    } else {
      $("cu-jobtitle").value = `${sport}: ${role}`;
    }
  }

  // Toggle the "Other" custom-role input when "Other" is picked
  $("cu-role").addEventListener("change", () => {
    const isOther = $("cu-role").value === "Other";
    $("cu-role-other-wrap").classList.toggle("hidden", !isOther);
    if (!isOther) $("cu-role-other").value = "";
    recomputeJobTitle();
  });
  $("cu-role-other").addEventListener("input", recomputeJobTitle);
  $("cu-group").addEventListener("change", recomputeJobTitle);

  $("create-user-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const first = $("cu-first").value.trim();
    const last = $("cu-last").value.trim();
    const personalEmail = $("cu-personal-email").value.trim();
    const domain = $("cu-domain").value;
    const upn = $("cu-upn").value.trim() || `${first.charAt(0).toLowerCase()}${last.toLowerCase()}@${domain}`;
    const mailNickname = upn.split("@")[0];
    const jobTitle = $("cu-jobtitle").value.trim();
    const groupId = $("cu-group").value;
    const asOwner = $("cu-as-owner").checked;
    const groupName = state.groups.find((g) => g.id === groupId)?.displayName || "";

    // OWNER MODE: file a MemberRequests row for admin approval; do not provision directly.
    if (role === "owner") {
      const roleSelect = $("cu-role").value;
      const roleOther = $("cu-role-other").value.trim();
      const roleVal = roleSelect === "Other" ? roleOther : roleSelect;
      try {
        await GRAPH.createMemberRequest({
          requestType: "Add",
          sportDisplayName: groupName,
          firstName: first,
          lastName: last,
          personalEmail,
          emailDomain: "@" + domain,
          role: roleVal,
        });
        closeCreateUserPanel();
        showToast(`Add request submitted — an admin will approve and provision the account.`);
        logAction("submitted Add request", `${first} ${last}`, null, { group: groupName, role: roleVal });
      } catch (err) {
        showError(`Could not submit Add request: ${err.message}`);
      }
      return;
    }

    const password = generateTempPassword();

    const form = $("create-user-form");
    const progress = $("create-user-progress");
    const result = $("create-user-result");
    form.classList.add("hidden");
    progress.classList.remove("hidden");
    progress.innerHTML = `<p class="loading">Creating ${escapeHtml(first)} ${escapeHtml(last)}…</p><ul id="cu-steps"></ul>`;
    const stepsEl = $("cu-steps");
    const stepLog = (msg, ok) => { stepsEl.insertAdjacentHTML("beforeend", `<li class="${ok ? 'step-ok' : 'step-err'}">${ok ? '✓' : '✗'} ${escapeHtml(msg)}</li>`); };

    let newUserId;
    try {
      const created = await GRAPH.createUser({
        displayName: `${first} ${last}`,
        givenName: first,
        surname: last,
        userPrincipalName: upn,
        mailNickname,
        jobTitle: jobTitle || undefined,
        password,
      });
      newUserId = created.id;
      stepLog(`User created (${upn})`, true);
    } catch (err) {
      stepLog(`Create user failed: ${err.message}`, false);
      // Show retry option by re-enabling form
      progress.classList.add("hidden");
      form.classList.remove("hidden");
      showError(`User creation failed: ${err.message}`);
      return;
    }

    // 2. Assign license (best-effort)
    try {
      await GRAPH.assignUserLicense(newUserId);
      stepLog(`Assigned EVAA license`, true);
    } catch (err) {
      stepLog(`License assign failed: ${err.message}`, false);
    }

    // 3. Add to group
    try {
      if (asOwner) await GRAPH.addOwner(groupId, newUserId);
      await GRAPH.addMember(groupId, newUserId);
      stepLog(`Added to ${groupName}${asOwner ? " (as owner + member)" : ""}`, true);
    } catch (err) {
      stepLog(`Group add failed: ${err.message}`, false);
    }

    // 4. Auto-send welcome email
    let emailSent = false;
    try {
      const subject = `Welcome to EVAA — your new ${domain === "evaasports.org" ? "Eastview Athletic Association" : "Apple Valley Fusion"} account`;
      const adminMail = AUTH.getAccount()?.username || "web-admin@evaasports.org";
      const html = buildWelcomeEmailHtml({ first, last, upn, password, groupName, jobTitle: jobTitle || "(not set)", adminMail });
      await GRAPH.sendMail([personalEmail], subject, html);
      stepLog(`Welcome email sent to ${personalEmail}`, true);
      emailSent = true;
    } catch (err) {
      stepLog(`Welcome email failed: ${err.message}`, false);
    }

    logAction("created user", `${first} ${last}`, newUserId, { upn, group: groupName, asOwner, emailSent });

    // 5. Show result modal
    progress.classList.add("hidden");
    result.classList.remove("hidden");
    const emailNotice = emailSent
      ? `<p class="success-note">✓ Welcome email sent to <strong>${escapeHtml(personalEmail)}</strong>. The credentials below are also shown here in case you need them as a backup.</p>`
      : `<p class="muted">Welcome email failed to send automatically — please email these credentials to <strong>${escapeHtml(personalEmail)}</strong> manually:</p>`;
    result.innerHTML = `<h4>✓ User created</h4>
      ${emailNotice}
      <div class="cred-block">
        <div><strong>Sign-in:</strong> <code>${escapeHtml(upn)}</code></div>
        <div><strong>Temporary password:</strong> <code id="cu-pwd">${escapeHtml(password)}</code> <button id="cu-copy-pwd" class="btn-link">Copy</button></div>
        <div class="muted">User must change password on first sign-in.</div>
      </div>
      <div class="modal-actions">
        <button id="cu-done" class="btn-primary">Done</button>
      </div>`;
    $("cu-copy-pwd").addEventListener("click", () => {
      navigator.clipboard.writeText(password).then(() => {
        $("cu-copy-pwd").textContent = "Copied!";
        setTimeout(() => { const el = document.getElementById("cu-copy-pwd"); if (el) el.textContent = "Copy"; }, 1500);
      });
    });
    $("cu-done").addEventListener("click", () => {
      closeCreateUserPanel();
      // Refresh whichever view is active
      if (activeTab === "groups" && currentDetailGroup) refreshDetail();
      else if (activeTab === "members" && currentDetailUser) refreshUserDetail();
    });
  });
})();
