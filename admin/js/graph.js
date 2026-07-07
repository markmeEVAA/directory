// EVAA Admin Portal — Microsoft Graph API wrapper
// Thin fetch-based client. All calls require a token from AUTH.getToken().

const GRAPH = (() => {
  const BASE = "https://graph.microsoft.com/v1.0";
  // Either of these groups grants full admin access in the portal:
  //   EVAA Portal Admins  — the explicit admins group
  //   EVAA - Leadership   — 7 executive officers (President, VP, Treasurer, Secretary,
  //                         Marketing, Operations, Safety). They get admin too.
  const ADMIN_GROUP_IDS = [
    "98d51c39-149a-4dbf-9e86-1510035d8239", // EVAA Portal Admins
    "12e5f9ce-d644-4052-aff8-b31e99c3acb9", // EVAA - Leadership
  ];

  async function callGraph(path, options = {}) {
    const token = await AUTH.getToken();
    const url = path.startsWith("http") ? path : `${BASE}${path}`;
    const resp = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      throw new Error(`Graph ${resp.status} ${resp.statusText}: ${body.slice(0, 300)}`);
    }
    // 204 No Content and 202 Accepted (e.g., /me/sendMail) return empty bodies.
    // Some other 2xx responses may also be empty. Read as text first; only parse
    // JSON if there's something there.
    if (resp.status === 204 || resp.status === 202) return null;
    const text = await resp.text();
    if (!text) return null;
    try { return JSON.parse(text); }
    catch { return null; }
  }

  // Paginated GET — follows @odata.nextLink until exhausted.
  async function callGraphAll(path) {
    const out = [];
    let next = path;
    while (next) {
      const data = await callGraph(next);
      out.push(...(data.value || []));
      next = data["@odata.nextLink"] || null;
    }
    return out;
  }

  async function getMe() {
    return callGraph("/me?$select=id,displayName,mail,userPrincipalName,jobTitle");
  }

  // True if the current user is a member of any admin-granting group
  // (EVAA Portal Admins OR EVAA - Leadership).
  async function isPortalAdmin() {
    try {
      const result = await callGraph("/me/checkMemberGroups", {
        method: "POST",
        body: JSON.stringify({ groupIds: ADMIN_GROUP_IDS }),
      });
      return Array.isArray(result.value) && result.value.some((id) => ADMIN_GROUP_IDS.includes(id));
    } catch (err) {
      console.error("Portal-admin check failed:", err);
      return false;
    }
  }

  // All EVAA/Fusion Unified groups (matches the existing helper-flow filter).
  // Note: no $orderby — Graph requires ConsistencyLevel:eventual header to combine
  // $filter+$orderby in advanced-query mode, and the JS already sorts client-side.
  async function listManagedGroups() {
    const filter = "(startswith(displayName,'EVAA') or startswith(displayName,'Fusion')) and groupTypes/any(c:c eq 'Unified')";
    const select = "id,displayName,mail,groupTypes,description";
    const path = `/groups?$filter=${encodeURIComponent(filter)}&$select=${select}&$top=100`;
    return callGraphAll(path);
  }

  async function listGroupMembers(groupId) {
    // accountEnabled is selected so the UI can badge disabled (already-offboarded)
    // members and block a duplicate removal request against them.
    const path = `/groups/${groupId}/members/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName,jobTitle,accountEnabled&$top=100`;
    return callGraphAll(path);
  }

  async function listGroupOwners(groupId) {
    const path = `/groups/${groupId}/owners/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`;
    return callGraphAll(path);
  }

  // List all users in the EVAA + Fusion domains (both active AND disabled).
  // Disabled accounts are surfaced so admins can re-enable them within the recovery window.
  // Used by the Manage Members tab for the paginated all-users view.
  async function listAllManagedUsers() {
    const filter = "endswith(userPrincipalName,'@evaasports.org') or endswith(userPrincipalName,'@avfusion.org')";
    const select = "id,displayName,mail,userPrincipalName,jobTitle,accountEnabled";
    // endswith requires advanced query, which requires ConsistencyLevel: eventual + $count
    const path = `/users?$filter=${encodeURIComponent(filter)}&$select=${select}&$count=true&$top=999`;
    const token = await AUTH.getToken();
    const resp = await fetch(`${BASE}${path}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ConsistencyLevel: "eventual",
      },
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      throw new Error(`Graph ${resp.status}: ${body.slice(0, 300)}`);
    }
    const data = await resp.json();
    return data.value || [];
  }

  // User search (for picking someone to add as owner/member).
  // Filters to enabled users in @evaasports.org and @avfusion.org.
  async function searchUsers(query) {
    if (!query || query.length < 2) return [];
    const safe = query.replace(/'/g, "''");
    const filter = `accountEnabled eq true and (startswith(displayName,'${safe}') or startswith(mail,'${safe}') or startswith(userPrincipalName,'${safe}'))`;
    const select = "id,displayName,mail,userPrincipalName,jobTitle";
    const path = `/users?$filter=${encodeURIComponent(filter)}&$select=${select}&$top=15`;
    const r = await callGraph(path);
    return r.value || [];
  }

  // Create a new M365 (Unified) group — mail-enabled, security-disabled. This is the
  // SAME type as the "EVAA - <Sport>" board-director groups (they get a mailbox +
  // SharePoint and can own/send-lock an email list, per EMAIL-LISTS.md). Returns the
  // created group object (id, displayName, mail, mailNickname, …).
  //
  // owners/members are passed inline via @odata.bind so the group is never ownerless,
  // even for an instant (Graph caps each bind array at 20 on create — callers add any
  // overflow afterward via addOwner/addMember). NOTE: a just-created group is subject to
  // replication lag — it is NOT immediately listable via listManagedGroups and a follow-up
  // write (addOwner/removeOwner) may 404 for a few seconds (see EMAIL-LISTS.md gotcha #4).
  // Callers should optimistically add the returned object to their local list and retry
  // any follow-up writes with backoff.
  async function createGroup({ displayName, mailNickname, description, ownerIds = [], memberIds = [] }) {
    const bind = (id) => `https://graph.microsoft.com/v1.0/directoryObjects/${id}`;
    const body = {
      displayName,
      mailNickname,
      groupTypes: ["Unified"],
      mailEnabled: true,
      securityEnabled: false,
    };
    if (description) body.description = description;
    // Dedupe ids; Graph rejects a bind array with duplicate references.
    const owners = [...new Set(ownerIds)].slice(0, 20);
    const members = [...new Set(memberIds)].slice(0, 20);
    if (owners.length) body["owners@odata.bind"] = owners.map(bind);
    if (members.length) body["members@odata.bind"] = members.map(bind);
    return callGraph("/groups", { method: "POST", body: JSON.stringify(body) });
  }

  // The connected SharePoint team-site URL for a Unified group, e.g.
  // "https://evaasports.sharepoint.com/sites/soccertravel". Returns null if the site
  // isn't provisioned yet (a just-created group lags) or the caller lacks access —
  // callers should treat null as "no link to show" rather than an error.
  async function getGroupSiteUrl(groupId) {
    try {
      const site = await callGraph(`/groups/${groupId}/sites/root?$select=webUrl`);
      return site && site.webUrl ? site.webUrl : null;
    } catch {
      return null;
    }
  }

  // Generic group-field update. Used by the inline rename on the group detail view.
  // patch is a partial group object, e.g. { displayName: "EVAA - Travel Soccer" }.
  // Note: this changes the friendly displayName only — it does NOT change the group's
  // mailNickname / primary email address (that's an Exchange operation).
  async function updateGroup(groupId, patch) {
    return callGraph(`/groups/${groupId}`, {
      method: "PATCH",
      body: JSON.stringify(patch),
    });
  }

  // Owner ops
  async function addOwner(groupId, userId) {
    return callGraph(`/groups/${groupId}/owners/$ref`, {
      method: "POST",
      body: JSON.stringify({
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
      }),
    });
  }
  async function removeOwner(groupId, userId) {
    return callGraph(`/groups/${groupId}/owners/${userId}/$ref`, { method: "DELETE" });
  }

  // Member ops
  async function addMember(groupId, userId) {
    return callGraph(`/groups/${groupId}/members/$ref`, {
      method: "POST",
      body: JSON.stringify({
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`,
      }),
    });
  }
  async function removeMember(groupId, userId) {
    return callGraph(`/groups/${groupId}/members/${userId}/$ref`, { method: "DELETE" });
  }

  // Get all groups a user is a direct member of.
  // Used to decide single-group remove vs full offboard for the smart prompt.
  async function getUserMemberOf(userId) {
    const path = `/users/${userId}/memberOf/microsoft.graph.group?$select=id,displayName,mail,groupTypes&$top=200`;
    return callGraphAll(path);
  }

  // Get all groups owned by this user (returns group objects).
  // Used to show "Member" vs "Director" vs "Both" on the user detail view.
  async function getUserOwnedGroups(userId) {
    const path = `/users/${userId}/ownedObjects/microsoft.graph.group?$select=id,displayName,mail&$top=200`;
    return callGraphAll(path);
  }

  // EVAA Standard license SKU (from project state file line 197 / 207).
  const EVAA_LICENSE_SKU_ID = "3b555118-da6a-4418-894f-7df1e2096870";

  // Remove the EVAA license from a user. POST /users/{id}/assignLicense with removeLicenses array.
  async function removeUserLicense(userId, skuId = EVAA_LICENSE_SKU_ID) {
    return callGraph(`/users/${userId}/assignLicense`, {
      method: "POST",
      body: JSON.stringify({
        addLicenses: [],
        removeLicenses: [skuId],
      }),
    });
  }

  async function disableUserAccount(userId) {
    return callGraph(`/users/${userId}`, {
      method: "PATCH",
      body: JSON.stringify({ accountEnabled: false }),
    });
  }

  async function enableUserAccount(userId) {
    return callGraph(`/users/${userId}`, {
      method: "PATCH",
      body: JSON.stringify({ accountEnabled: true }),
    });
  }

  // Permanently delete a user. Graph DELETE is a SOFT delete: the account moves to
  // the Entra deleted-items bin, recoverable for 30 days, then auto-purged. Requires
  // User.ReadWrite.All (in scope) AND the signed-in admin to hold a directory role
  // that permits user deletion (e.g. User Administrator) — otherwise Graph returns
  // 403 Authorization_RequestDenied, handled by the caller.
  async function deleteUser(userId) {
    return callGraph(`/users/${userId}`, { method: "DELETE" });
  }

  // Generic user-field update. Used by the inline-edit UI on user detail view.
  // patch is a partial user object, e.g. { jobTitle: "Baseball: VP" } or { displayName: "..." }
  async function updateUser(userId, patch) {
    return callGraph(`/users/${userId}`, {
      method: "PATCH",
      body: JSON.stringify(patch),
    });
  }

  // Tenant license catalog (SKU id -> friendly name + part number).
  // Used to translate a user's assignedLicenses[].skuId into something human-readable.
  async function getSubscribedSkus() {
    const result = await callGraph("/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits");
    return result.value || [];
  }

  // Create a new user via Graph. Returns the created user object including `id`.
  async function createUser({ displayName, givenName, surname, userPrincipalName, mailNickname, jobTitle, password }) {
    const body = {
      accountEnabled: true,
      displayName,
      givenName,
      surname,
      mailNickname,
      userPrincipalName,
      usageLocation: "US",
      passwordProfile: {
        forceChangePasswordNextSignIn: true,
        password,
      },
    };
    if (jobTitle) body.jobTitle = jobTitle;
    return callGraph("/users", { method: "POST", body: JSON.stringify(body) });
  }

  // Assign the EVAA license to a user.
  async function assignUserLicense(userId, skuId = EVAA_LICENSE_SKU_ID) {
    return callGraph(`/users/${userId}/assignLicense`, {
      method: "POST",
      body: JSON.stringify({
        addLicenses: [{ skuId, disabledPlans: [] }],
        removeLicenses: [],
      }),
    });
  }

  // Send mail as the signed-in admin (delegated Mail.Send).
  // toRecipients: array of email address strings.
  // htmlBody: HTML string (use \n for newlines inside <pre> if needed).
  async function sendMail(toRecipients, subject, htmlBody, opts = {}) {
    const body = {
      message: {
        subject,
        body: { contentType: "HTML", content: htmlBody },
        toRecipients: toRecipients.map((addr) => ({ emailAddress: { address: addr } })),
      },
      saveToSentItems: opts.saveToSentItems !== false,
    };
    if (opts.cc && opts.cc.length) {
      body.message.ccRecipients = opts.cc.map((addr) => ({ emailAddress: { address: addr } }));
    }
    if (opts.bcc && opts.bcc.length) {
      body.message.bccRecipients = opts.bcc.map((addr) => ({ emailAddress: { address: addr } }));
    }
    return callGraph(`/me/sendMail`, {
      method: "POST",
      body: JSON.stringify(body),
    });
  }

  // Return the subset of managed group IDs that the current user OWNS.
  // Uses /me/ownedObjects, filters to the managed-group IDs we already loaded.
  // Empty array means "not an owner of anything managed".
  async function getOwnedManagedGroupIds(managedGroupIds) {
    if (!managedGroupIds || managedGroupIds.length === 0) return [];
    const owned = await callGraphAll(
      `/me/ownedObjects/microsoft.graph.group?$select=id&$top=200`
    );
    const ownedSet = new Set(owned.map((g) => g.id));
    return managedGroupIds.filter((id) => ownedSet.has(id));
  }

  // ---------------- Submit Member Request (owner-mode approval flow) ----------------
  // Owner-mode Add/Remove rows are created via the "EVAA - Submit Member Request"
  // helper flow, which runs as the flow owner (web-admin) and has full SP rights.
  // The user does NOT need direct SP access — this bypasses the per-user permission
  // issue that direct Graph SP writes would hit for non-admin group owners.
  //
  // The flow's HTTP trigger URL is SAS-signed. Anyone with this URL can submit a
  // row, but the row lands in MemberRequests with Status=Pending and an admin
  // still has to Approve it via the existing Approval Flow before any provisioning
  // happens. So it's safe to ship the URL in client-side JS.
  const SUBMIT_MEMBER_REQUEST_URL =
    "https://defaultb5897a1bb85b42bd8e619b021b67d2.ce.environment.api.powerplatform.com:443/" +
    "powerautomate/automations/direct/workflows/3e7e4752b3314da9b462c92cc700c3eb/" +
    "triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&" +
    "sig=zaahy0L624fFCgB0Qmu7sNGV-OQbAahEIpRtnhDg6uY";

  // Submit a row to MemberRequests via the helper flow. Mirrors the canvas-app
  // AddScreen / RemoveConfirmScreen Patch. requestType: "Add" or "Remove".
  // For Add: firstName, lastName, personalEmail, emailDomain ("@evaasports.org" or
  //   "@avfusion.org"), role, sportDisplayName.
  // For Remove: firstName, lastName, memberId, memberEmail, sportDisplayName.
  async function createMemberRequest(payload) {
    const me = await getMe();
    const body = {
      requestType: payload.requestType,
      sportDisplayName: payload.sportDisplayName,
      firstName: payload.firstName || "",
      lastName: payload.lastName || "",
      personalEmail: payload.personalEmail || "n/a",
      emailDomain: payload.emailDomain || "@evaasports.org",
      role: payload.role || "",
      memberId: payload.memberId || "",
      memberEmail: payload.memberEmail || "",
      requesterEmail: me?.userPrincipalName || me?.mail || "",
      requesterName: me?.displayName || "",
      removalDisposition: payload.removalDisposition || "",
      transferTo: payload.transferTo || "",
    };
    const resp = await fetch(SUBMIT_MEMBER_REQUEST_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`Submit-request flow ${resp.status}: ${errText.slice(0, 300)}`);
    }
    return resp.json().catch(() => ({ status: "submitted" }));
  }

  // ---------------- Submit Password Reset Request (owner-mode approval flow) ----------------
  // Owner-mode Reset Password is routed via the "EVAA - Submit Password Reset"
  // helper flow → writes a Pending row to PasswordResetRequests SP list →
  // separate Approval Flow triggers, sends an approval email to portaladmin@,
  // and on Approve generates a new temp password, PATCHes the user via Graph,
  // and emails credentials to the personal address.
  //
  // SAS-signed URL — safe to ship in client JS because the action still requires
  // admin approval before any password change happens.
  const SUBMIT_PASSWORD_RESET_URL =
    "https://defaultb5897a1bb85b42bd8e619b021b67d2.ce.environment.api.powerplatform.com:443/" +
    "powerautomate/automations/direct/workflows/65aa545968584eab8a6e97c3e7aba019/" +
    "triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&" +
    "sig=CArZq1A_udxiYzBdxEgt4XnhHqFcZ6jJ4oT0q7CNQmo";

  async function submitPasswordReset(payload) {
    const me = await getMe();
    const body = {
      requesterEmail: me?.userPrincipalName || me?.mail || "",
      requesterName: me?.displayName || "",
      targetUserId: payload.targetUserId,
      targetUserUpn: payload.targetUserUpn,
      targetUserDisplayName: payload.targetUserDisplayName,
      personalEmail: payload.personalEmail,
    };
    const resp = await fetch(SUBMIT_PASSWORD_RESET_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`Submit-password-reset flow ${resp.status}: ${errText.slice(0, 300)}`);
    }
    return resp.json().catch(() => ({ status: "submitted" }));
  }

  // ---------------- AdminActionLog (Phase 2.4 audit log) ----------------
  // SharePoint list at evaasports.sharepoint.com/sites/EVAABoardPortal/Lists/AdminActionLog
  // Columns (Mark added via SP UI on 5/22/2026):
  //   Title (built-in, auto-summary)
  //   Actor (text — UPN of admin)
  //   ActionType (choice: AddMember/RemoveMember/AddOwner/RemoveOwner/CreateUser/
  //               DisableUser/EnableUser/Offboard/PreserveData/ResetPassword/RestoreUser/Other)
  //   TargetUser (text — UPN or displayName; blank for group-only ops)
  //   TargetGroup (text — group displayName; blank for user-level ops)
  //   Result (choice: Success/Failure/Partial)
  //   ErrorDetail (multi-line text)
  //   Notes (multi-line text)
  const ADMIN_ACTION_LOG_SITE_ID =
    "evaasports.sharepoint.com,5c93dacd-279c-41bd-a4b0-64288b689f69,3c4714c8-a098-4f4b-bdd9-ad7a69c13740";
  const ADMIN_ACTION_LOG_LIST_ID = "5edbe46a-ba12-4cf2-9dfa-937fa47297c4";

  // Map a free-text action string (legacy logAction shape) into one of the
  // SP Choice values. Order matters — more specific patterns checked first.
  function mapActionType(action) {
    const s = (action || "").toLowerCase();
    if (s.includes("created user")) return "CreateUser";
    if (s.includes("disabled user with data preserved")) return "PreserveData";
    if (s.includes("offboarded fully")) return "Offboard";
    if (s.includes("re-enabled") || s.includes("re-licensed")) return "RestoreUser";
    if (s.includes("reset password") || s.includes("password reset")) return "ResetPassword";
    if (s.includes("added owner")) return "AddOwner";
    if (s.includes("removed owner")) return "RemoveOwner";
    if (s.includes("add request")) return "AddMember";
    if (s.includes("remove request")) return "RemoveMember";
    if (s.includes("added") && (s.includes("member") || s.includes("user to group"))) return "AddMember";
    if (s.includes("removed") && (s.includes("member") || s.includes("user from group") || s.includes("per-group"))) return "RemoveMember";
    if (s.includes("disabled")) return "DisableUser";
    if (s.includes("enabled")) return "EnableUser";
    return "Other";
  }

  // Fire-and-forget write to AdminActionLog SP list.
  // Errors are swallowed (logged to console.warn) so admin actions never block on logging.
  async function logAuditEntry({ actor, action, targetName, targetId, targetGroup, result, errorDetail, notes }) {
    try {
      const actionType = mapActionType(action);
      const titleParts = [actionType, "by", actor || "(unknown)"];
      if (targetName) titleParts.push("-", targetName);
      if (targetGroup) titleParts.push("in", targetGroup);
      const title = titleParts.join(" ").slice(0, 250);
      const fields = {
        Title: title,
        Actor: actor || "(unknown)",
        ActionType: actionType,
        TargetUser: targetName || "",
        TargetGroup: targetGroup || "",
        Result: result || "Success",
        ErrorDetail: errorDetail || "",
        Notes: notes || "",
      };
      await callGraph(
        `/sites/${ADMIN_ACTION_LOG_SITE_ID}/lists/${ADMIN_ACTION_LOG_LIST_ID}/items`,
        { method: "POST", body: JSON.stringify({ fields }) }
      );
    } catch (err) {
      console.warn("[AUDIT-FALLBACK]", err.message || err, { action, targetName, targetId, targetGroup });
    }
  }

  // Fetch recent audit log entries (viewer tab).
  // Returns the full Graph response so callers can use @odata.nextLink for pagination.
  // SP $orderby support via Graph is limited; we fetch by most-recently-modified at the
  // list-item level (lastModifiedDateTime). Client sorts/filters further.
  async function listAuditLog({ top = 100, nextLink = null } = {}) {
    if (nextLink) {
      // nextLink is already a full URL with $skiptoken; use raw fetch
      const token = await AUTH.getToken();
      const resp = await fetch(nextLink, {
        headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
      });
      if (!resp.ok) throw new Error(`Audit log fetch ${resp.status}`);
      return resp.json();
    }
    const path = `/sites/${ADMIN_ACTION_LOG_SITE_ID}/lists/${ADMIN_ACTION_LOG_LIST_ID}/items?$expand=fields&$top=${top}&$orderby=lastModifiedDateTime desc`;
    return await callGraph(path);
  }

  return {
    getMe,
    isPortalAdmin,
    getOwnedManagedGroupIds,
    listManagedGroups,
    listGroupMembers,
    listGroupOwners,
    listAllManagedUsers,
    searchUsers,
    createGroup,
    updateGroup,
    getGroupSiteUrl,
    addOwner, removeOwner,
    addMember, removeMember,
    getUserMemberOf,
    getUserOwnedGroups,
    removeUserLicense,
    disableUserAccount,
    enableUserAccount,
    deleteUser,
    createUser,
    assignUserLicense,
    updateUser,
    sendMail,
    getSubscribedSkus,
    createMemberRequest,
    submitPasswordReset,
    logAuditEntry,
    listAuditLog,
    EVAA_LICENSE_SKU_ID,
  };
})();
