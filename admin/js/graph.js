// EVAA Admin Portal — Microsoft Graph API wrapper
// Thin fetch-based client. All calls require a token from AUTH.getToken().

const GRAPH = (() => {
  const BASE = "https://graph.microsoft.com/v1.0";
  const PORTAL_ADMINS_GROUP_ID = "98d51c39-149a-4dbf-9e86-1510035d8239";

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
    if (resp.status === 204) return null;
    return resp.json();
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

  // True if the current user is a member (or owner-as-member) of EVAA Portal Admins.
  async function isPortalAdmin() {
    try {
      // Most efficient check: /me/checkMemberGroups
      const result = await callGraph("/me/checkMemberGroups", {
        method: "POST",
        body: JSON.stringify({ groupIds: [PORTAL_ADMINS_GROUP_ID] }),
      });
      return Array.isArray(result.value) && result.value.includes(PORTAL_ADMINS_GROUP_ID);
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
    const path = `/groups/${groupId}/members/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`;
    return callGraphAll(path);
  }

  async function listGroupOwners(groupId) {
    const path = `/groups/${groupId}/owners/microsoft.graph.user?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`;
    return callGraphAll(path);
  }

  // List all enabled users in the EVAA + Fusion domains.
  // Used by the Manage Members tab for the paginated all-users view.
  async function listAllManagedUsers() {
    const filter = "accountEnabled eq true and (endswith(userPrincipalName,'@evaasports.org') or endswith(userPrincipalName,'@avfusion.org'))";
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

  // Generic user-field update. Used by the inline-edit UI on user detail view.
  // patch is a partial user object, e.g. { jobTitle: "Baseball: VP" } or { displayName: "..." }
  async function updateUser(userId, patch) {
    return callGraph(`/users/${userId}`, {
      method: "PATCH",
      body: JSON.stringify(patch),
    });
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
    return callGraph(`/me/sendMail`, {
      method: "POST",
      body: JSON.stringify(body),
    });
  }

  return {
    getMe,
    isPortalAdmin,
    listManagedGroups,
    listGroupMembers,
    listGroupOwners,
    listAllManagedUsers,
    searchUsers,
    addOwner, removeOwner,
    addMember, removeMember,
    getUserMemberOf,
    getUserOwnedGroups,
    removeUserLicense,
    disableUserAccount,
    enableUserAccount,
    createUser,
    assignUserLicense,
    updateUser,
    sendMail,
    EVAA_LICENSE_SKU_ID,
  };
})();
