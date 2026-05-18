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

  return {
    getMe,
    isPortalAdmin,
    listManagedGroups,
    listGroupMembers,
    listGroupOwners,
    searchUsers,
    addOwner, removeOwner,
    addMember, removeMember,
  };
})();
