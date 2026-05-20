// EVAA Admin Portal — MSAL authentication
// Uses Microsoft Authentication Library (MSAL) v3 browser SPA flow.
// Tenant + Client ID match the existing EVAA app registration (74b79a84-...).

const AUTH = (() => {
  const TENANT_ID = "b5897a1b-b85b-42bd-8e61-9b021b67d2ce";
  const CLIENT_ID = "74b79a84-6ea0-4c32-beae-25ed9bdc249f";

  // Scopes the admin portal needs. All must be granted (delegated) on the app reg.
  //   User.Read                  — current user profile
  //   Group.Read.All             — list groups + members + owners
  //   GroupMember.Read.All       — read group memberships
  //   Group.ReadWrite.All        — add/remove members + owners
  //   User.ReadWrite.All         — create / update / disable / enable users; assign licenses
  //   Directory.AccessAsUser.All — REQUIRED in addition to User.ReadWrite.All for
  //                                passwordProfile updates (admin-mode password reset).
  //                                Per MS docs (May 2026), User.ReadWrite.All alone is no
  //                                longer sufficient for password reset / user disable —
  //                                you also need Directory.AccessAsUser.All. Without this
  //                                scope, Graph returns 403 Authorization_RequestDenied.
  //   Mail.Send                  — send welcome + admin-notify emails as the signed-in user
  //   Sites.ReadWrite.All        — write rows to MemberRequests SharePoint list (owner-mode Add/Remove)
  const SCOPES = [
    "User.Read",
    "Group.Read.All",
    "GroupMember.Read.All",
    "Group.ReadWrite.All",
    "User.ReadWrite.All",
    "Directory.AccessAsUser.All",
    "Mail.Send",
    "Sites.ReadWrite.All",
  ];

  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };

  let msalClient = null;
  let activeAccount = null;

  async function init() {
    // MSAL v3 browser SDK exposes itself as the global `msal`.
    if (typeof window.msal === "undefined" || !window.msal.PublicClientApplication) {
      throw new Error("MSAL library failed to load (window.msal is undefined). Check network / CDN.");
    }
    msalClient = new window.msal.PublicClientApplication(msalConfig);
    await msalClient.initialize();

    // Pick up a cached account if one exists (popup mode never produces a
    // redirect response, but a cached account survives reload within the
    // sessionStorage TTL).
    const accounts = msalClient.getAllAccounts();
    if (accounts.length > 0) {
      msalClient.setActiveAccount(accounts[0]);
      activeAccount = accounts[0];
    }
    return activeAccount;
  }

  async function signIn() {
    if (!msalClient) throw new Error("Auth not initialized");
    const result = await msalClient.loginPopup({ scopes: SCOPES });
    if (result && result.account) {
      msalClient.setActiveAccount(result.account);
      activeAccount = result.account;
    }
    return activeAccount;
  }

  async function signOut() {
    if (!msalClient) return;
    return msalClient.logoutPopup({
      postLogoutRedirectUri: msalConfig.auth.redirectUri,
    });
  }

  async function getToken() {
    if (!msalClient) throw new Error("Auth not initialized");
    if (!activeAccount) throw new Error("No signed-in account");
    const request = { scopes: SCOPES, account: activeAccount };
    try {
      const result = await msalClient.acquireTokenSilent(request);
      return result.accessToken;
    } catch (err) {
      console.warn("Silent token acquisition failed, falling back to popup:", err);
      const result = await msalClient.acquireTokenPopup(request);
      return result?.accessToken;
    }
  }

  function getAccount() {
    return activeAccount;
  }

  return { init, signIn, signOut, getToken, getAccount };
})();
