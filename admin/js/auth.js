// EVAA Admin Portal — MSAL authentication
// Uses Microsoft Authentication Library (MSAL) v3 browser SPA flow.
// Tenant + Client ID match the existing EVAA app registration (74b79a84-...).

const AUTH = (() => {
  const TENANT_ID = "b5897a1b-b85b-42bd-8e61-9b021b67d2ce";
  const CLIENT_ID = "74b79a84-6ea0-4c32-beae-25ed9bdc249f";

  // Scopes the admin portal needs. All must be granted (delegated) on the app reg.
  //   User.Read                — current user profile
  //   Group.Read.All           — list groups + members + owners
  //   GroupMember.Read.All     — read group memberships
  //   Group.ReadWrite.All      — add/remove members + owners
  //   User.ReadWrite.All       — create / update / disable / enable users; assign licenses
  //   Mail.Send                — send welcome + admin-notify emails as the signed-in user
  const SCOPES = [
    "User.Read",
    "Group.Read.All",
    "GroupMember.Read.All",
    "Group.ReadWrite.All",
    "User.ReadWrite.All",
    "Mail.Send",
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

    // Handle redirect response if we just came back from sign-in
    const response = await msalClient.handleRedirectPromise();
    if (response && response.account) {
      msalClient.setActiveAccount(response.account);
      activeAccount = response.account;
    } else {
      // Try to pick up a cached account
      const accounts = msalClient.getAllAccounts();
      if (accounts.length > 0) {
        msalClient.setActiveAccount(accounts[0]);
        activeAccount = accounts[0];
      }
    }
    return activeAccount;
  }

  async function signIn() {
    if (!msalClient) throw new Error("Auth not initialized");
    return msalClient.loginRedirect({ scopes: SCOPES });
  }

  async function signOut() {
    if (!msalClient) return;
    return msalClient.logoutRedirect({
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
      console.warn("Silent token acquisition failed, falling back to interactive:", err);
      const result = await msalClient.acquireTokenRedirect(request);
      return result?.accessToken;
    }
  }

  function getAccount() {
    return activeAccount;
  }

  return { init, signIn, signOut, getToken, getAccount };
})();
