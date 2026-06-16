// EVAA Finance Portal — MSAL authentication
// Mirrors /admin/js/auth.js: same client app reg, MSAL v3 SPA flow.
// Scope set is intentionally smaller — the submission form only needs User.Read,
// the treasurer console needs Sites.ReadWrite.All + Mail.Send for audit/notifications.

const AUTH = (() => {
  const TENANT_ID = "b5897a1b-b85b-42bd-8e61-9b021b67d2ce";
  const CLIENT_ID = "74b79a84-6ea0-4c32-beae-25ed9bdc249f";

  // Detect which page we're on and request only the scopes that page needs.
  // app.js (submission form) calls init() with consoleMode=false; admin-app.js with true.
  // Submitter needs Sites.Read.All to fetch FinanceFormOptions (dropdown source of truth).
  // Sites.Read.All is already admin-consented for this app reg via the existing
  // Sites.ReadWrite.All grant — Read is a subset.
  const SCOPES_SUBMITTER = ["User.Read", "Sites.Read.All"];
  const SCOPES_CONSOLE = ["User.Read", "Sites.ReadWrite.All", "Mail.Send"];

  let SCOPES = SCOPES_SUBMITTER;

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

  async function init({ consoleMode = false } = {}) {
    SCOPES = consoleMode ? SCOPES_CONSOLE : SCOPES_SUBMITTER;
    if (typeof window.msal === "undefined" || !window.msal.PublicClientApplication) {
      throw new Error("MSAL library failed to load (window.msal is undefined). Check network / CDN.");
    }
    msalClient = new window.msal.PublicClientApplication(msalConfig);
    await msalClient.initialize();
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
