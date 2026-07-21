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
      // Safari/iOS: keep auth state in a cookie too, so the redirect round-trip
      // survives ITP / sessionStorage quirks. Required for reliable redirect auth.
      storeAuthStateInCookie: true,
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

    // Redirect flow (Safari/iOS-safe — popups are blocked by Safari's popup blocker
    // and its non-synchronous-open rule). handleRedirectPromise() MUST run on every
    // load: it returns the auth result when we've just come back from Microsoft
    // sign-in, or null on a normal load.
    let redirectResult = null;
    try {
      redirectResult = await msalClient.handleRedirectPromise();
    } catch (e) {
      console.warn("handleRedirectPromise failed:", e);
    }

    if (redirectResult && redirectResult.account) {
      activeAccount = redirectResult.account;
      msalClient.setActiveAccount(activeAccount);
    } else {
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
    // Navigates away to Microsoft and returns to redirectUri; does NOT resolve with
    // an account. On return, init()/handleRedirectPromise() picks up the signed-in account.
    await msalClient.loginRedirect({ scopes: SCOPES });
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
      console.warn("Silent token acquisition failed, falling back to redirect:", err);
      // Interactive fallback via redirect (no popup). Navigates away and returns
      // to the page; the caller's in-flight promise won't resolve — that's expected.
      await msalClient.acquireTokenRedirect(request);
      return undefined;
    }
  }

  function getAccount() {
    return activeAccount;
  }

  return { init, signIn, signOut, getToken, getAccount };
})();
