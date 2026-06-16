// EVAA Finance Portal — Microsoft Graph API wrapper + helper-flow submit
// Mirrors /admin/js/graph.js: thin fetch-based client.
// Submission path uses the helper-flow pattern (SAS-signed Power Automate trigger URL).
// Console path uses Graph directly with the treasurer's Sites.ReadWrite.All scope.

const GRAPH = (() => {
  const BASE = "https://graph.microsoft.com/v1.0";

  // SharePoint site path → site ID lookup is cached for the session.
  // Same site the existing admin portal uses (EVAABoardPortal).
  const SITE_HOSTNAME = "evaasports.sharepoint.com";
  const SITE_PATH = "/sites/EVAABoardPortal";
  let _siteId = null;

  // List + library names (must match the SP provisioning)
  const FINANCE_LIST = "FinanceRequests";
  const RECEIPT_LIBRARY = "FinanceReceipts";
  const AUDIT_LIST = "AdminActionLog";
  const OPTIONS_LIST = "FinanceFormOptions";

  // Option types — must match FinanceFormOptions OptionType choice column
  const OPTION_TYPES = ["Sport", "ExpenseCategory", "ProgramType", "TravelingSubtype", "Season", "VendorCardinality", "RequestType"];

  // Treasurer / Leadership group — gates console access.
  // Mirrors /admin/js/graph.js ADMIN_GROUP_IDS.
  const LEADERSHIP_GROUP_IDS = [
    "98d51c39-149a-4dbf-9e86-1510035d8239", // EVAA Portal Admins
    "12e5f9ce-d644-4052-aff8-b31e99c3acb9", // EVAA - Leadership (includes treasurer)
  ];

  // Helper-flow URL — SAS-signed Power Automate trigger for "EVAA - Submit Finance Request".
  // Mirrors the pattern in /admin/js/graph.js SUBMIT_MEMBER_REQUEST_URL.
  // Safe to ship client-side: rows land with Status=Submitted and require treasurer review.
  const SUBMIT_FINANCE_REQUEST_URL =
    "https://defaultb5897a1bb85b42bd8e619b021b67d2.ce.environment.api.powerplatform.com:443/" +
    "powerautomate/automations/direct/workflows/522553095c8a44789aadab98eabeb80f/" +
    "triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&" +
    "sig=voMKfSydAOdaC52hbwvFLcGNYtlOkQ-_GYlyeAz8Nac";

  // ─── Core ──────────────────────────────────────────────────────────────────
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
    if (resp.status === 204 || resp.status === 202) return null;
    const text = await resp.text();
    if (!text) return null;
    try { return JSON.parse(text); } catch { return null; }
  }

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
    return callGraph("/me?$select=id,displayName,mail,userPrincipalName,jobTitle,givenName,surname,businessPhones");
  }

  async function isTreasurerOrAdmin() {
    try {
      const result = await callGraph("/me/checkMemberGroups", {
        method: "POST",
        body: JSON.stringify({ groupIds: LEADERSHIP_GROUP_IDS }),
      });
      return Array.isArray(result.value) && result.value.some((id) => LEADERSHIP_GROUP_IDS.includes(id));
    } catch (err) {
      console.error("Leadership check failed:", err);
      return false;
    }
  }

  // ─── Site ID lookup (cached) ───────────────────────────────────────────────
  async function getSiteId() {
    if (_siteId) return _siteId;
    const site = await callGraph(`/sites/${SITE_HOSTNAME}:${SITE_PATH}?$select=id`);
    _siteId = site.id;
    return _siteId;
  }

  // ─── Submission path (helper flow) ─────────────────────────────────────────
  // Compress an image file to a target ≤ ~2 MB with max dimension 1600 px.
  // Iterates JPEG quality down if the first pass is still too big.
  // Returns a Blob (with a `.name` matching the original) or the original file
  // unchanged for non-image files / on any failure (fail-open so weird files still upload).
  const COMPRESS_TARGET_BYTES = 2 * 1024 * 1024;  // 2 MB
  const COMPRESS_MAX_DIM = 1600;
  async function maybeCompressImage(file) {
    if (!file || !file.type || !file.type.startsWith("image/")) return file;
    if (file.size <= COMPRESS_TARGET_BYTES) return file; // already small enough
    try {
      const dataUri = await new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.onerror = () => reject(r.error);
        r.readAsDataURL(file);
      });
      const img = await new Promise((resolve, reject) => {
        const i = new Image();
        i.onload = () => resolve(i);
        i.onerror = () => reject(new Error("image decode failed"));
        i.src = dataUri;
      });
      // Resize to fit within MAX_DIM × MAX_DIM
      let { width: w, height: h } = img;
      const scale = Math.min(1, COMPRESS_MAX_DIM / Math.max(w, h));
      w = Math.round(w * scale);
      h = Math.round(h * scale);
      const canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);

      // Iterate quality: 0.85 → 0.75 → 0.65 → 0.55 → 0.45 until ≤ target
      const qualities = [0.85, 0.75, 0.65, 0.55, 0.45];
      for (const q of qualities) {
        const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", q));
        if (!blob) continue;
        if (blob.size <= COMPRESS_TARGET_BYTES || q === qualities[qualities.length - 1]) {
          // Keep original filename but switch extension to .jpg for clarity
          const baseName = (file.name || "receipt").replace(/\.[^.]+$/, "") + ".jpg";
          return new File([blob], baseName, { type: "image/jpeg", lastModified: Date.now() });
        }
      }
      return file; // unreachable, but safe
    } catch (e) {
      console.warn("Image compression failed, sending original:", e);
      return file;
    }
  }

  // Reads a File object (from <input type="file">) into a base64 string.
  // For images, runs through maybeCompressImage first.
  // Returns { fileName, mimeType, base64, originalBytes, sentBytes } or null if empty.
  async function readFileAsBase64(file) {
    if (!file) return null;
    const originalBytes = file.size;
    const compressed = await maybeCompressImage(file);
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUri = reader.result; // "data:image/jpeg;base64,XXXXX"
        const idx = dataUri.indexOf(",");
        resolve({
          fileName: compressed.name,
          mimeType: compressed.type || "application/octet-stream",
          base64: idx >= 0 ? dataUri.slice(idx + 1) : dataUri,
          originalBytes,
          sentBytes: compressed.size,
        });
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(compressed);
    });
  }

  // Generate a client-side UUID for the request — idempotency hint for the flow.
  function newClientRequestId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    // Fallback for older browsers
    return "frq-" + Date.now() + "-" + Math.random().toString(36).slice(2, 10);
  }

  // Post the form payload to the helper flow.
  // `values` is the object the form built (with file-typed entries replaced by File objects).
  async function submitFinanceRequest(values, receiptFile) {
    if (SUBMIT_FINANCE_REQUEST_URL.startsWith("TODO_")) {
      throw new Error(
        "Submit flow URL not configured yet. After deploying " +
        "'EVAA - Submit Finance Request' flow, paste its SAS-signed URL into graph.js."
      );
    }
    const receipt = receiptFile ? await readFileAsBase64(receiptFile) : null;
    const me = await getMe();
    const payload = {
      ...values,
      // Trust client-supplied contact info but stamp the signed-in identity so the
      // flow can cross-check / overwrite if values were tampered with.
      _signedInUpn: me?.userPrincipalName || me?.mail || "",
      _signedInName: me?.displayName || "",
      Receipt: receipt,
      SourceSystem: "Portal",
      ClientRequestId: newClientRequestId(),
    };
    const resp = await fetch(SUBMIT_FINANCE_REQUEST_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      throw new Error(`Submit-finance flow ${resp.status}: ${errText.slice(0, 300)}`);
    }
    return resp.json().catch(() => ({ status: "submitted" }));
  }

  // ─── FinanceFormOptions — dropdown source of truth ─────────────────────────
  // Cached per session so we don't re-fetch on every page render.
  let _optionsCache = null;
  async function getFinanceFormOptions({ force = false } = {}) {
    if (!force && _optionsCache) return _optionsCache;
    const siteId = await getSiteId();
    const items = await callGraphAll(
      `/sites/${siteId}/lists/${encodeURIComponent(OPTIONS_LIST)}/items?expand=fields&$top=999`
    );
    // Shape: [{ id, title, type, code, order, active }]
    const opts = items.map((it) => {
      const f = it.fields || {};
      return {
        id: it.id,
        title: f.Title || "",
        type: f.OptionType || "",
        code: f.AccountCode || "",
        order: typeof f.DisplayOrder === "number" ? f.DisplayOrder : 0,
        active: f.Active !== false,
      };
    });
    _optionsCache = opts;
    return opts;
  }

  // Return active options for a given type, sorted by DisplayOrder.
  async function getOptionsByType(type) {
    const all = await getFinanceFormOptions();
    return all
      .filter((o) => o.type === type && o.active)
      .sort((a, b) => (a.order - b.order) || a.title.localeCompare(b.title));
  }

  // Lookup the account code for a category title (or any type+title pair).
  // Returns the code string, or "" if not found / no code defined.
  async function lookupAccountCode(type, title) {
    const all = await getFinanceFormOptions();
    const match = all.find((o) => o.type === type && o.title === title);
    return (match && match.code) || "";
  }

  // CRUD for the Options tab in the admin console
  async function createFormOption({ title, type, code, order, active }) {
    const siteId = await getSiteId();
    const fields = { Title: title, OptionType: type, AccountCode: code || "", DisplayOrder: order || 0, Active: active !== false };
    _optionsCache = null;
    return callGraph(`/sites/${siteId}/lists/${encodeURIComponent(OPTIONS_LIST)}/items`, {
      method: "POST",
      body: JSON.stringify({ fields }),
    });
  }

  async function updateFormOption(itemId, patchFields) {
    const siteId = await getSiteId();
    _optionsCache = null;
    return callGraph(`/sites/${siteId}/lists/${encodeURIComponent(OPTIONS_LIST)}/items/${itemId}/fields`, {
      method: "PATCH",
      body: JSON.stringify(patchFields),
    });
  }

  async function deleteFormOption(itemId) {
    const siteId = await getSiteId();
    _optionsCache = null;
    return callGraph(`/sites/${siteId}/lists/${encodeURIComponent(OPTIONS_LIST)}/items/${itemId}`, {
      method: "DELETE",
    });
  }

  // ─── Treasurer console — Graph CRUD on FinanceRequests ─────────────────────

  // List all FinanceRequests with all fields expanded.
  // Caller filters/sorts client-side (the volume is low).
  async function listFinanceRequests() {
    const siteId = await getSiteId();
    const path = `/sites/${siteId}/lists/${encodeURIComponent(FINANCE_LIST)}/items?expand=fields&$top=999`;
    return callGraphAll(path);
  }

  async function getFinanceRequest(itemId) {
    const siteId = await getSiteId();
    return callGraph(`/sites/${siteId}/lists/${encodeURIComponent(FINANCE_LIST)}/items/${itemId}?expand=fields`);
  }

  // The submission flow uploads receipts to FinanceReceipts named "{itemId}__{filename}",
  // but does NOT patch the FinanceRequests item's ReceiptUrl field (removed from the flow
  // to dodge SP UpdateItem's required-field constraint). So the console finds the receipt
  // by scanning the library for files starting with that prefix.
  // Returns { name, webUrl } or null if no receipt was uploaded.
  async function getReceiptForRequest(itemId) {
    const siteId = await getSiteId();
    try {
      const items = await callGraphAll(
        `/sites/${siteId}/lists/${encodeURIComponent(RECEIPT_LIBRARY)}/items?expand=driveItem,fields(select=FileLeafRef)&$top=999`
      );
      const prefix = String(itemId) + "__";
      const match = items.find((it) => {
        const name = (it.fields && it.fields.FileLeafRef) || (it.driveItem && it.driveItem.name) || "";
        return name.startsWith(prefix);
      });
      if (!match) return null;
      const di = match.driveItem || {};
      const name = di.name || (match.fields && match.fields.FileLeafRef) || "receipt";
      // Fetch the @microsoft.graph.downloadUrl for inline preview — short-lived pre-authed URL.
      // The expand=driveItem already gives us name + webUrl but not downloadUrl, so fetch
      // the drive item directly.
      let downloadUrl = null;
      const driveId = di.parentReference && di.parentReference.driveId;
      if (driveId && di.id) {
        const full = await callGraph(`/drives/${driveId}/items/${di.id}?$select=name,webUrl,@microsoft.graph.downloadUrl,file`);
        downloadUrl = full && full["@microsoft.graph.downloadUrl"];
      }
      return {
        name,
        webUrl: di.webUrl,
        downloadUrl,
        mimeType: di.file && di.file.mimeType || (name.toLowerCase().endsWith(".pdf") ? "application/pdf" : "image/jpeg"),
      };
    } catch (e) {
      console.warn("Receipt lookup failed (non-fatal):", e);
      return null;
    }
  }

  // Update item status (+ optional notes). Audit-logs the change.
  // `transition` is the new Status string.
  async function updateRequestStatus(itemId, transition, treasurerNotes) {
    const siteId = await getSiteId();
    const me = await getMe();
    const decisionBy = me?.userPrincipalName || me?.mail || "(unknown)";
    const fields = {
      Status: transition,
      DecisionBy: decisionBy,
      DecisionAt: new Date().toISOString(),
    };
    if (typeof treasurerNotes === "string") fields.TreasurerNotes = treasurerNotes;
    const result = await callGraph(
      `/sites/${siteId}/lists/${encodeURIComponent(FINANCE_LIST)}/items/${itemId}/fields`,
      {
        method: "PATCH",
        body: JSON.stringify(fields),
      }
    );
    // Audit log
    await logAuditEntry({
      actionType: "Other",
      actor: decisionBy,
      title: `FinanceRequest ${itemId} → ${transition}`,
      notes: `Status changed to ${transition}.` + (treasurerNotes ? ` Notes: ${treasurerNotes.slice(0, 400)}` : ""),
      targetUser: "",
    }).catch((e) => console.warn("AuditLog write failed (non-fatal):", e));
    return result;
  }

  // Mirror of the existing /admin AdminActionLog write pattern.
  async function logAuditEntry({ actionType, actor, title, notes, targetUser }) {
    const siteId = await getSiteId();
    const fields = {
      Title: title,
      Actor: actor,
      ActionType: actionType,        // Choice value (use exact API value e.g. "Other")
      TargetUser: targetUser || "",
      Notes: notes || "",
    };
    return callGraph(
      `/sites/${siteId}/lists/${encodeURIComponent(AUDIT_LIST)}/items`,
      {
        method: "POST",
        body: JSON.stringify({ fields }),
      }
    );
  }

  return {
    callGraph,
    callGraphAll,
    getMe,
    isTreasurerOrAdmin,
    getSiteId,
    submitFinanceRequest,
    listFinanceRequests,
    getFinanceRequest,
    getReceiptForRequest,
    updateRequestStatus,
    logAuditEntry,
    readFileAsBase64,
    getFinanceFormOptions,
    getOptionsByType,
    lookupAccountCode,
    createFormOption,
    updateFormOption,
    deleteFormOption,
    OPTION_TYPES,
  };
})();
