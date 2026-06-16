# Power Automate flow spec — `EVAA - Submit Finance Request`

Mirrors the existing helper-flow pattern (e.g., `EVAA - Submit Member Request`).

> ⚠️ **TESTING NOTE (2026-06-15):** Treasurer notifications temporarily routed to
> `web-admin@evaasports.org` instead of `treasurer@evaasports.org` while Mark validates
> the flow. Swap back before announcing the feature.

## Purpose
Receives a JSON payload from the public `/finance/index.html` submission form (signed-in EVAA user, but with no direct SharePoint write rights), then performs the privileged work:

1. Creates a new item in the `FinanceRequests` SharePoint list.
2. If a receipt file was attached, uploads it to the `FinanceReceipts` document library and writes the resulting URL back to the list item.
3. Emails `web-admin@evaasports.org` with the new submission summary + link.
4. Emails the submitter a confirmation.

Submitters never need SharePoint permissions because the flow runs as `web-admin@evaasports.org` (the flow owner).

---

## Trigger

| Property | Value |
|---|---|
| Type | **When an HTTP request is received** |
| Method | POST |
| Auth | SAS-signed (signature in the URL — safe to ship in client JS, same pattern as the existing helper flows) |
| Content-type | `application/json` |

### Request body schema

```jsonc
{
  "RequestType": "Check Request | Deposit Request | Credit Card Use",

  "FirstName": "string",
  "LastName": "string",
  "ContactEmail": "string",
  "Phone1": "string",
  "Phone2": "string | null",
  "DateRequested": "YYYY-MM-DD",

  "Sport": "string (from schema.js SPORTS)",
  "ProgramType": "string | null",
  "TravelingSubtype": "string | null",
  "Season": "string | null",
  "Category": "string (from schema.js EXPENSE_CATEGORY)",
  "Amount": "number (USD)",
  "Notes": "string | null",

  "VendorCardinality": "Single Vendor | Multiple Vendors | null",
  "VendorName": "string | null",
  "PayerName": "string | null",
  "VendorNote_Multi": "string | null",
  "PayeeAddress": "string | null",
  "PayeeCity": "string | null",
  "PayeeState": "string | null",
  "PayeeZip": "string | null",
  "AltPayeeName": "string | null",
  "AltPayeeAddress": "string | null",
  "AltPayeeCity": "string | null",
  "AltPayeeState": "string | null",
  "AltPayeeZip": "string | null",
  "CardLast4": "string | null",
  "TransactionDate": "YYYY-MM-DD | null",

  "Receipt": {
    "fileName": "string",
    "mimeType": "string",
    "base64": "string (data URI body — strip the data:...;base64, prefix client-side)"
  },

  "SourceSystem": "Portal",
  "ClientRequestId": "string (a UUID the client generates so retries are idempotent)"
}
```

Submitter identity is determined client-side from the MSAL account and put in `FirstName`/`LastName`/`ContactEmail`. The flow trusts these (the same trust model as `Submit Member Request`).

---

## Actions

### 1. Compute_Title (Compose)
```
@{concat(triggerBody()?['RequestType'], ' — ', triggerBody()?['Sport'], ' — $',
         string(triggerBody()?['Amount']), ' — ',
         coalesce(triggerBody()?['VendorName'], triggerBody()?['PayerName'], '(unspecified)'))}
```

### 2. Create_FinanceRequest_Item (SharePoint → Create item)
- Site: `https://evaasports.sharepoint.com/sites/EVAABoardPortal`
- List: `FinanceRequests`
- Map every payload field to its column. `Title` = output of step 1. `Status` = `Submitted`. `SubmittedByEmail` = `triggerBody()?['ContactEmail']`. `SubmittedByName` = `concat(FirstName, ' ', LastName)`.
- **Choice fields** use the `/Value` suffix style (e.g., `RequestType/Value`).
- Capture the returned item id for next step.

### 3. (Condition) — `length(triggerBody()?['Receipt']?['base64']) > 0`

#### True branch (receipt present):
- **Upload_Receipt** (SharePoint → Create file)
  - Site path: same as list
  - Folder path: `/FinanceReceipts`
  - File name: `@{concat(string(outputs('Create_FinanceRequest_Item')?['body/ID']), '__', triggerBody()?['Receipt']?['fileName'])}`
  - File content: `@base64ToBinary(triggerBody()?['Receipt']?['base64'])`

> 📝 **Why no separate patch step**: the original spec had a `Patch_Receipt_Url` action to write the receipt's URL onto the `FinanceRequests` item. The SP "Update item" operation strictly requires all required columns (Title, Amount, etc.) even on partial updates, which made the action error-prone. The treasurer console derives the receipt URL on-demand by scanning the `FinanceReceipts` library for files prefixed with the item ID (`GRAPH.getReceiptForRequest()` in `/finance/js/graph.js`). Simpler and avoids the constraint.

#### Else branch: no-op.

### 4. Email_Treasurer (Office 365 Outlook → Send email V2)
- To: `web-admin@evaasports.org`
- Subject: `New {RequestType}: {Sport} — ${Amount} — {VendorName/PayerName}`
- Body (HTML; same style language as the existing Approval Flow emails):
  - Greeting
  - Summary table of all submitted fields
  - Receipt link (if uploaded)
  - Big CTA button → treasurer console URL with `?id={ListItemId}` deep-link parameter
  - Reply-to: the submitter's `ContactEmail` (so a quick "we need more info" reply goes back to them)

### 5. Email_Submitter (Office 365 Outlook → Send email V2)
- To: `triggerBody()?['ContactEmail']`
- Subject: `Your EVAA {RequestType} was received`
- Body: confirmation summary + "We'll review and follow up. You can reply to this email with questions."

### 6. Response (HTTP Response)
- Status: 200
- Body: `{ "status": "submitted", "id": "@{outputs('Create_FinanceRequest_Item')?['body/ID']}" }`

### Error handling
- Wrap steps 2–5 in a Scope.
- ERROR_EMAIL handler (configured to run on scope failure) sends to `web-admin@evaasports.org` with the run-history URL — mirrors the existing Approval Flow's `ERROR EMAIL_Send_an_email_(V2)_3_1` pattern.

---

## After build — connect the front end

1. After the flow saves, copy the generated **HTTP POST URL** (with the SAS signature) from the trigger's "URL" panel.
2. Paste it into `/finance/js/graph.js` as `SUBMIT_FINANCE_REQUEST_URL`.
3. Test by submitting a small Check request and confirming:
   - Row appears in `FinanceRequests` with `Status = Submitted`
   - Receipt file appears in `FinanceReceipts` (if you attached one)
   - Treasurer + submitter emails arrive

---

## Treasurer console — separate path

The treasurer console (`/finance/admin.html`) does NOT go through this flow. It calls Microsoft Graph directly using the treasurer's `Sites.ReadWrite.All` scope to:
- List `FinanceRequests`
- Update item Status + DecisionAt + DecisionBy + TreasurerNotes
- Write `AdminActionLog` entries for each status change (audit trail)

No second flow needed because treasurers are full admins.
