# SharePoint provisioning spec — Finance Requests

Site: `https://evaasports.sharepoint.com/sites/EVAABoardPortal` (existing, reuse)

Two artifacts to create.

---

## 1. List: `FinanceRequests`

| Internal name | Display name | Type | Required | Notes |
|---|---|---|---|---|
| `Title` | Title | Single line of text | yes (auto-fill on create) | Set by helper flow: `"{RequestType} — {Sport} — {Amount} — {VendorName or PayerName}"`. |
| `RequestType` | Request Type | Choice (single) | yes | Options: `Check Request`, `Deposit Request`, `Credit Card Use`. **Use API values, not display names** (gotcha #33). |
| `Status` | Status | Choice (single) | yes, default `Submitted` | Options in workflow order: `Submitted`, `Under Review`, `Approved`, `Paid/Deposited`, `Denied`. |
| `SubmittedByEmail` | Submitted By (Email) | Single line of text | yes | **Text, not Person.** Person columns need full Claims format and have bitten the existing flows (gotcha #3). |
| `SubmittedByName` | Submitted By (Name) | Single line of text | yes | |
| `SubmitterPhone` | Submitter Phone | Single line of text | no | |
| `SubmitterAltPhone` | Submitter Alternate Phone | Single line of text | no | |
| `DateRequested` | Date Requested | Date and time (date only) | yes | |
| `Sport` | Sport | Choice (single) | yes | Mirror schema.js SPORTS exactly. Include `Multi-Sport / Operations` + `Other (note in details)`. |
| `ProgramType` | Program Type | Choice (single) | no | Options: `Traveling`, `In-House`, `All`, `N/A`. |
| `TravelingSubtype` | Traveling Subtype | Choice (single) | no | Options: `Fall Traveling`, `Spring Traveling`, `Summer Traveling`, `Winter Traveling`. |
| `Season` | Season | Choice (single) | no | Options: `This request is for the current season.`, `This request is for the upcoming season.` |
| `Category` | Category | Choice (single) | yes | Mirror schema.js EXPENSE_CATEGORY exactly. |
| `Amount` | Amount | Currency (USD) | yes | 2 decimal places. |
| `Notes` | Notes | Multiple lines of text (plain) | no | Enhanced rich text off; keep as plain to simplify Graph reads. |
| `VendorCardinality` | Vendor Cardinality | Choice (single) | no | Options: `Single Vendor`, `Multiple Vendors`. |
| `VendorName` | Vendor Name | Single line of text | no | Check/Credit Card single-vendor case. |
| `PayerName` | Payer Name | Single line of text | no | Deposit case. |
| `VendorNote_Multi` | Multi-Vendor Note | Multiple lines (plain) | no | Used when VendorCardinality = Multiple Vendors. |
| `PayeeAddress` | Payee Address | Single line of text | no | |
| `PayeeCity` | Payee City | Single line of text | no | |
| `PayeeState` | Payee State | Single line of text | no | |
| `PayeeZip` | Payee Zip | Single line of text | no | |
| `AltPayeeName` | Alt Payee Name | Single line of text | no | "Make check payable to" — Check Request only. |
| `AltPayeeAddress` | Alt Payee Address | Single line of text | no | |
| `AltPayeeCity` | Alt Payee City | Single line of text | no | |
| `AltPayeeState` | Alt Payee State | Single line of text | no | |
| `AltPayeeZip` | Alt Payee Zip | Single line of text | no | |
| `ReceiptUrl` | Receipt URL | Hyperlink | no | Set by helper flow after uploading the receipt file to `FinanceReceipts`. URL points to the document in the library. |
| `ReceiptFileName` | Receipt File Name | Single line of text | no | Original uploaded filename. |
| `SourceSystem` | Source System | Choice (single) | yes, default `Portal` | Options: `Portal`, `SportsEngine`. Lets the future SE-history-sync writer distinguish. |
| `SE_SubmissionId` | SE Submission ID | Single line of text | no | For dedup when SE-history-sync runs later. Empty for portal-submitted items. |
| `TreasurerNotes` | Treasurer Notes | Multiple lines (plain) | no | Treasurer-only edits via console. |
| `DecisionAt` | Decision At | Date and time (full datetime) | no | Auto-set when console moves Status out of `Submitted` for the first time. |
| `DecisionBy` | Decision By | Single line of text | no | UPN of the treasurer making the decision. |

### Views
- **Default**: All Items, sorted by Created descending.
- **Open Queue** (treasurer's main view): Status ∈ {`Submitted`, `Under Review`, `Approved`}, sorted by DateRequested ascending.
- **Closed**: Status ∈ {`Paid/Deposited`, `Denied`}, sorted by DecisionAt descending.

### Permissions
- Inherit from site (default). Submitters do NOT need direct access — the helper flow writes on their behalf.
- Treasurer (`EVAA - Leadership` members) read/write via the admin console using existing Sites.ReadWrite.All scope.

---

## 2. Document Library: `FinanceReceipts`

| Setting | Value |
|---|---|
| Type | Document library |
| Versioning | Major versions, retain last 10 |
| Require check-out | No |
| File-naming convention | `{ListItemId}__{Original-filename}` — set by the helper flow |
| Folder structure | Flat (flow uses ListItemId prefix to avoid name collisions) |
| Allowed file types | All; flow will reject anything > 10 MB |
| Permissions | Inherit from site |

The helper flow uploads the receipt here, captures the returned web URL, and writes it to `FinanceRequests.ReceiptUrl` on the same item.

---

## How to provision

Two options:

**Option A — UI (fastest, recommended):**
1. Go to `https://evaasports.sharepoint.com/sites/EVAABoardPortal` → Site contents → New → List → `FinanceRequests`
2. Add each column per the table above. Use API values (no spaces) for choice options where the spec calls it out.
3. Site contents → New → Document library → `FinanceReceipts`
4. Configure the 3 views listed above.

**Option B — PowerShell (PnP) or Graph batch:**
A provisioning script can be added later. Not built now since it's a one-time setup and the UI is faster for this scale.

---

## Reuse / consistency

- Field-naming follows the convention used in MemberRequests (no spaces, PascalCase).
- Choice columns mirror the JS schema enums in `/finance/js/schema.js` exactly. **When updating one, update both** — the form renders from JS, but SP rejects writes for choice values not in its options list.
- The audit-log writer pattern (`AuditLog_*` actions in the Approval Flow) is reused for status changes in the treasurer console — see flow spec.
