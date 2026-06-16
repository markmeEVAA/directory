// EVAA Finance Requests — form schema
// Config-driven so the UI in app.js renders whatever lives here.
// Source: SportsEngine survey #6743 ("Credit Card / Deposits / Check Request Form"),
// reconstructed from three real submissions in the treasurer mail thread (2020–21):
// 2 Check variants (single + multiple vendor) and 1 Deposit. The SE form was a
// single set of questions with a type selector at the top — this schema mirrors that.
//
// Design notes:
//   - One unified question set across Check / Deposit / Credit Card.
//   - Only narrow conditional bits:
//       * Vendor/Payer NAME label flips from "Vendor name" → "Payer name / company" for Deposit (matches SE).
//       * The "Recipient (if different)" trailing block only shows for Check (per Mark — that's
//         the alternate "make the check payable to" mailing address. Deposit and CC don't have
//         a check to mail.)
//   - Sport gives the treasurer enough context to reconcile against the credit card statement;
//     no separate CardLast4 / TransactionDate fields. If the submitter wants to note them, the
//     Notes field is fine.
//
// Field shape:
//   { id, label, type, required, options?, when?, help?, labelByType? }
//   labelByType: { [RequestType]: "alternate label" } — picked at render time.
//   type: text | email | tel | date | number | currency | choice | textarea | file
//   when: optional predicate { field, equals: <value or array> } controlling visibility

const SCHEMA = (() => {
  // ─── SHARED ENUMS ─────────────────────────────────────────────────────────────
  const REQUEST_TYPES = ["Check Request", "Deposit Request", "Credit Card Use"];

  const SPORTS = [
    "Baseball",
    "Basketball",
    "Cross Country Running",
    "Football",
    "Lacrosse",
    "Soccer",
    "Softball",
    "Tennis",
    "Track",
    "Trap",
    "Volleyball",
    "Wrestling",
    "XC Ski",
    "Multi-Sport / Operations",
    "Other (note in details)",
  ];

  // Program type (the existing SE form just had "Traveling" — we add In-House etc.)
  const PROGRAM_TYPE = ["Traveling", "In-House", "All", "N/A"];

  // Only when PROGRAM_TYPE === "Traveling"
  const TRAVELING_SUBTYPE = ["Fall Traveling", "Spring Traveling", "Summer Traveling", "Winter Traveling"];

  // Seasons (existing SE)
  const SEASON = ["This request is for the current season.", "This request is for the upcoming season."];

  // Expense categories — extracted from real submissions + the typical EVAA list
  const EXPENSE_CATEGORY = [
    "Coaching Fees",
    "Tournament Registration Fees",
    "Tournament Registration Fees Income",
    "Concession Income",
    "Equipment",
    "Uniforms",
    "Field/Facility Rental",
    "Officials / Refs",
    "Awards / Trophies",
    "Insurance",
    "Background Checks",
    "Training / Clinics",
    "Travel / Lodging",
    "Marketing / Promotion",
    "Office / Supplies",
    "Bank Fees",
    "Reimbursement",
    "Miscellaneous Income",
    "Other (note in details)",
  ];

  // Vendor cardinality (same question for all three types, per SE)
  const VENDOR_CARDINALITY = ["Single Vendor", "Multiple Vendors"];

  // ─── FORM SECTIONS ────────────────────────────────────────────────────────────

  const sections = [
    {
      id: "request_type",
      title: "Request type",
      fields: [
        {
          id: "RequestType",
          label: "Is this a check request, deposit request, or credit card use?",
          type: "choice",
          required: true,
          options: REQUEST_TYPES,
          help: "All other questions are the same regardless of choice. We just need to know which.",
        },
      ],
    },
    {
      id: "contact",
      title: "Your contact info",
      fields: [
        { id: "FirstName", label: "First name", type: "text", required: true },
        { id: "LastName", label: "Last name", type: "text", required: true },
        { id: "ContactEmail", label: "Email", type: "email", required: true, help: "We use this for follow-ups." },
        { id: "Phone1", label: "Phone", type: "tel", required: true },
        { id: "Phone2", label: "Alternate phone (optional)", type: "tel", required: false },
        { id: "DateRequested", label: "Date requested", type: "date", required: true, help: "Defaults to today." },
      ],
    },
    {
      id: "counterparty",
      title: "Vendor / Payer",
      fields: [
        {
          id: "VendorCardinality",
          label: "Single or multiple?",
          labelByType: {
            "Check Request": "Single vendor or multiple vendors?",
            "Credit Card Use": "Single vendor or multiple vendors?",
            "Deposit Request": "Single payer or multiple payers?",
          },
          type: "choice",
          required: true,
          options: VENDOR_CARDINALITY,
        },
        {
          id: "VendorName",
          label: "Vendor / payer name",
          labelByType: {
            "Check Request": "Vendor name",
            "Credit Card Use": "Vendor / merchant name",
            "Deposit Request": "Payer name / company",
          },
          type: "text",
          required: true,
          when: { field: "VendorCardinality", equals: "Single Vendor" },
          help: "Who the money goes to (Check / CC) or comes from (Deposit).",
        },
        {
          id: "VendorNote_Multi",
          label: "Brief summary (multiple)",
          type: "textarea",
          required: false,
          when: { field: "VendorCardinality", equals: "Multiple Vendors" },
          help: "Quick description; attach a document below with the full list if needed.",
        },
        { id: "PayeeAddress", label: "Street address", type: "text", required: false },
        { id: "PayeeCity", label: "City", type: "text", required: false },
        { id: "PayeeState", label: "State / Province", type: "text", required: false },
        { id: "PayeeZip", label: "Zip", type: "text", required: false },
      ],
    },
    {
      id: "details",
      title: "Details",
      fields: [
        { id: "Sport", label: "Sport / Program", type: "choice", required: true, options: SPORTS,
          help: "This is how the treasurer reconciles the request — pick the right sport even if the spend is shared." },
        { id: "ProgramType", label: "Program type", type: "choice", required: false, options: PROGRAM_TYPE },
        {
          id: "TravelingSubtype",
          label: "Which traveling season?",
          type: "choice",
          required: false,
          options: TRAVELING_SUBTYPE,
          when: { field: "ProgramType", equals: "Traveling" },
        },
        { id: "Season", label: "What season is this request for?", type: "choice", required: false, options: SEASON },
        { id: "Category", label: "Expense / income category", type: "choice", required: true, options: EXPENSE_CATEGORY },
        {
          id: "Amount",
          label: "Total amount",
          type: "currency",
          required: true,
          help: "USD. Enter as a number — e.g. 1400 or 1400.00",
        },
        {
          id: "Receipt",
          label: "Attach receipt / supporting document",
          type: "file",
          required: false,
          help: "Photo of receipt or PDF. On mobile, tap to use your camera. Images over 2 MB are auto-resized.",
        },
        {
          id: "Notes",
          label: "Notes / Invoice # / Additional context",
          type: "textarea",
          required: false,
          help: "Anything the treasurer should know — invoice number, card last 4, transaction date, etc.",
        },
      ],
    },
    {
      id: "alternate_payee",
      title: "Where to send the check",
      help: "Skip if the check should go to the vendor address above.",
      when: { field: "RequestType", equals: "Check Request" },
      fields: [
        { id: "AltPayeeName", label: "Recipient name", type: "text", required: false,
          help: "Who is the check made out to, if different from the vendor?" },
        { id: "AltPayeeAddress", label: "Street address", type: "text", required: false },
        { id: "AltPayeeCity", label: "City", type: "text", required: false },
        { id: "AltPayeeState", label: "State / Province", type: "text", required: false },
        { id: "AltPayeeZip", label: "Zip", type: "text", required: false },
      ],
    },
  ];

  // ─── EVALUATOR ────────────────────────────────────────────────────────────────
  function isVisible(node, values) {
    const cond = node.when;
    if (!cond) return true;
    const actual = values[cond.field];
    if (Array.isArray(cond.equals)) return cond.equals.includes(actual);
    return actual === cond.equals;
  }

  // Pick the right label given the chosen RequestType (supports labelByType override).
  function labelFor(field, values) {
    if (field.labelByType && values && values.RequestType && field.labelByType[values.RequestType]) {
      return field.labelByType[values.RequestType];
    }
    return field.label;
  }

  // Flatten visible fields (in section order) for validation + submit payload.
  function visibleFields(values) {
    const out = [];
    for (const sec of sections) {
      if (!isVisible(sec, values)) continue;
      for (const f of sec.fields) {
        if (!isVisible(f, values)) continue;
        out.push({ section: sec.id, ...f });
      }
    }
    return out;
  }

  return {
    sections,
    isVisible,
    labelFor,
    visibleFields,
    enums: { REQUEST_TYPES, SPORTS, PROGRAM_TYPE, TRAVELING_SUBTYPE, SEASON, EXPENSE_CATEGORY, VENDOR_CARDINALITY },
    sourceSurveyId: 6743,
    sourceUrl: "https://evaa.sportngin.com/survey/show/6743",
  };
})();
