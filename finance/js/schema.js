// EVAA Finance Requests — form schema
// Schema definition + dynamic option loading from FinanceFormOptions SP list.
//
// The form *shape* (sections, field IDs, conditional visibility) is fixed here.
// The dropdown *contents* live in SharePoint list "FinanceFormOptions" and are
// loaded async via GRAPH.getOptionsByType(). Call `await SCHEMA.load()` once on
// page boot before the first render — see app.js.
//
// To add a new sport / category / etc., treasurer edits the SP list. No code change.

const SCHEMA = (() => {
  // ─── DYNAMIC ENUMS — baked defaults, refreshed by load() ──────────────────
  // These defaults are a SNAPSHOT of the FinanceFormOptions SP list (2026-07-20).
  // They are populated (not empty) on purpose: the submission form is used by
  // coaches / members / AV Fusion folks who do NOT have access to the board
  // SharePoint site, so their SCHEMA.load() read of the options list fails. With
  // real values baked in here, the form still shows every dropdown for everyone.
  // The treasurer console (board members, who DO have site access) refreshes these
  // from the live list via load(). ⚠️ When the treasurer edits options in the SP
  // list, re-snapshot this block so the public form stays in sync (or move options
  // to a public feed — see note in load()).
  const enums = {
    REQUEST_TYPES: ["Check Request", "Deposit Request", "Credit Card Use"],
    SPORTS: [
      "Baseball", "Basketball", "Cross Country Running", "Football", "Lacrosse",
      "Soccer", "Softball", "Tennis", "Track", "Trap", "Volleyball", "Wrestling",
      "XC Ski", "Multi-Sport / Operations", "Other (note in details)",
    ],
    PROGRAM_TYPE: ["Traveling", "In-House", "All", "N/A"],
    TRAVELING_SUBTYPE: ["Fall Traveling", "Spring Traveling", "Summer Traveling", "Winter Traveling"],
    SEASON: [
      "This request is for the current season.",
      "This request is for the upcoming season.",
    ],
    EXPENSE_CATEGORY: [
      "Coaching Fees", "Tournament Registration Fees", "Tournament Registration Fees Income",
      "Concession Income", "Equipment", "Uniforms", "Field/Facility Rental", "Officials / Refs",
      "Awards / Trophies", "Insurance", "Background Checks", "Training / Clinics",
      "Travel / Lodging", "Marketing / Promotion", "Office / Supplies", "Bank Fees",
      "Reimbursement", "Miscellaneous Income", "Other (note in details)",
    ],
    VENDOR_CARDINALITY: ["Single Vendor", "Multiple Vendors"],
  };

  // Async loader — pulls FinanceFormOptions and populates `enums`.
  // Idempotent (cached in GRAPH layer); safe to await multiple times.
  async function load() {
    if (typeof GRAPH === "undefined" || !GRAPH.getOptionsByType) return; // caller didn't load graph.js yet — defaults stay
    try {
      const [sports, cats, prog, travel, season, vendor, req] = await Promise.all([
        GRAPH.getOptionsByType("Sport"),
        GRAPH.getOptionsByType("ExpenseCategory"),
        GRAPH.getOptionsByType("ProgramType"),
        GRAPH.getOptionsByType("TravelingSubtype"),
        GRAPH.getOptionsByType("Season"),
        GRAPH.getOptionsByType("VendorCardinality"),
        GRAPH.getOptionsByType("RequestType"),
      ]);
      // Only overwrite a default when the read actually returned rows. A user
      // without board-site access can get an empty/trimmed result — in that case
      // keep the baked defaults instead of blanking the dropdowns.
      if (sports.length) enums.SPORTS            = sports.map((o) => o.title);
      if (cats.length)   enums.EXPENSE_CATEGORY  = cats.map((o) => o.title);
      if (prog.length)   enums.PROGRAM_TYPE      = prog.map((o) => o.title);
      if (travel.length) enums.TRAVELING_SUBTYPE = travel.map((o) => o.title);
      if (season.length) enums.SEASON            = season.map((o) => o.title);
      if (vendor.length) enums.VENDOR_CARDINALITY = vendor.map((o) => o.title);
      if (req.length)    enums.REQUEST_TYPES     = req.map((o) => o.title);
    } catch (e) {
      console.warn("SCHEMA.load() failed — falling back to defaults:", e);
    }
  }

  // ─── FORM SECTIONS ────────────────────────────────────────────────────────
  // Sections reference enums.* indirectly via a getter so the renderer always
  // sees the latest values after load().
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
          get options() { return enums.REQUEST_TYPES; },
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
          get options() { return enums.VENDOR_CARDINALITY; },
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
        // Vendor address — required only for Check Request (we need to mail the check
        // *somewhere*). Optional for Deposit (informational) and Credit Card Use.
        { id: "PayeeAddress", label: "Street address", type: "text",
          requiredWhen: { field: "RequestType", equals: "Check Request" } },
        { id: "PayeeCity", label: "City", type: "text",
          requiredWhen: { field: "RequestType", equals: "Check Request" } },
        { id: "PayeeState", label: "State / Province", type: "text",
          requiredWhen: { field: "RequestType", equals: "Check Request" } },
        { id: "PayeeZip", label: "Zip", type: "text",
          requiredWhen: { field: "RequestType", equals: "Check Request" } },
        // "Send to a different address" toggle — Check only.
        // When checked, the AltPayee* fields below it become visible inline.
        {
          id: "DifferentMailingAddress",
          label: "Send the check to a different address",
          type: "boolean",
          required: false,
          when: { field: "RequestType", equals: "Check Request" },
          help: "Check this if the check should be mailed somewhere other than the vendor address.",
        },
        // Alt mailing address — appears directly under the checkbox when it's ticked.
        // Required iff visible.
        { id: "AltPayeeName", label: "Recipient name", type: "text",
          when: { field: "DifferentMailingAddress", equals: true },
          requiredWhen: { field: "DifferentMailingAddress", equals: true },
          help: "Who the check is made out to / sent to." },
        { id: "AltPayeeAddress", label: "Street address", type: "text",
          when: { field: "DifferentMailingAddress", equals: true },
          requiredWhen: { field: "DifferentMailingAddress", equals: true } },
        { id: "AltPayeeCity", label: "City", type: "text",
          when: { field: "DifferentMailingAddress", equals: true },
          requiredWhen: { field: "DifferentMailingAddress", equals: true } },
        { id: "AltPayeeState", label: "State / Province", type: "text",
          when: { field: "DifferentMailingAddress", equals: true },
          requiredWhen: { field: "DifferentMailingAddress", equals: true } },
        { id: "AltPayeeZip", label: "Zip", type: "text",
          when: { field: "DifferentMailingAddress", equals: true },
          requiredWhen: { field: "DifferentMailingAddress", equals: true } },
      ],
    },
    {
      id: "details",
      title: "Details",
      fields: [
        { id: "Sport", label: "Sport / Program", type: "choice", required: true, get options() { return enums.SPORTS; },
          help: "This is how the treasurer reconciles the request — pick the right sport even if the spend is shared." },
        { id: "ProgramType", label: "Program type", type: "choice", required: false, get options() { return enums.PROGRAM_TYPE; } },
        {
          id: "TravelingSubtype",
          label: "Which traveling season?",
          type: "choice",
          required: false,
          get options() { return enums.TRAVELING_SUBTYPE; },
          when: { field: "ProgramType", equals: "Traveling" },
        },
        { id: "Season", label: "What season is this request for?", type: "choice", required: false, get options() { return enums.SEASON; } },
        { id: "Category", label: "Expense / income category", type: "choice", required: true, get options() { return enums.EXPENSE_CATEGORY; } },
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
    // (alternate_payee section removed — alt-address fields moved inline above,
    // appearing right under the "Send to a different address" checkbox.)
  ];

  // ─── EVALUATOR ────────────────────────────────────────────────────────────
  function isVisible(node, values) {
    const cond = node.when;
    if (!cond) return true;
    const actual = values[cond.field];
    if (Array.isArray(cond.equals)) return cond.equals.includes(actual);
    return actual === cond.equals;
  }

  function labelFor(field, values) {
    if (field.labelByType && values && values.RequestType && field.labelByType[values.RequestType]) {
      return field.labelByType[values.RequestType];
    }
    return field.label;
  }

  // Resolve a field's required-ness given current form values.
  // Honors both `required: true` and `requiredWhen: { field, equals }`.
  function isRequired(field, values) {
    if (field.required) return true;
    const rw = field.requiredWhen;
    if (!rw) return false;
    const actual = values[rw.field];
    if (Array.isArray(rw.equals)) return rw.equals.includes(actual);
    return actual === rw.equals;
  }

  // Set of field IDs whose value-change should trigger a full re-render
  // (because some other field's visibility/required-ness depends on them).
  // Computed from all `when` and `requiredWhen` predicates across the schema.
  function gatingFieldIds() {
    const ids = new Set();
    for (const sec of sections) {
      if (sec.when) ids.add(sec.when.field);
      for (const f of sec.fields) {
        if (f.when) ids.add(f.when.field);
        if (f.requiredWhen) ids.add(f.requiredWhen.field);
      }
    }
    return ids;
  }

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
    load,
    sections,
    isVisible,
    labelFor,
    isRequired,
    visibleFields,
    gatingFieldIds,
    enums,
    sourceSurveyId: 6743,
    sourceUrl: "https://evaa.sportngin.com/survey/show/6743",
  };
})();
