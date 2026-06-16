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
  // ─── DYNAMIC ENUMS — populated by load() ──────────────────────────────────
  // Defaults are sensible fallbacks so the form can render even if the options
  // list isn't reachable. Real values overwrite these on load().
  const enums = {
    REQUEST_TYPES: ["Check Request", "Deposit Request", "Credit Card Use"],
    SPORTS: [],
    PROGRAM_TYPE: [],
    TRAVELING_SUBTYPE: [],
    SEASON: [],
    EXPENSE_CATEGORY: [],
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
      enums.SPORTS            = sports.map((o) => o.title);
      enums.EXPENSE_CATEGORY  = cats.map((o) => o.title);
      enums.PROGRAM_TYPE      = prog.map((o) => o.title);
      enums.TRAVELING_SUBTYPE = travel.map((o) => o.title);
      enums.SEASON            = season.map((o) => o.title);
      if (vendor.length) enums.VENDOR_CARDINALITY = vendor.map((o) => o.title);
      if (req.length) enums.REQUEST_TYPES = req.map((o) => o.title);
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
    visibleFields,
    enums,
    sourceSurveyId: 6743,
    sourceUrl: "https://evaa.sportngin.com/survey/show/6743",
  };
})();
