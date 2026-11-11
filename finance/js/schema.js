// EVAA Finance Requests — form schema
// Schema definition + dynamic option loading from FinanceFormOptions SP list.
//
// The form *shape* (sections, field IDs, conditional visibility) is fixed here.
// The dropdown *contents* live in SharePoint list "FinanceFormOptions" and are
// loaded async by SCHEMA.load(). Call `await SCHEMA.load()` once on page boot
// before the first render — see app.js.
//
// To add a new sport / category / program etc., the treasurer edits the options in
// the console (Options tab). No code change.
//
// Conventions the form relies on (documented in the console's Options tab too):
//   • Categories starting "Revenue - " show only on Deposit Requests; categories
//     starting "Expense - " show only on Check Requests and Credit Card Use.
//     Anything without either prefix shows on every request type.
//   • Program rows carry a Sport (ForSport) and a Program type (ForProgramType).
//     The submitter picks Sport → Program type → Program; only programs matching
//     both show. Leadership / finance / portal admins see every sport's programs.

const SCHEMA = (() => {
  const W9 = "W-9 Submission";

  // ─── DYNAMIC ENUMS — baked defaults, refreshed by load() ──────────────────
  // Snapshot of FinanceFormOptions (2026-11-08). Kept populated so the form still
  // works if both the live feed and options.json are unreachable.
  const enums = {
    REQUEST_TYPES: ["Check Request", "Deposit Request", "Credit Card Use", W9],
    SPORTS: [
      "Baseball", "Basketball", "Cross Country Running", "Cross Country Skiing", "Football", "Lacrosse",
      "Soccer", "Softball", "Swimming", "Tennis", "Track", "Volleyball", "Wrestling",
      "General Fund", "Multi-Sport / Operations", "Other (note in details)",
    ],
    PROGRAM_TYPE: ["Traveling", "In-House", "Tournament", "Fundraising", "All", "N/A"],
    // [{ title, sport, programType }]
    PROGRAMS: [
      { title: "Basketball - Boys", sport: "Basketball", programType: "Traveling" },
      { title: "Basketball - Girls", sport: "Basketball", programType: "Traveling" },
      { title: "Basketball - Boys Tournament", sport: "Basketball", programType: "Tournament" },
      { title: "Basketball - Girls Tournament", sport: "Basketball", programType: "Tournament" },
      { title: "Basketball - Fundraising", sport: "Basketball", programType: "Fundraising" },
      { title: "Volleyball - JO Girls", sport: "Volleyball", programType: "Traveling" },
      { title: "Volleyball - JO Boys", sport: "Volleyball", programType: "Traveling" },
      { title: "Baseball - 12U Program", sport: "Baseball", programType: "Traveling" },
    ],
    TRAVELING_SUBTYPE: ["Fall Traveling", "Spring Traveling", "Summer Traveling", "Winter Traveling"],
    SEASON: [
      "This request is for the current season.",
      "This request is for the upcoming season.",
    ],
    EXPENSE_CATEGORY: [
      "Revenue - Player Registrations", "Revenue - Clinic Registrations", "Revenue - Concessions",
      "Revenue - Tournament Team Registrations", "Revenue - Tournament Gate Fee", "Revenue - Tournament Contributions",
      "Revenue - Parent Contributions", "Revenue - Fundraising", "Revenue - Contributions & Donations",
      "Revenue - EVAA Sportswear", "Revenue - Miscellaneous",
      "Expense - Advertising", "Expense - Accounting Fees", "Expense - Bank Service Fees", "Expense - Clinic Expenses",
      "Expense - Clinic Coaching Fees", "Expense - Coaching Fees", "Expense - Coaches Apparel",
      "Expense - Concession Supplies", "Expense - Contributions & Donations", "Expense - Equipment",
      "Expense - EVAA Sponsored Events", "Expense - EVAA Sportswear Apparel", "Expense - Facilities Rental",
      "Expense - Field Maintenance", "Expense - Field Supplies", "Expense - Fundraising", "Expense - Insurance",
      "Expense - Internet Costs", "Expense - League Entry Fees", "Expense - Licenses & Permits",
      "Expense - Marketing Costs", "Expense - Membership Dues", "Expense - Meeting Costs", "Expense - Office Supplies",
      "Expense - Other Costs", "Expense - Parent Fundraising Distribution", "Expense - Player Development",
      "Expense - Player Financial Support", "Expense - Player Uniforms", "Expense - PlayMetrics Fees",
      "Expense - Postage", "Expense - Professional Fees", "Expense - Printing", "Expense - Referees & Umpires",
      "Expense - Risk Management", "Expense - Scholarships", "Expense - Sports Engine Fees",
      "Expense - Storage Rental", "Expense - Tax Services", "Expense - Team Celebrations",
      "Expense - Tournament Team Registrations", "Expense - Tournament Other Costs", "Expense - Travel",
      "Expense - Trophies, Medals & Ribbons", "Expense - Tryout Costs", "Expense - Website Development",
    ],
    VENDOR_CARDINALITY: ["Single Vendor", "Multiple Vendors"],
  };

  // Viewer context. Leadership / finance / portal admins see every sport's programs.
  const context = { seeAllPrograms: false };
  function setContext(c) { Object.assign(context, c || {}); }

  // Live options feed — SAS-signed Power Automate trigger for "EVAA - Get Finance Options".
  // Reads FinanceFormOptions as web-admin and returns active rows, so console edits reach
  // every submitter (coaches don't have Board Portal SharePoint access). Read-only.
  const LIVE_OPTIONS_URL = "https://defaultb5897a1bb85b42bd8e619b021b67d2.ce.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/07/workflows/b0cda7b0ad7946f0a45789b21e72ce41/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=rDARDeSDKYu84yyc0tCBoDINSQ0y3hln2Q--mSI-5IM";

  // Apply a { OptionType: [...] } map onto enums, guarding empties so a
  // partial source never blanks a populated dropdown.
  function applyOptions(j) {
    if (j.Sport?.length)             enums.SPORTS             = j.Sport;
    if (j.ExpenseCategory?.length)   enums.EXPENSE_CATEGORY   = j.ExpenseCategory;
    if (j.ProgramType?.length)       enums.PROGRAM_TYPE       = j.ProgramType;
    if (Array.isArray(j.Program))    enums.PROGRAMS           = j.Program;
    if (j.TravelingSubtype?.length)  enums.TRAVELING_SUBTYPE  = j.TravelingSubtype;
    if (j.Season?.length)            enums.SEASON             = j.Season;
    if (j.VendorCardinality?.length) enums.VENDOR_CARDINALITY = j.VendorCardinality;
    if (j.RequestType?.length)       enums.REQUEST_TYPES      = j.RequestType;
  }

  // Turn flat option rows [{ title, type, order, sport, programType }] (live feed or
  // Graph) into the { OptionType: [...] } map applyOptions() expects.
  function rowsToMap(rows) {
    const sorted = rows.slice().sort((a, b) => ((a.order || 0) - (b.order || 0)) || String(a.title).localeCompare(String(b.title)));
    const m = {};
    for (const r of sorted) {
      if (!r.type || !r.title) continue;
      (m[r.type] = m[r.type] || []).push(
        r.type === "Program" ? { title: r.title, sport: r.sport || "", programType: r.programType || "" } : r.title
      );
    }
    if (!m.Program) m.Program = []; // an empty program list is valid (all deactivated)
    return m;
  }

  async function fetchWithTimeout(url, opts, ms) {
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), ms);
    try { return await fetch(url, { ...opts, signal: ctl.signal }); }
    finally { clearTimeout(t); }
  }

  // Async loader — populates `enums`, most-current source first:
  //   1) Live feed flow (no auth needed; reflects console edits immediately).
  //   2) options.json snapshot shipped with the site.
  //   3) Live Graph read of FinanceFormOptions — for anyone with site access.
  //   4) Baked defaults already in `enums`.
  async function load() {
    if (!LIVE_OPTIONS_URL.startsWith("TODO_")) {
      try {
        const resp = await fetchWithTimeout(LIVE_OPTIONS_URL, { cache: "no-store" }, 6000);
        if (resp.ok) {
          const j = await resp.json();
          if (j && Array.isArray(j.options) && j.options.length) { applyOptions(rowsToMap(j.options)); return; }
        }
      } catch (e) {
        console.warn("Live options feed unavailable, trying options.json:", e);
      }
    }

    try {
      const url = new URL("./options.json", document.baseURI).href + "?_=" + Date.now();
      const resp = await fetch(url, { cache: "no-store" });
      if (resp.ok) {
        const j = await resp.json();
        if (j && !j.error) { applyOptions(j); return; }
      }
    } catch (e) {
      console.warn("options.json unavailable, trying Graph:", e);
    }

    if (typeof GRAPH === "undefined" || !GRAPH.getFinanceFormOptions) return;
    try {
      const all = await GRAPH.getFinanceFormOptions();
      applyOptions(rowsToMap(all.filter((o) => o.active)));
    } catch (e) {
      console.warn("SCHEMA.load() Graph read failed — falling back to defaults:", e);
    }
  }

  // ─── Dependent option lists ───────────────────────────────────────────────
  function programsFor(values) {
    const kind = values.ProgramType;
    if (!kind) return [];
    return enums.PROGRAMS
      .filter((p) => p.programType === kind && (context.seeAllPrograms || p.sport === values.Sport))
      .map((p) => p.title);
  }

  function categoriesFor(values) {
    const all = enums.EXPENSE_CATEGORY;
    const t = values.RequestType;
    const isRev = (c) => /^revenue\s*-/i.test(c);
    const isExp = (c) => /^expense\s*-/i.test(c);
    if (t === "Deposit Request") return all.filter((c) => !isExp(c));
    if (t === "Check Request" || t === "Credit Card Use") return all.filter((c) => !isRev(c));
    return all;
  }

  const notW9 = (v) => v.RequestType !== W9;
  const isW9 = (v) => v.RequestType === W9;

  // ─── FORM SECTIONS ────────────────────────────────────────────────────────
  // Choice fields expose `options` (getter, static list) or `optionsFor(values)`
  // (list depends on other answers). `showIf(values)` hides a field/section;
  // `dependsOn` lists the fields that trigger a re-render when they change.
  const sections = [
    {
      id: "request_type",
      title: "Request type",
      fields: [
        {
          id: "RequestType",
          label: "What are you submitting?",
          type: "choice",
          required: true,
          get options() { return enums.REQUEST_TYPES; },
          help: "Check request, deposit, credit card use — or a vendor's W-9 form.",
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
      id: "w9",
      title: "W-9 details",
      showIf: isW9,
      dependsOn: ["RequestType"],
      help: "W-9s contain tax ID numbers, so they're stored in a separate, restricted folder that only the treasurer's billing group can open.",
      fields: [
        { id: "W9Name", label: "Vendor / payee name (as shown on the W-9)", type: "text", required: true },
        { id: "W9Sport", label: "Sport / Program this vendor works with (optional)", type: "choice", required: false,
          get options() { return enums.SPORTS; } },
        { id: "W9File", label: "Attach the W-9", type: "file", required: true,
          help: "PDF or a clear photo of the signed W-9." },
        { id: "W9Notes", label: "Notes (optional)", type: "textarea", required: false },
      ],
    },
    {
      id: "counterparty",
      title: "Vendor / Payer",
      showIf: notW9,
      dependsOn: ["RequestType"],
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
        {
          id: "DifferentMailingAddress",
          label: "Send the check to a different address",
          type: "boolean",
          required: false,
          when: { field: "RequestType", equals: "Check Request" },
          help: "Check this if the check should be mailed somewhere other than the vendor address.",
        },
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
      showIf: notW9,
      dependsOn: ["RequestType"],
      fields: [
        { id: "Sport", label: "Sport / Program", type: "choice", required: true, get options() { return enums.SPORTS; },
          help: "This is how the treasurer reconciles the request — pick the right sport even if the spend is shared." },
        { id: "ProgramType", label: "Program type", type: "choice", required: false,
          get options() { return enums.PROGRAM_TYPE; },
          help: "Traveling, in-house, tournament, etc. Picking one shows the matching programs for your sport." },
        {
          id: "Program",
          label: "Which program?",
          type: "choice",
          required: false,
          optionsFor: programsFor,
          showIf: (v) => programsFor(v).length > 0,
          dependsOn: ["Sport", "ProgramType"],
        },
        {
          id: "TravelingSubtype",
          label: "Which traveling season?",
          type: "choice",
          required: false,
          get options() { return enums.TRAVELING_SUBTYPE; },
          when: { field: "ProgramType", equals: "Traveling" },
        },
        { id: "Season", label: "What season is this request for?", type: "choice", required: false, get options() { return enums.SEASON; } },
        { id: "Category", label: "Expense / income category", type: "choice", required: true,
          optionsFor: categoriesFor, dependsOn: ["RequestType"],
          labelByType: {
            "Deposit Request": "Income category",
            "Check Request": "Expense category",
            "Credit Card Use": "Expense category",
          } },
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
  ];

  // ─── EVALUATOR ────────────────────────────────────────────────────────────
  function isVisible(node, values) {
    if (typeof node.showIf === "function" && !node.showIf(values)) return false;
    const cond = node.when;
    if (!cond) return true;
    const actual = values[cond.field];
    if (Array.isArray(cond.equals)) return cond.equals.includes(actual);
    return actual === cond.equals;
  }

  function optionsFor(field, values) {
    if (typeof field.optionsFor === "function") return field.optionsFor(values);
    return field.options || [];
  }

  function labelFor(field, values) {
    if (field.labelByType && values && values.RequestType && field.labelByType[values.RequestType]) {
      return field.labelByType[values.RequestType];
    }
    return field.label;
  }

  function isRequired(field, values) {
    if (field.required) return true;
    const rw = field.requiredWhen;
    if (!rw) return false;
    const actual = values[rw.field];
    if (Array.isArray(rw.equals)) return rw.equals.includes(actual);
    return actual === rw.equals;
  }

  // Field IDs whose change should trigger a full re-render (some other field's
  // visibility, required-ness or option list depends on them).
  function gatingFieldIds() {
    const ids = new Set();
    const addNode = (n) => {
      if (n.when) ids.add(n.when.field);
      if (n.requiredWhen) ids.add(n.requiredWhen.field);
      (n.dependsOn || []).forEach((d) => ids.add(d));
    };
    for (const sec of sections) {
      addNode(sec);
      sec.fields.forEach(addNode);
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
    setContext,
    sections,
    isVisible,
    optionsFor,
    labelFor,
    isRequired,
    visibleFields,
    gatingFieldIds,
    enums,
    W9,
    sourceSurveyId: 6743,
    sourceUrl: "https://evaa.sportngin.com/survey/show/6743",
  };
})();
