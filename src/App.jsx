import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════════════
   RSSP Student Budget Tool — Standalone v8

   CHANGES from v7:
     - Furniture & household items default changed from $500 to $0
     - Excel export: removed Assumptions section (Living Expenses + Income)
       as these are duplicated in the Budget Summary
     - Excel export: removed Accommodation type line; instead appended
       "(single)" or "(sharing)" after Accommodation in Budget Summary
     - Excel export: removed all "Weekly" / "Amount" column C headers;
       added subtitle line "All amounts are WEEKLY and are in A$"
     - Excel export: column A width reduced to 5; section headings bleed
       across into empty column B cells
     - Assumptions section: Living Expenses heading → "Living Expenses (Weekly)"

   PURPOSE:  Normalised weekly income-vs-expenses view for RSSP students.
   PERSISTENCE: Uses window.storage API.
   CALC LOGIC: Identical to RSSP Budgeting Tool v14 (see v6 header for detail).

   DATA SOURCES:
     Youth Allowance max rate: $677.20/fn — Services Australia
     Income test: free area $539, taper1 $646, 50c/60c — Services Australia
     Rent Assistance: indexed 20 Mar 2026 — DVA / Services Australia
     SSL: $1,349/semester — Services Australia
     Refugee Student Loan: Skill Path / Spark Finance
     Utility estimates: AER reference bills, ABS household expenditure survey
   ═══════════════════════════════════════════════════════════════════════ */

const STORAGE_KEY = "rssp-budget-assumptions";

const storage = {
  get: async (key) => {
    if (window.storage) return window.storage.get(key);
    const value = localStorage.getItem(key);
    return value ? { value } : null;
  },
  set: async (key, value) => {
    if (window.storage) return window.storage.set(key, value);
    localStorage.setItem(key, value);
  },
  delete: async (key) => {
    if (window.storage) return window.storage.delete(key);
    localStorage.removeItem(key);
  },
};

/* Living expense categories */
const LC_LABELS = {
  accom: "Accommodation",
  utilities: "Utilities (Electricity, Gas, Water)",
  transport: "Transport",
  food: "Food",
  personal: "Personal",
  clothing: "Clothing",
  entertainment: "Entertainment",
  other: "Other",
};

/* Hints shown below specific living cost input fields */
const LC_HINTS = {
  personal: "(eg. gym membership, haircut, cosmetics, mobile phone)",
  entertainment: "(eg. cinema, live music/sport, sporting clubs, night out)",
  other: "(eg. study materials, medical expenses)",
};

/* Rent Assistance rates — indexed 20 March 2026 */
const RA_TABLE = [
  { key: "single", situation: "Single",        threshold: 154.80, max: 219.40, get ceiling() { return +(this.threshold + this.max / 0.75).toFixed(2); } },
  { key: "sharer", situation: "Single, sharer", threshold: 154.80, max: 146.27, get ceiling() { return +(this.threshold + this.max / 0.75).toFixed(2); } },
];

/* ═══ UTILITY ESTIMATION DATA ═══ */
const ACCOM_TYPES = [
  { key: "student", label: "Purpose-built student accommodation", bundled: true },
  { key: "sharehouse", label: "Sharehouse (3–4 people)" },
  { key: "solo", label: "Solo rental (1-bed apartment)" },
  { key: "family", label: "Living with family / homestay", bundled: true },
];

const CLIMATE_ZONES = [
  { key: "warm", label: "Warm", cities: "Brisbane, Darwin, Cairns, Townsville" },
  { key: "mild", label: "Mild", cities: "Sydney, Perth, Adelaide, Gold Coast" },
  { key: "cool", label: "Cool", cities: "Melbourne, Canberra, Hobart, Geelong" },
];

const UTIL_ELECTRICITY = { sharehouse: 14, solo: 28 };
const UTIL_GAS = {
  sharehouse: { warm: 0, mild: 5, cool: 8 },
  solo:       { warm: 0, mild: 8, cool: 15 },
};
const UTIL_WATER = { sharehouse: 3, solo: 8 };
const UTIL_INTERNET = { sharehouse: 6, solo: 20 };

function getUtilityEstimate(accomType, climateZone) {
  const accom = ACCOM_TYPES.find(a => a.key === accomType);
  if (!accom || accom.bundled) {
    return { electricity: 0, gas: 0, water: 0, internet: 0, total: 0, bundled: true };
  }
  const electricity = UTIL_ELECTRICITY[accomType] || 0;
  const gas = (UTIL_GAS[accomType] && UTIL_GAS[accomType][climateZone]) || 0;
  const water = UTIL_WATER[accomType] || 0;
  const internet = UTIL_INTERNET[accomType] || 0;
  const total = electricity + gas + water + internet;
  return { electricity, gas, water, internet, total, bundled: false };
}

/* ═══ UTILITY DETAIL CONTENT ═══ */
const UTIL_INFO = {
  electricity: {
    title: "Electricity",
    icon: "⚡",
    description: "Electricity is typically the largest utility expense for students. Your usage depends on whether you have electric heating/cooling, how many appliances you use, and the size of your dwelling.",
    details: [
      "Average retail rate in Australia is approximately 30c per kWh plus a daily supply charge of around $1/day.",
      "A 1-bedroom apartment typically uses 3,500–5,000 kWh per year. In a sharehouse, your share is roughly 1,500–2,500 kWh per year.",
      "Air conditioning in summer (or electric heating in winter) is the biggest single driver of electricity costs.",
      "LED lighting, switching off standby appliances, and using cold-water washing can meaningfully reduce your bill.",
    ],
    tipToSave: "Compare energy plans using the government's free Energy Made Easy tool (energymadeeasy.gov.au). Switching retailers can save $200–400/year.",
  },
  gas: {
    title: "Gas",
    icon: "🔥",
    description: "Gas costs vary significantly depending on where you live in Australia. In warmer climates (Brisbane, Darwin), many homes have no gas connection at all. In cooler climates (Melbourne, Canberra, Hobart), gas heating is common and can add substantially to winter bills.",
    details: [
      "In Melbourne and Canberra, gas heating can add $10–20/week in peak winter months, averaging $8–12/week over the year per person.",
      "In Sydney and Perth, gas usage is moderate — mainly for hot water and cooking — averaging $5–8/week per person.",
      "In Brisbane and Darwin, most student accommodation has no gas connection. Hot water and cooking are electric.",
      "If your accommodation has a gas connection, you will pay a daily supply charge (~60–90c/day) regardless of usage.",
    ],
    tipToSave: "If you have gas heating, set the thermostat to 18–20°C. Each degree above 20°C adds roughly 10% to your heating bill. Draught-proofing doors and windows makes a big difference.",
  },
  water: {
    title: "Water",
    icon: "💧",
    description: "Water billing varies by state and lease type. In many apartment complexes and purpose-built student accommodation, water is included in rent. Where it is separately metered, a single person's usage is typically modest.",
    details: [
      "Typical single-person water usage costs $5–10/week where separately metered.",
      "In Victoria, landlords can only charge tenants for water if the property meets water efficiency standards (low-flow showerheads, dual-flush toilets).",
      "In NSW and Queensland, landlords can generally pass on water usage charges (but not the fixed service charge).",
      "Most purpose-built student accommodation includes water in the rent — check your lease.",
    ],
    tipToSave: "Shorter showers are the single biggest water-saving action. A 4-minute shower uses roughly 36 litres vs 90+ litres for a 10-minute shower.",
  },
  internet: {
    title: "Internet",
    icon: "📶",
    description: "Internet is relatively consistent in price across Australia. The main variable is how many people you split the plan with. Purpose-built student accommodation usually includes WiFi in the rent.",
    details: [
      "A standard NBN 50 plan (suitable for most students) costs $70–90/month.",
      "In a sharehouse of 3–4 people, your share is roughly $5–10/week.",
      "Living alone, you pay the full plan cost: approximately $17–22/week.",
      "Some newer apartment complexes include internet in body corporate fees or strata, so check before signing up for a separate plan.",
    ],
    tipToSave: "If you're in a sharehouse, NBN 50 is usually sufficient for 3–4 people streaming and studying simultaneously. Avoid signing up for NBN 100 unless someone needs it for work.",
  },
};

const C = { navy: "#385592", coral: "#de5240", cyan: "#cdf0f1", teal: "#2dcd9e" };

const DEFAULTS = () => ({
  livingCosts: Object.fromEntries(Object.keys(LC_LABELS).map(k => [k, 0])),
  hoursPerWeek: 0, hourlyWage: 0, raType: "single", otherNote: "",
});

/* ═══ CALCULATION FUNCTIONS ═══ */
function calcRA(fnRent, threshold, maxRA, taper) {
  if (fnRent <= threshold) return 0;
  return Math.min(Math.round((fnRent - threshold) * taper * 100) / 100, maxRA);
}

function calcIncomeTestReduction(fnWages, freeArea, taper1End, taper1Rate, taper2Rate) {
  if (fnWages <= freeArea) return 0;
  if (fnWages <= taper1End) return Math.round((fnWages - freeArea) * taper1Rate * 100) / 100;
  return Math.round(((taper1End - freeArea) * taper1Rate + (fnWages - taper1End) * taper2Rate) * 100) / 100;
}

/* ═══ FORMATTING ═══ */
const fmt = v => {
  if (v == null) return "-";
  return v < 0
    ? `($${Math.abs(Math.round(v)).toLocaleString("en-AU")})`
    : `$${Math.round(v).toLocaleString("en-AU")}`;
};
const fmt2 = v => {
  if (v == null) return "-";
  const abs = Math.abs(v).toFixed(2);
  return v < 0 ? `($${abs})` : `$${abs}`;
};

/* ═══ UI COMPONENTS ═══ */
const Inp = ({ label, value, onChange, min, max, step, note, warn, dollar, disabled, placeholder }) => (
  <div className="flex flex-col gap-0.5">
    <label className="text-xs font-medium" style={{ color: C.navy }}>{label}</label>
    <div className={dollar ? "relative" : ""}>
      {dollar && <span className="absolute left-2 top-1/2 -translate-y-1/2 text-sm text-gray-400">$</span>}
      <input type="number" value={value === 0 && placeholder ? "" : value}
        onChange={e => { const raw = e.target.value; onChange(raw === "" ? 0 : parseFloat(raw) || 0); }}
        placeholder={placeholder} min={min} max={max} step={step || 1} disabled={disabled}
        className={`border rounded py-1.5 text-sm bg-white focus:outline-none focus:ring-2 w-full ${dollar ? "pl-6 pr-2" : "px-2"} ${disabled ? "opacity-50" : ""}`}
        style={{ borderColor: warn ? C.coral : "#d1d5db" }} />
    </div>
    {note && <span className={`text-xs ${warn ? "font-medium" : ""}`} style={{ color: warn ? C.coral : "#9ca3af" }}>{note}</span>}
  </div>
);
const Section = ({ title, children }) => (
  <div className="mb-6">
    <div className="flex items-center gap-2 mb-3 pb-1.5" style={{ borderBottom: `2px solid ${C.teal}` }}>
      <h3 className="text-xs font-bold uppercase tracking-wide" style={{ color: C.navy }}>{title}</h3>
    </div>
    {children}
  </div>
);
const SubHeading = ({ children }) => (
  <p className="text-xs font-bold mb-2" style={{ color: C.navy }}>{children}</p>
);
const GreyNote = ({ children }) => (
  <div className="text-xs mb-3" style={{ color: "#6b7280" }}>{children}</div>
);
const SectionDivider = ({ title }) => (
  <div className="mb-4 -mx-5 px-5 py-3" style={{ backgroundColor: `${C.navy}10`, borderTop: `2px solid ${C.navy}30`, borderBottom: `2px solid ${C.navy}30` }}>
    <h2 className="text-sm font-bold uppercase tracking-wide" style={{ color: C.navy }}>{title}</h2>
  </div>
);
const FN = () => <span style={{ fontWeight: 700, textDecoration: "underline" }}>Fortnightly</span>;

const LOGO = "data:image/png;base64,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";

/* ═══ UTILITY INFO CARD ═══ */
const UtilityCard = ({ info, estimate, isBundled }) => {
  const [open, setOpen] = useState(false);
  return (
    <div className="rounded-lg border overflow-hidden mb-4" style={{ borderColor: "#e5e7eb" }}>
      <button onClick={() => setOpen(!open)} className="w-full flex items-center justify-between p-4 text-left transition hover:bg-gray-50">
        <div className="flex items-center gap-3">
          <span className="text-2xl">{info.icon}</span>
          <div>
            <h4 className="text-sm font-bold" style={{ color: C.navy }}>{info.title}</h4>
            <p className="text-xs" style={{ color: "#6b7280" }}>
              {isBundled ? "Typically included in your rent" : `Estimated: $${estimate}/week`}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {!isBundled && (
            <span className="font-mono font-bold text-sm px-3 py-1 rounded" style={{ backgroundColor: `${C.teal}20`, color: C.navy }}>
              ${estimate}/wk
            </span>
          )}
          {isBundled && (
            <span className="text-xs font-semibold px-3 py-1 rounded" style={{ backgroundColor: C.cyan, color: C.navy }}>
              $0/wk
            </span>
          )}
          <svg className={`w-4 h-4 transition-transform ${open ? "rotate-180" : ""}`} style={{ color: C.navy }} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
        </div>
      </button>
      {open && (
        <div className="px-4 pb-4 border-t" style={{ borderColor: "#f3f4f6" }}>
          <p className="text-xs mt-3 mb-3" style={{ color: "#374151" }}>{info.description}</p>
          <div className="space-y-2 mb-3">
            {info.details.map((d, i) => (
              <div key={i} className="flex gap-2 text-xs" style={{ color: "#6b7280" }}>
                <span style={{ color: C.teal, flexShrink: 0 }}>•</span>
                <span>{d}</span>
              </div>
            ))}
          </div>
          <div className="p-3 rounded-lg text-xs" style={{ backgroundColor: "#f0fdf4", border: "1px solid #bbf7d0" }}>
            <span className="font-semibold" style={{ color: "#065f46" }}>Tip to save: </span>
            <span style={{ color: "#065f46" }}>{info.tipToSave}</span>
          </div>
        </div>
      )}
    </div>
  );
};

/* ═══ MAIN APPLICATION ═══ */
export default function App() {
  const [tab, setTab] = useState("budget");
  const [loaded, setLoaded] = useState(false);

  const [livingCosts, setLivingCosts] = useState(DEFAULTS().livingCosts);
  const [hoursPerWeek, setHoursPerWeek] = useState(0);
  const [hourlyWage, setHourlyWage] = useState(0);
  const [raType, setRaType] = useState("single");
  const [otherNote, setOtherNote] = useState("");

  const [utilAccomType, setUtilAccomType] = useState("sharehouse");
  const [utilClimate, setUtilClimate] = useState("cool");

  /* Funding summary state (new in v7) */
  const [furnitureCost, setFurnitureCost] = useState(0);
  const [estimatedSavingsInput, setEstimatedSavingsInput] = useState(0);

  const yaMaxRate = 677.20;
  const freeArea = 539;
  const taper1End = 646;
  const taper1Rate = 0.50;
  const taper2Rate = 0.60;

  const raRow = RA_TABLE.find(r => r.key === raType) || RA_TABLE[0];

  /* ── Load ── */
  useEffect(() => {
    (async () => {
      try {
        const result = await storage.get(STORAGE_KEY);
        if (result && result.value) {
          const s = JSON.parse(result.value);
          if (s.livingCosts) setLivingCosts(prev => ({ ...DEFAULTS().livingCosts, ...s.livingCosts }));
          if (s.hoursPerWeek != null) setHoursPerWeek(s.hoursPerWeek);
          if (s.hourlyWage != null) setHourlyWage(s.hourlyWage);
          if (s.raType) setRaType(s.raType);
          if (s.otherNote != null) setOtherNote(s.otherNote);
          if (s.utilAccomType) setUtilAccomType(s.utilAccomType);
          if (s.utilClimate) setUtilClimate(s.utilClimate);
          if (s.furnitureCost != null) setFurnitureCost(s.furnitureCost);
          if (s.estimatedSavingsInput != null) setEstimatedSavingsInput(s.estimatedSavingsInput);
        }
      } catch (_) {}
      setLoaded(true);
    })();
  }, []);

  /* ── Save ── */
  const save = useCallback(async (lc, hpw, hw, rt, on, uat, uc, fc, esi) => {
    try {
      await storage.set(STORAGE_KEY, JSON.stringify({
        livingCosts: lc, hoursPerWeek: hpw, hourlyWage: hw, raType: rt, otherNote: on,
        utilAccomType: uat, utilClimate: uc, furnitureCost: fc, estimatedSavingsInput: esi,
      }));
    } catch (_) {}
  }, []);

  useEffect(() => {
    if (loaded) save(livingCosts, hoursPerWeek, hourlyWage, raType, otherNote, utilAccomType, utilClimate, furnitureCost, estimatedSavingsInput);
  }, [livingCosts, hoursPerWeek, hourlyWage, raType, otherNote, utilAccomType, utilClimate, furnitureCost, estimatedSavingsInput, loaded, save]);

  /* ── Reset ── */
  const resetAll = async () => {
    const d = DEFAULTS();
    setLivingCosts(d.livingCosts);
    setHoursPerWeek(d.hoursPerWeek);
    setHourlyWage(d.hourlyWage);
    setRaType(d.raType);
    setOtherNote(d.otherNote);
    setFurnitureCost(0);
    setEstimatedSavingsInput(0);
    try { await storage.delete(STORAGE_KEY); } catch (_) {}
  };

  /* ── Derived ── */
  const upLC = (k, v) => setLivingCosts(prev => ({ ...prev, [k]: v }));
  const weeklyAccom = livingCosts.accom || 0;
  const weeklyWages = hoursPerWeek * hourlyWage;
  const fnAccomDisplay = weeklyAccom * 2;
  const currentRA = calcRA(fnAccomDisplay, raRow.threshold, raRow.max, 0.75);
  const rentalBond = weeklyAccom * 4;
  const rentInAdvance = weeklyAccom * 2;

  const utilEstimate = useMemo(() => getUtilityEstimate(utilAccomType, utilClimate), [utilAccomType, utilClimate]);

  const applyUtilityEstimate = () => {
    upLC("utilities", utilEstimate.total);
    setTab("budget");
  };

  /* ── Budget calculation ── */
  const budget = useMemo(() => {
    const fnWages = weeklyWages * 2;
    const fnAccom = weeklyAccom * 2;
    const craPerFn = calcRA(fnAccom, raRow.threshold, raRow.max, 0.75);
    const combinedMaxFn = yaMaxRate + craPerFn;
    const incTestRedFn = calcIncomeTestReduction(fnWages, freeArea, taper1End, taper1Rate, taper2Rate);
    const netGovPerFn = Math.max(0, combinedMaxFn - incTestRedFn);
    const weeklyGov = netGovPerFn / 2;
    const weeklyYAMax = yaMaxRate / 2;
    const weeklyRAMax = craPerFn / 2;
    const weeklyReduction = Math.min(incTestRedFn / 2, (yaMaxRate + craPerFn) / 2);
    const expenseItems = Object.entries(LC_LABELS).map(([k, label]) => ({ key: k, label, amount: livingCosts[k] || 0 }));
    const totalExpense = expenseItems.reduce((s, item) => s + item.amount, 0);
    const totalIncome = weeklyWages + weeklyGov;
    const net = totalIncome - totalExpense;
    return { weeklyGov, weeklyYAMax, weeklyRAMax, weeklyReduction, expenseItems, totalExpense, totalIncome, net };
  }, [livingCosts, hoursPerWeek, hourlyWage, raRow, weeklyWages, weeklyAccom]);

  /* ── CSV Download ── */
  const downloadXLSX = () => {

    /* Helper: sanitize unicode chars for Excel compatibility */
    const san = s => String(s ?? "").replace(/\u2014/g, "-").replace(/\u2013/g, "-").replace(/\u2018|\u2019/g, "'").replace(/\u201c|\u201d/g, '"');

    /* ── Build rows ── */
    const data = [];
    const push = (...args) => data.push(args.map(a => typeof a === "string" ? san(a) : a));
    const blank = () => data.push([]);

    /* Header */
    push("Student Budget Summary");
    push("All amounts are WEEKLY and are in A$");
    push("Export Date", new Date().toLocaleDateString("en-AU"));
    blank();

    /* Budget Summary */
    push("BUDGET SUMMARY");
    blank();
    push("", "Income", "");
    push("", "  Part-time wages (" + hoursPerWeek + " hrs x $" + hourlyWage + "/hr)", Math.round(weeklyWages));
    push("", "  Government payment (net)", Math.round(budget.weeklyGov));
    push("", "    Youth Allowance (max)", Math.round(budget.weeklyYAMax));
    push("", "    Rent Assistance (max)", Math.round(budget.weeklyRAMax));
    push("", "    Less: Income test reduction", budget.weeklyReduction > 0 ? -Math.round(budget.weeklyReduction) : 0);
    push("", "Total Weekly Income", Math.round(budget.totalIncome));
    blank();
    push("", "Living Expenses", "");
    const accomSuffix = raType === "single" ? " (single)" : " (sharing)";
    budget.expenseItems.forEach(item => {
      const label = item.key === "accom" ? san(item.label) + accomSuffix : san(item.label);
      push("", "  " + label, item.amount);
    });
    push("", "Total Weekly Living Expenses", budget.totalExpense);
    blank();
    push("", "Net Weekly Amount", Math.round(budget.net));
    blank();

    /* Funding Summary */
    push("FUNDING SUMMARY");
    blank();
    push("", "Upfront Costs", "");
    push("", "  Rental bond (4 weeks rent)", rentalBond);
    push("", "  Rent in advance (2 weeks)", rentInAdvance);
    push("", "  Furniture & household items", furnitureCost);
    push("", "Total Upfront Costs", totalUpfrontCosts);
    blank();
    push("", "Less: Estimated savings at move-out", estimatedSavingsInput > 0 ? -estimatedSavingsInput : 0);
    blank();
    push("", fundingGap > 0 ? "Estimated Funding Gap" : "Estimated Surplus After Upfront Costs", Math.round(fundingGap));

    /* ── Create workbook ── */
    const ws = XLSX.utils.aoa_to_sheet(data);

    /* Column widths */
    ws["!cols"] = [
      { wch: 5 },   /* A: narrow — section headings bleed into B */
      { wch: 42 },  /* B: item labels */
      { wch: 14 },  /* C: amounts */
    ];

    /* Number format for currency columns — apply to column C */
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: 2 });
      const cell = ws[addr];
      if (cell && typeof cell.v === "number") {
        cell.z = '"$"#,##0';
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Budget Summary");
    XLSX.writeFile(wb, "RSSP_Budget_Summary.xlsx");
  };

  const isDeficit = budget.net < 0;
  const lcTotal = Object.keys(LC_LABELS).reduce((s, k) => s + (livingCosts[k] || 0), 0);

  /* Funding summary calculations */
  const totalUpfrontCosts = rentalBond + rentInAdvance + furnitureCost;
  const fundingGap = totalUpfrontCosts - estimatedSavingsInput;
  const TABS = [["budget", "Budget"], ["utilities", "Utilities Guide"], ["funding", "Funding"]];

  return (
    <div className="min-h-screen" style={{ backgroundColor: "#f8f9fb" }}>
      <div className="max-w-4xl mx-auto p-4">
        {/* Header */}
        <div className="rounded-lg p-4 mb-4 flex items-center justify-between" style={{ background: `linear-gradient(135deg, ${C.navy} 0%, #4a6aaa 100%)` }}>
          <div>
            <h1 className="text-lg font-bold text-white">RSSP Student Budget Tool</h1>
            <p className="text-xs text-white opacity-70">All currency values are in Australian dollars</p>
          </div>
          <img src={LOGO} alt="Skill Path" style={{ height: 48, objectFit: "contain" }} />
        </div>

        {/* Tab bar + buttons */}
        <div className="flex items-center justify-between mb-4 border-b" style={{ borderColor: "#e5e7eb" }}>
          <div className="flex gap-1">
            {TABS.map(([k, l]) => (
              <button key={k} onClick={() => setTab(k)}
                className="px-4 py-2.5 text-sm font-medium border-b-2 -mb-px transition whitespace-nowrap"
                style={{ borderColor: tab === k ? C.teal : "transparent", color: tab === k ? C.navy : "#9ca3af", fontWeight: tab === k ? 700 : 500 }}>
                {l}
              </button>
            ))}
          </div>
          <div className="flex gap-2 mb-1">
            <button onClick={downloadXLSX} className="px-3 py-1.5 text-xs rounded transition"
              style={{ backgroundColor: C.navy, color: "white", border: `1px solid ${C.navy}` }}>
              Download to Spreadsheet
            </button>
            <button onClick={resetAll} className="px-3 py-1.5 text-xs rounded transition"
              style={{ backgroundColor: "#f3f4f6", color: C.navy, border: "1px solid #e5e7eb" }}>
              Reset Assumptions
            </button>
          </div>
        </div>

        {/* ═══ TAB 1: BUDGET ═══ */}
        {tab === "budget" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Assumptions" />

            <Section title="Living Expenses (Weekly)">
              <GreyNote>
                You can compare your assumed living expenses to those in the Government's cost of living calculator for students which is available{" "}
                <a href="https://costofliving.studyaustralia.gov.au" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.
                {" "}For help estimating your utility costs, see the <button onClick={() => setTab("utilities")} className="font-semibold underline" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Utilities Guide</button> tab.
              </GreyNote>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-3">
                {Object.entries(LC_LABELS).map(([k, label]) => (
                  <Inp key={k} label={label} value={livingCosts[k]} onChange={v => upLC(k, v)} min={0} step={1} dollar placeholder="0" note={LC_HINTS[k]} />
                ))}
              </div>
              <div className="flex justify-between items-center p-2 rounded text-sm" style={{ backgroundColor: C.cyan }}>
                <span className="font-medium" style={{ color: C.navy }}>Total Weekly</span>
                <span className="font-mono font-semibold" style={{ color: C.navy }}>${lcTotal.toLocaleString()}/wk</span>
              </div>
            </Section>

            <Section title="Income Sources">
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-3">
                <Inp label="Part time work — hours/week" value={hoursPerWeek} onChange={v => setHoursPerWeek(v)} min={0} max={48} placeholder="0" />
                <Inp label="Hourly wage" value={hourlyWage} onChange={v => setHourlyWage(v)} step={0.5} dollar placeholder="0" />
              </div>
              <p className="text-xs" style={{ color: "#6b7280" }}>
                If you start working or your income from work changes you will need to notify Centrelink. See{" "}
                <a href="https://www.servicesaustralia.gov.au/when-to-report-your-income-to-centrelink?context=43916" target="_blank" rel="noopener noreferrer"
                  style={{ color: C.navy, textDecoration: "underline", fontWeight: 700 }}>
                  this page
                </a>{" "}for how to report changes in income to Centrelink.
              </p>
            </Section>

            <Section title="Youth Allowance">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">Youth Allowance is an Australian Government payment which students receive with the amount paid depending on their situation. This budget model assumes you are between the ages of 15–25, single with no children. If you are older than 25 you will likely be eligible for Austudy payments which are similar.</p>
                <p className="mb-2">The Youth Allowance payments reduce in accordance with a <span style={{ textDecoration: "underline" }}>personal income test</span>. You can earn up to $539 per fortnight before your payment is affected (the "income free area"). For each dollar earned between $539 and $646, your combined payment reduces by 50 cents. For each dollar above $646, your combined payment reduces by 60 cents. The reduction is applied to your total payment from the Government (including any Rent Assistance you may be entitled to — see below).</p>
                <p className="mb-2">An "Income Bank" allows you to accumulate unused income free area credits in low-income fortnights to offset higher-income fortnights.</p>
                <p>More info:{" "}<a href="https://www.servicesaustralia.gov.au/youth-allowance" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>Youth Allowance</a></p>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <Inp label="Youth Allowance Max (fortnightly)" value={yaMaxRate} onChange={() => {}} step={0.1} dollar disabled note="Single, no children, 18+, away from home" />
              </div>
            </Section>

            <Section title="Rent Assistance">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <SubHeading>How Rent Assistance Works</SubHeading>
                <p className="mb-1">Rent Assistance is an additional government payment for eligible students.</p>
                <p className="mb-1">Rent Assistance is calculated based on the rent you pay. For every $1 of fortnightly rent you pay above the relevant rent threshold, you receive 75 cents in Rent Assistance, up to a maximum amount.</p>
                <p className="mb-1">Rent Assistance is not subject to a separate income test. It is added to your Youth Allowance payment to form a combined maximum rate. The personal income test reduction is then applied to this combined total. This means your Rent Assistance is only affected once the income test reduction exceeds your base Youth Allowance amount.</p>
                <p className="mb-4">More info from the Australian Government{" "}<a href="https://www.servicesaustralia.gov.au/how-much-rent-assistance-you-can-get?context=22206" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.</p>

                <SubHeading>Rent Assistance and Purpose Built Student Accommodation (PBSA)</SubHeading>
                <p className="mb-1">
                  Students paying for accommodation in a PBSA property should be eligible for Rent Assistance, which will most likely be considered "Shared" accommodation.
                </p>
                <p className="mb-1">
                  If the property includes catering (regular meals), a portion of the total payment may need to be excluded so that only the rental amount is included. Centrelink uses a <strong>two-thirds rule</strong> to calculate the rent component when a student is paying "Board and Lodging" which includes regular meals.
                </p>
                <p>
                  You will need to speak to Centrelink to confirm the appropriate payments.
                </p>
              </div>

              <div className="mb-4">
                <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Your accommodation type</label>
                <div className="flex gap-3">
                  {RA_TABLE.map(row => (
                    <button key={row.key} onClick={() => setRaType(row.key)} className="px-4 py-2 rounded text-sm font-medium transition"
                      style={{ backgroundColor: raType === row.key ? C.navy : "#f3f4f6", color: raType === row.key ? "white" : C.navy, border: `2px solid ${raType === row.key ? C.navy : "#e5e7eb"}` }}>
                      {row.situation}
                    </button>
                  ))}
                </div>
              </div>
              <div className="overflow-x-auto mb-4">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.cyan }}>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>If you're</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Your <FN /> rent is more than</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>To get the maximum payment your <FN /> rent is at least</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>The maximum <FN /> payment is</th>
                    </tr>
                  </thead>
                  <tbody>
                    {RA_TABLE.map(row => (
                      <tr key={row.key} style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: raType === row.key ? `${C.teal}15` : "transparent" }}>
                        <td className="p-2" style={{ fontWeight: raType === row.key ? 600 : 400 }}>{row.situation}</td>
                        <td className="p-2 font-mono">{fmt2(row.threshold)}</td>
                        <td className="p-2 font-mono">{fmt2(row.ceiling)}</td>
                        <td className="p-2 font-mono font-semibold">{fmt2(row.max)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="p-3 rounded-lg mb-4" style={{ backgroundColor: `${C.teal}10`, border: `1px solid ${C.teal}` }}>
                <h4 className="text-xs font-bold mb-1" style={{ color: C.navy }}>Your Estimated Maximum Rent Assistance <span className="font-normal">(before income test)</span></h4>
                <p className="text-xs mb-3" style={{ color: "#6b7280" }}>
                  Your estimated rent assistance payments based on the rent and accommodation type is shown below. This is <span style={{ fontWeight: 700, textDecoration: "underline" }}>BEFORE</span> any reduction due to the income test which is calculated in the budget summary below.
                </p>
                <div className="grid grid-cols-2 gap-x-8 gap-y-2 text-xs max-w-md">
                  <div><span className="text-gray-500">Weekly rent</span><div className="font-mono font-semibold" style={{ color: C.navy }}>{fmt2(weeklyAccom)}</div></div>
                  <div><span className="text-gray-500">Fortnightly rent</span><div className="font-mono font-semibold" style={{ color: C.navy }}>{fmt2(fnAccomDisplay)}</div></div>
                  <div><span className="text-gray-500">Rent Assistance (weekly)</span><div className="font-mono font-semibold" style={{ color: C.teal }}>{fmt2(currentRA / 2)}</div></div>
                  <div><span className="text-gray-500">Rent Assistance (fortnightly)</span><div className="font-mono font-semibold" style={{ color: C.teal }}>{fmt2(currentRA)}</div></div>
                </div>
                {fnAccomDisplay <= raRow.threshold && weeklyAccom > 0 && (
                  <p className="text-xs mt-2" style={{ color: C.coral }}>Your fortnightly rent of {fmt2(fnAccomDisplay)} is below the threshold of {fmt2(raRow.threshold)}. You do not qualify for Rent Assistance.</p>
                )}
                {weeklyAccom === 0 && (
                  <p className="text-xs mt-2" style={{ color: "#9ca3af" }}>Enter your weekly accommodation cost above to calculate your Rent Assistance entitlement.</p>
                )}
              </div>
            </Section>

            {/* ═══ BUDGET SUMMARY ═══ */}
            <SectionDivider title="Budget Summary" />
            <GreyNote>This shows your weekly ongoing budget based on the assumptions above without any university contribution or one-off income sources.</GreyNote>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead><tr style={{ backgroundColor: C.navy }}>
                  <th className="text-left px-4 py-2.5 font-semibold text-white">Item</th>
                  <th className="text-right px-4 py-2.5 font-semibold text-white w-32">Weekly Amount</th>
                </tr></thead>
                <tbody>
                  <tr><td colSpan={2} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.teal}15`, color: C.navy }}>Income</td></tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Part-time wages ({hoursPerWeek} hrs × ${hourlyWage}/hr)</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(Math.round(weeklyWages))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Government payment (net)</td>
                    <td className="px-4 py-2 text-right font-mono font-semibold" style={{ color: C.navy }}>{fmt(Math.round(budget.weeklyGov))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: "#6b7280", paddingLeft: 44 }}>Youth Allowance (max)</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: "#6b7280" }}>{fmt(Math.round(budget.weeklyYAMax))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: "#6b7280", paddingLeft: 44 }}>Rent Assistance (max)</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: "#6b7280" }}>{fmt(Math.round(budget.weeklyRAMax))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: budget.weeklyReduction > 0 ? C.coral : "#6b7280", paddingLeft: 44 }}>Less: Income test reduction</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: budget.weeklyReduction > 0 ? C.coral : "#6b7280" }}>
                      {budget.weeklyReduction > 0 ? `($${Math.round(budget.weeklyReduction).toLocaleString("en-AU")})` : fmt(0)}
                    </td>
                  </tr>
                  <tr style={{ borderTop: "2px solid #e5e7eb", backgroundColor: "#f0fdf4" }}>
                    <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total Weekly Income</td>
                    <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(Math.round(budget.totalIncome))}</td>
                  </tr>
                  <tr><td colSpan={2} className="py-1"></td></tr>
                  <tr><td colSpan={2} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.coral}10`, color: C.navy }}>Living Expenses</td></tr>
                  {budget.expenseItems.map((item, i) => (
                    <tr key={`exp-${i}`} style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>
                        {item.label}
                      </td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(item.amount)}</td>
                    </tr>
                  ))}
                  <tr style={{ borderTop: "2px solid #e5e7eb", backgroundColor: "#fef2f2" }}>
                    <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total Weekly Living Expenses</td>
                    <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(budget.totalExpense)}</td>
                  </tr>
                  <tr><td colSpan={2} className="py-1"></td></tr>
                  <tr style={{ backgroundColor: isDeficit ? `${C.coral}15` : `${C.teal}15`, borderTop: `3px solid ${isDeficit ? C.coral : C.teal}` }}>
                    <td className="px-4 py-3 font-bold" style={{ color: isDeficit ? C.coral : "#065f46" }}>Net Weekly Amount</td>
                    <td className="px-4 py-3 text-right font-mono font-bold text-base" style={{ color: isDeficit ? C.coral : "#065f46" }}>{fmt(Math.round(budget.net))}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="mt-5 p-4 rounded-lg" style={{ backgroundColor: isDeficit ? "#fef2f2" : "#f0fdf4", border: `2px solid ${isDeficit ? C.coral : C.teal}` }}>
              <p className="text-sm font-bold" style={{ color: isDeficit ? C.coral : "#065f46" }}>
                Based on your assumptions you are spending {isDeficit ? "more" : "less"} than you earn by ${Math.abs(Math.round(budget.net)).toLocaleString("en-AU")} per week.
              </p>
            </div>
            <div className="mt-6 p-4 rounded text-xs" style={{ backgroundColor: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d" }}>
              <strong>Disclaimer:</strong> This budgeting tool is provided to students in the RSSP as a guide to assist them in understanding the potential costs associated with studying in Australia. The model is indicative only and has not taken into account the student's personal situation.
            </div>
          </div>
        )}

        {/* ═══ TAB 2: UTILITIES GUIDE ═══ */}
        {tab === "utilities" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Utilities Guide" />
            <GreyNote>
              This guide helps you estimate your weekly utility costs based on your accommodation type and where in Australia you are living. Utility expenses can vary significantly and will depend heavily on climate and whether the accommodation uses gas and electricity. You can use this tool to generate an estimate, then apply it to your budget.
            </GreyNote>

            <Section title="Your Situation">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-4">
                <div>
                  <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Accommodation type</label>
                  <div className="flex flex-col gap-2">
                    {ACCOM_TYPES.map(a => (
                      <button key={a.key} onClick={() => setUtilAccomType(a.key)}
                        className="px-4 py-2.5 rounded text-xs font-medium transition text-left"
                        style={{
                          backgroundColor: utilAccomType === a.key ? C.navy : "#f3f4f6",
                          color: utilAccomType === a.key ? "white" : C.navy,
                          border: `2px solid ${utilAccomType === a.key ? C.navy : "#e5e7eb"}`,
                        }}>
                        {a.label}
                        {a.bundled && <span className="ml-2 opacity-70">(utilities usually included)</span>}
                      </button>
                    ))}
                  </div>
                </div>
                <div>
                  <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Climate zone</label>
                  <div className="flex flex-col gap-2">
                    {CLIMATE_ZONES.map(cz => (
                      <button key={cz.key} onClick={() => setUtilClimate(cz.key)}
                        className="px-4 py-2.5 rounded text-xs font-medium transition text-left"
                        style={{
                          backgroundColor: utilClimate === cz.key ? C.navy : "#f3f4f6",
                          color: utilClimate === cz.key ? "white" : C.navy,
                          border: `2px solid ${utilClimate === cz.key ? C.navy : "#e5e7eb"}`,
                        }}>
                        <span className="font-semibold">{cz.label}</span>
                        <span className="ml-2 opacity-70">— {cz.cities}</span>
                      </button>
                    ))}
                  </div>
                  <p className="text-xs mt-2" style={{ color: "#9ca3af" }}>Climate zone primarily affects gas/heating costs. Select the zone closest to your university city.</p>
                </div>
              </div>
            </Section>

            <Section title="Your Estimated Weekly Utility Costs">
              {utilEstimate.bundled ? (
                <div className="p-4 rounded-lg mb-4" style={{ backgroundColor: C.cyan, border: `1px solid ${C.navy}20` }}>
                  <p className="text-sm font-semibold" style={{ color: C.navy }}>$0/week — utilities are typically included in your rent</p>
                  <p className="text-xs mt-1" style={{ color: "#6b7280" }}>
                    {utilAccomType === "student"
                      ? "Purpose-built student accommodation almost always includes electricity, gas, water and internet in the weekly rent. Check your lease to confirm."
                      : "When living with family or in a homestay arrangement, utility costs are typically covered by the household. You may want to discuss contributing to household expenses."}
                  </p>
                </div>
              ) : (
                <>
                  <div className="overflow-x-auto mb-4">
                    <table className="w-full text-sm">
                      <thead>
                        <tr style={{ backgroundColor: C.navy }}>
                          <th className="text-left px-4 py-2 font-semibold text-white">Utility</th>
                          <th className="text-right px-4 py-2 font-semibold text-white w-28">Weekly Est.</th>
                          <th className="text-right px-4 py-2 font-semibold text-white w-28">Annual Est.</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[
                          { label: "⚡ Electricity", val: utilEstimate.electricity },
                          { label: "🔥 Gas", val: utilEstimate.gas },
                          { label: "💧 Water", val: utilEstimate.water },
                          { label: "📶 Internet", val: utilEstimate.internet },
                        ].map((row, i) => (
                          <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                            <td className="px-4 py-2" style={{ color: C.navy }}>{row.label}</td>
                            <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>${row.val}/wk</td>
                            <td className="px-4 py-2 text-right font-mono text-xs" style={{ color: "#6b7280" }}>${(row.val * 52).toLocaleString("en-AU")}/yr</td>
                          </tr>
                        ))}
                        <tr style={{ borderTop: `2px solid ${C.teal}`, backgroundColor: `${C.teal}15` }}>
                          <td className="px-4 py-2.5 font-bold" style={{ color: C.navy }}>Total Utilities</td>
                          <td className="px-4 py-2.5 text-right font-mono font-bold" style={{ color: C.navy }}>${utilEstimate.total}/wk</td>
                          <td className="px-4 py-2.5 text-right font-mono font-bold text-xs" style={{ color: "#6b7280" }}>${(utilEstimate.total * 52).toLocaleString("en-AU")}/yr</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div className="flex items-center gap-4 p-4 rounded-lg mb-4" style={{ backgroundColor: `${C.teal}10`, border: `1px solid ${C.teal}` }}>
                    <div className="flex-1">
                      <p className="text-xs font-semibold" style={{ color: C.navy }}>Apply this estimate to your budget?</p>
                      <p className="text-xs" style={{ color: "#6b7280" }}>
                        This will set your Utilities field on the Budget tab to <strong>${utilEstimate.total}/week</strong>.
                        {livingCosts.utilities > 0 && (
                          <span> It will replace the current value of <strong>${livingCosts.utilities}/week</strong>.</span>
                        )}
                      </p>
                    </div>
                    <button onClick={applyUtilityEstimate}
                      className="px-5 py-2.5 rounded-lg text-sm font-bold text-white transition whitespace-nowrap"
                      style={{ backgroundColor: C.navy }}>
                      Use ${utilEstimate.total}/wk →
                    </button>
                  </div>
                </>
              )}
            </Section>

            <Section title="Understanding Each Utility">
              <GreyNote>Click on each utility below to learn more about what drives costs and how to save money. These estimates are indicative and based on typical usage patterns for students in 2024–25.</GreyNote>
              <UtilityCard info={UTIL_INFO.electricity} estimate={utilEstimate.electricity} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.gas} estimate={utilEstimate.gas} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.water} estimate={utilEstimate.water} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.internet} estimate={utilEstimate.internet} isBundled={utilEstimate.bundled} />
            </Section>

            <Section title="Quick Reference — Estimated Weekly Utility Costs">
              <GreyNote>This table summarises the estimated weekly per-person utility cost by accommodation type and climate zone. All values are in $/week.</GreyNote>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.navy }}>
                      <th className="text-left px-3 py-2 font-semibold text-white">Accommodation Type</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Warm</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Mild</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Cool</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ACCOM_TYPES.map(a => (
                      <tr key={a.key} style={{
                        borderBottom: "1px solid #f3f4f6",
                        backgroundColor: a.key === utilAccomType ? `${C.teal}15` : "transparent",
                      }}>
                        <td className="px-3 py-2" style={{ color: C.navy, fontWeight: a.key === utilAccomType ? 600 : 400 }}>{a.label}</td>
                        {CLIMATE_ZONES.map(cz => {
                          const est = getUtilityEstimate(a.key, cz.key);
                          const isActive = a.key === utilAccomType && cz.key === utilClimate;
                          return (
                            <td key={cz.key} className="px-3 py-2 text-center font-mono" style={{
                              color: est.bundled ? "#9ca3af" : C.navy,
                              fontWeight: isActive ? 700 : 400,
                              backgroundColor: isActive ? `${C.teal}30` : "transparent",
                            }}>
                              {est.bundled ? "Included" : `$${est.total}`}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>

            <div className="p-4 rounded text-xs" style={{ backgroundColor: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d" }}>
              <strong>Note:</strong> These estimates are indicative rules of thumb based on average retail energy pricing and typical student usage patterns. Your actual costs will depend on your specific energy retailer, plan, appliances, usage habits, and the energy efficiency of your dwelling. Prices are based on 2024–25 data and are subject to change.
            </div>
          </div>
        )}

        {/* ═══ TAB 3: FUNDING ═══ */}
        {tab === "funding" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Funding Summary" />
            <GreyNote>
              When you move into your own accommodation you will likely have a number of upfront costs to pay. This section estimates those costs and compares them against the savings you may have accumulated before you move out.
            </GreyNote>

            <Section title="Estimated Move-Out Costs & Savings">
              <div className="overflow-x-auto mb-4">
                <table className="w-full text-sm">
                  <thead>
                    <tr style={{ backgroundColor: C.navy }}>
                      <th className="text-left px-4 py-2.5 font-semibold text-white">Item</th>
                      <th className="text-left px-4 py-2.5 font-semibold text-white w-48">Basis</th>
                      <th className="text-right px-4 py-2.5 font-semibold text-white w-32">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr><td colSpan={3} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.coral}10`, color: C.navy }}>Upfront Costs</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Rental bond *</td>
                      <td className="px-4 py-2 text-xs" style={{ color: "#6b7280" }}>4 weeks × ${weeklyAccom}/wk</td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(rentalBond)}</td>
                    </tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Rent in advance</td>
                      <td className="px-4 py-2 text-xs" style={{ color: "#6b7280" }}>2 weeks × ${weeklyAccom}/wk</td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(rentInAdvance)}</td>
                    </tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Furniture & household items</td>
                      <td className="px-4 py-2">
                        <div className="relative" style={{ maxWidth: 120 }}>
                          <span className="absolute left-2 top-1/2 -translate-y-1/2 text-sm text-gray-400">$</span>
                          <input type="number" value={furnitureCost === 0 ? "" : furnitureCost}
                            onChange={e => setFurnitureCost(e.target.value === "" ? 0 : parseFloat(e.target.value) || 0)}
                            placeholder="0" min={0} step={50}
                            className="border rounded py-1 pl-6 pr-2 text-sm bg-white focus:outline-none focus:ring-2 w-full"
                            style={{ borderColor: "#d1d5db" }} />
                        </div>
                      </td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(furnitureCost)}</td>
                    </tr>
                    <tr style={{ borderTop: "2px solid #e5e7eb" }}>
                      <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total Upfront Costs</td>
                      <td className="px-4 py-2"></td>
                      <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(totalUpfrontCosts)}</td>
                    </tr>
                    <tr><td colSpan={3} className="py-1"></td></tr>
                    <tr><td colSpan={3} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.teal}15`, color: C.navy }}>Estimated Savings</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: "#065f46", paddingLeft: 28 }}>Less: Estimated savings at move-out</td>
                      <td className="px-4 py-2">
                        <div className="relative" style={{ maxWidth: 120 }}>
                          <span className="absolute left-2 top-1/2 -translate-y-1/2 text-sm text-gray-400">$</span>
                          <input type="number" value={estimatedSavingsInput === 0 ? "" : estimatedSavingsInput}
                            onChange={e => setEstimatedSavingsInput(e.target.value === "" ? 0 : parseFloat(e.target.value) || 0)}
                            placeholder="0" min={0} step={100}
                            className="border rounded py-1 pl-6 pr-2 text-sm bg-white focus:outline-none focus:ring-2 w-full"
                            style={{ borderColor: "#d1d5db" }} />
                        </div>
                      </td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: estimatedSavingsInput > 0 ? "#065f46" : "#9ca3af" }}>
                        {estimatedSavingsInput > 0 ? `($${estimatedSavingsInput.toLocaleString("en-AU")})` : fmt(0)}
                      </td>
                    </tr>
                    <tr><td colSpan={3} className="py-1"></td></tr>
                    <tr style={{ borderTop: `3px solid ${fundingGap > 0 ? C.coral : C.teal}`, backgroundColor: fundingGap > 0 ? `${C.coral}15` : `${C.teal}15` }}>
                      <td className="px-4 py-3 font-bold" style={{ color: fundingGap > 0 ? C.coral : "#065f46" }}>
                        {fundingGap > 0 ? "Estimated Funding Gap" : "Estimated Surplus After Upfront Costs"}
                      </td>
                      <td className="px-4 py-3"></td>
                      <td className="px-4 py-3 text-right font-mono font-bold text-base" style={{ color: fundingGap > 0 ? C.coral : "#065f46" }}>
                        {fmt(Math.abs(Math.round(fundingGap)))}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div className="p-4 rounded-lg mb-4" style={{ backgroundColor: fundingGap > 0 ? "#fef2f2" : "#f0fdf4", border: `2px solid ${fundingGap > 0 ? C.coral : C.teal}` }}>
                <p className="text-sm font-bold" style={{ color: fundingGap > 0 ? C.coral : "#065f46" }}>
                  {fundingGap > 0
                    ? `Based on your assumptions, you will need approximately $${Math.round(fundingGap).toLocaleString("en-AU")} in additional funding to cover your upfront accommodation costs. The loan options below may help bridge this gap.`
                    : `Based on your assumptions, your estimated savings should cover your upfront accommodation costs${fundingGap < 0 ? ` with approximately $${Math.abs(Math.round(fundingGap)).toLocaleString("en-AU")} to spare` : ""}.`
                  }
                </p>
              </div>
              <p className="text-xs" style={{ color: "#9ca3af" }}>
                * Rental bonds are refundable security deposits, typically equal to 4 weeks' rent, designed to cover potential damages or non-payment of rent. The bond is usually paid by the tenant at the start of a lease and held by a government authority on behalf of the landlord.
              </p>
            </Section>

            <SectionDivider title="Funding Options" />
            <GreyNote>
              If your budget indicates that you may not have enough money to cover your expected costs you may need to take out a loan. Below are two loan options that may help you fund your expenses.
            </GreyNote>
            <Section title="1. Student Start-up Loan (SSL)">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">The Student Start-up Loan (SSL) is a voluntary, tax-free loan from the Australian Government designed to help eligible higher education students with the costs of study, including textbooks, equipment, travel and living expenses.</p>
                <p className="mb-2">You can borrow <strong style={{ color: C.navy }}>$1,349 per semester</strong> (up to twice per calendar year, i.e. up to $2,698 per year) for the duration of your course, provided you continue to meet the eligibility requirements and apply each period.</p>
                <p className="mb-2">To be eligible, you must receive at least $1 of Youth Allowance (as a student), Austudy, or ABSTUDY Living Allowance in the relevant fortnight.</p>
                <p className="mb-2">The SSL is added to your HELP debt and is repaid through the tax system once your income exceeds the compulsory repayment threshold (same as HECS-HELP debts). The loan is interest-free but is subject to annual indexation, which means the total amount you repay will be more than you borrow. Indexation is applied on 1 June each year once the debt is at least 11 months old.</p>
                <p>More info:{" "}<a href="https://www.servicesaustralia.gov.au/student-start-up-loan" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>Services Australia — Student Start-up Loan</a></p>
              </div>
              <div className="p-3 rounded text-xs" style={{ backgroundColor: "#f0f4ff", border: `1px solid ${C.navy}20` }}>
                <div className="grid grid-cols-2 gap-2" style={{ maxWidth: 400 }}>
                  <div style={{ color: C.navy }}>Amount per semester</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$1,349</div>
                  <div style={{ color: C.navy }}>Maximum per year</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$2,698</div>
                  <div style={{ color: C.navy }}>Interest rate</div><div className="font-mono font-semibold" style={{ color: C.navy }}>0% (indexed to CPI)</div>
                  <div style={{ color: C.navy }}>Repayment</div><div className="font-semibold" style={{ color: C.navy }}>Via tax system (HELP debt)</div>
                </div>
              </div>
            </Section>
            <Section title="2. Refugee Student Loan Program">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">The Refugee Student Loan Program has been developed by Skill Path for students in the RSSP. It is administered by Spark Finance. These loans are for up to <strong style={{ color: C.navy }}>$5,000/year</strong> during study. They have a <strong style={{ color: C.navy }}>7% interest rate</strong> and a <strong style={{ color: C.navy }}>7 year term</strong> with no repayment required until after you complete your course.</p>
                <p>More information about this loan is on the RSSP Student Hub in Notion{" "}<a href="https://www.notion.so/Student-Loan-Scheme-with-Spark-Finance-2c4199deb36c806b9819f6ad2565706c?source=copy_link" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.</p>
              </div>
              <div className="p-3 rounded text-xs" style={{ backgroundColor: "#f0f4ff", border: `1px solid ${C.navy}20` }}>
                <div className="grid grid-cols-2 gap-2" style={{ maxWidth: 400 }}>
                  <div style={{ color: C.navy }}>Maximum per year</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$5,000</div>
                  <div style={{ color: C.navy }}>Interest rate</div><div className="font-mono font-semibold" style={{ color: C.navy }}>7%</div>
                  <div style={{ color: C.navy }}>Loan term</div><div className="font-mono font-semibold" style={{ color: C.navy }}>7 years</div>
                  <div style={{ color: C.navy }}>Repayment during study</div><div className="font-semibold" style={{ color: C.navy }}>None required</div>
                  <div style={{ color: C.navy }}>Administered by</div><div className="font-semibold" style={{ color: C.navy }}>Spark Finance</div>
                </div>
              </div>
            </Section>
          </div>
        )}
      </div>
    </div>
  );
}