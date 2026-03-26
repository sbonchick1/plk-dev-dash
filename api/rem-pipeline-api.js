const fetch = require("node-fetch");

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  const TOKEN          = process.env.SMARTSHEET_TOKEN;
  const PIPELINE_ID    = "8717733601798020";
  const BUDGET_SHEET_ID= "7864670409936772";

  try {
    const [pipelineResp, budgetResp] = await Promise.all([
      fetch(`https://api.smartsheet.com/2.0/sheets/${PIPELINE_ID}`,    { headers: { Authorization: "Bearer " + TOKEN } }),
      fetch(`https://api.smartsheet.com/2.0/sheets/${BUDGET_SHEET_ID}`, { headers: { Authorization: "Bearer " + TOKEN } }),
    ]);

    const [pipeline, budgetData] = await Promise.all([
      pipelineResp.text().then(t => JSON.parse(t)),
      budgetResp.text().then(t => JSON.parse(t)),
    ]);

    // ── Pipeline column map ──────────────────────────────────────────────────
    const colMap = {};
    (pipeline.columns || []).forEach((c, i) => { colMap[c.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      return cell.displayValue !== undefined && cell.displayValue !== null
        ? cell.displayValue : cell.value !== undefined ? cell.value : null;
    }

    function parseDate(v) {
      if (!v) return null;
      const d = new Date(v);
      return isNaN(d.getTime()) ? null : d.getTime();
    }

    // ── Filter to Prospect + SA rows only ────────────────────────────────────
    const EARLY_STATUSES = ["Prospect", "SA"];

    const rows = (pipeline.rows || [])
      .map(row => ({
        division:    get(row, "Division"),
        status:      get(row, "Status"),
        riskLevel:   get(row, "Risk Level"),
        sipId:       get(row, "SIP ID"),
        restNum:     get(row, "Rest No"),
        fz:          get(row, "FZ"),
        address:     get(row, "Address"),
        city:        get(row, "City"),
        state:       get(row, "ST"),
        fzOpenDate:  get(row, "FZ Projected Open Date"),
        plkOpenDate: get(row, "PLK Projected Open Date"),
        lastComment: get(row, "Last Comments"),
        saStart:     get(row, "SA Start"),
        saStartMs:   parseDate(get(row, "SA Start")),
      }))
      .filter(r => r.division && r.status && EARLY_STATUSES.includes(r.status));

    // ── SA Budget: rows 26-31 (0-indexed: 25-30), column A = division, column B = SA budget ──
    const colA = budgetData.columns.find(c => c.title === "A");
    const colB = budgetData.columns.find(c => c.title === "B");

    if (!colA || !colB) {
      return res.status(500).json({
        error: "Could not find budget columns A or B",
        foundColumns: (budgetData.columns || []).map(c => c.title)
      });
    }

    function getCellById(row, columnId) {
      const cell = (row.cells || []).find(c => c.columnId === columnId);
      if (!cell) return null;
      return cell.value !== undefined ? cell.value : null;
    }

    // rows 26–31 = index 25–30 (SA budget section)
    const saTargets = {};
    (budgetData.rows || []).forEach((row, index) => {
      if (index < 25 || index > 30) return;
      const div       = getCellById(row, colA.id);
      const targetRaw = getCellById(row, colB.id);
      const target    = typeof targetRaw === "number"
        ? targetRaw
        : parseFloat(String(targetRaw || "").replace(/[,$]/g, ""));
      if (div && !isNaN(target) && target > 0) saTargets[div] = target;
    });

    res.status(200).json({ rows, saTargets, lastUpdated: new Date().toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
