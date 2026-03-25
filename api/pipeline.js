const fetch = require("node-fetch");

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  const TOKEN            = process.env.SMARTSHEET_TOKEN;
  const PIPELINE_SHEET_ID = "8717733601798020";
  const BUDGET_SHEET_ID   = "7864670409936772";

  try {
    // ── Pipeline sheet ──────────────────────────────────────────────────────
    const pipelineResponse = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${PIPELINE_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const pipelineText = await pipelineResponse.text();
    let pipelineData;
    try { pipelineData = JSON.parse(pipelineText); }
    catch(e) { return res.status(500).json({ error: "Failed to parse pipeline response", raw: pipelineText.slice(0,300) }); }
    if (!pipelineData.columns || !pipelineData.rows) {
      return res.status(500).json({ error: "Unexpected pipeline response", detail: pipelineData });
    }

    // ── Budget sheet ────────────────────────────────────────────────────────
    const budgetResponse = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${BUDGET_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const budgetText = await budgetResponse.text();
    let budgetData;
    try { budgetData = JSON.parse(budgetText); }
    catch(e) { return res.status(500).json({ error: "Failed to parse budget response", raw: budgetText.slice(0,300) }); }

    // ── Column map ──────────────────────────────────────────────────────────
    const colMap = {};
    pipelineData.columns.forEach(function(col, i) { colMap[col.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.displayValue !== undefined) return cell.displayValue;
      if (cell.value       !== undefined) return cell.value;
      return null;
    }

    // Parse a date string to ms timestamp; returns null if unparseable
    function parseDate(v) {
      if (!v) return null;
      const d = new Date(v);
      return isNaN(d.getTime()) ? null : d.getTime();
    }

    // ── Parse rows — now includes stage-date fields ─────────────────────────
    const rows = pipelineData.rows.map(function(row) {
      return {
        // existing fields
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
        // NEW: stage date fields (raw string + parsed ms)
        saStart:          get(row, "SA Start"),
        pcStart:          get(row, "PC Start"),
        ucStart:          get(row, "UC Start"),
        permitStarted:    get(row, "Permit Started"),
        permitCompleted:  get(row, "Permit Completed"),
        // parsed as timestamps for easy arithmetic on the frontend
        saStartMs:         parseDate(get(row, "SA Start")),
        pcStartMs:         parseDate(get(row, "PC Start")),
        ucStartMs:         parseDate(get(row, "UC Start")),
        permitStartedMs:   parseDate(get(row, "Permit Started")),
        permitCompletedMs: parseDate(get(row, "Permit Completed")),
        fzOpenDateMs:      parseDate(get(row, "FZ Projected Open Date")),
      };
    }).filter(function(r) {
      return r.division && r.status;
    });

    // ── Budget map ──────────────────────────────────────────────────────────
    const colA = budgetData.columns.find(function(c) { return c.title === "A"; });
    const colB = budgetData.columns.find(function(c) { return c.title === "B"; });
    if (!colA || !colB) {
      return res.status(500).json({
        error: "Could not find budget columns A or B",
        foundColumns: budgetData.columns.map(function(c) { return c.title; })
      });
    }

    function getCellById(row, columnId) {
      const cell = row.cells.find(function(c) { return c.columnId === columnId; });
      if (!cell) return null;
      return cell.value !== undefined ? cell.value : null;
    }

    const budgets = {};
    budgetData.rows.forEach(function(row, index) {
      if (index < 2 || index > 7) return;
      const div       = getCellById(row, colA.id);
      const budgetRaw = getCellById(row, colB.id);
      const budget    = typeof budgetRaw === "number"
        ? budgetRaw
        : parseFloat(String(budgetRaw || "").replace(/[,$]/g, ""));
      if (div && !isNaN(budget) && budget > 0) budgets[div] = budget;
    });

    res.status(200).json({ rows, budgets, lastUpdated: new Date().toISOString() });

  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
