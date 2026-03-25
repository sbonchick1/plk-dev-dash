const fetch = require("node-fetch");

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  const TOKEN            = process.env.SMARTSHEET_TOKEN;
  const CLOSURE_SHEET_ID = "554959277188996";
  const BUDGET_SHEET_ID  = "7864670409936772";

  try {
    // ── Closure sheet ───────────────────────────────────────────────────────
    const closureResp = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${CLOSURE_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const closureText = await closureResp.text();
    let closureData;
    try { closureData = JSON.parse(closureText); }
    catch(e) { return res.status(500).json({ error: "Failed to parse closure response", raw: closureText.slice(0,300) }); }
    if (!closureData.columns || !closureData.rows) {
      return res.status(500).json({ error: "Unexpected closure response", detail: closureData });
    }

    // ── Budget sheet ────────────────────────────────────────────────────────
    const budgetResp = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${BUDGET_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const budgetText = await budgetResp.text();
    let budgetData;
    try { budgetData = JSON.parse(budgetText); }
    catch(e) { return res.status(500).json({ error: "Failed to parse budget response", raw: budgetText.slice(0,300) }); }

    // ── Column map ──────────────────────────────────────────────────────────
    const colMap = {};
    closureData.columns.forEach(function(col, i) { colMap[col.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.displayValue !== undefined) return cell.displayValue;
      if (cell.value       !== undefined) return cell.value;
      return null;
    }

    function parseNum(v) {
      if (v === null || v === undefined) return null;
      const s = String(v).replace(/[$,%]/g,"").replace(/,/g,"").trim();
      if (!s || s === "-" || s === "NA" || s === "N/A") return null;
      const n = parseFloat(s);
      return isNaN(n) ? null : n;
    }

    // ── Parse rows ──────────────────────────────────────────────────────────
    const rows = closureData.rows.map(function(row) {
      return {
        division:       get(row, "Division"),
        restNum:        get(row, "Rest. No"),
        fz:             get(row, "FZ"),
        address:        get(row, "Address"),
        city:           get(row, "City"),
        state:          get(row, "ST"),
        dateOfClosure:  get(row, "Date of Closure"),
        closureRisk:    get(row, "Closure Risk Level"),
        closureBucket:  get(row, "Closure Bucket"),
        closureReason:  get(row, "Closure Reason"),
        plkControl:     get(row, "PLK Control?"),
        ars2025:        parseNum(get(row, "2025 ARS")),
        ttmEbitda:      parseNum(get(row, "TTM EBITDA")),
        comments:       get(row, "Comments"),
      };
    }).filter(function(r) {
      return r.division && r.closureBucket;
    });

    // ── Budget map — column C for closures ──────────────────────────────────
    const colA = budgetData.columns.find(function(c) { return c.title === "A"; });
    const colC = budgetData.columns.find(function(c) { return c.title === "C"; });
    if (!colA || !colC) {
      return res.status(500).json({
        error: "Could not find budget columns A or C",
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
      const budgetRaw = getCellById(row, colC.id);
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
