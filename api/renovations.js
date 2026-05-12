const fetch = require("node-fetch");

const SHEET_ID = "3928541238349700";

// Map raw Smartsheet status values → consolidated stage buckets
const STAGE_MAP = {
  "Not Started":                        "Not Started",
  "360° Survey Quoted":            "360 Phase",
  "360° Survey Complete":          "360 Phase",
  "360° Survey Scheduled":         "360 Phase",
  "RSCF / Scope Finalization":          "360 Phase",
  "Pending Design Submittal":           "In Design",
  "Permitting":                         "Permitting",
  "Under Construction":                 "Under Construction",
  "Punch List - FZ Action":             "Punch List",
  "Punch List - CM Review":             "Punch List",
  "Punch List - Pending Photos":        "Punch List",
  "Complete":                           "Complete",
  "On hold / pending FZ reply":         "360 Phase",
  "CGBR Complete / Waiting for FZ":     "Permitting",
};

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  const TOKEN = process.env.SMARTSHEET_TOKEN;
  if (!TOKEN) return res.status(500).json({ error: "SMARTSHEET_TOKEN not set" });

  try {
    const response = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const text = await response.text();
    let data;
    try { data = JSON.parse(text); }
    catch(e) { return res.status(500).json({ error: "Failed to parse Smartsheet response", raw: text.slice(0, 300) }); }

    if (!data.columns || !data.rows)
      return res.status(500).json({ error: "Unexpected Smartsheet response", detail: data });

    // Debug: return all column names so we can verify exact titles
    if (req.query && req.query.debug === "1") {
      return res.status(200).json({
        columns: data.columns.map(function(c, i) { return { index: i, id: c.id, title: c.title }; }),
        rowCount: data.rows.length
      });
    }

    const colMap = {};
    data.columns.forEach(function(col, i) { colMap[col.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.displayValue !== undefined) return cell.displayValue;
      if (cell.value       !== undefined) return cell.value;
      return null;
    }

    const rows = data.rows.map(function(row) {
      const rawStatus = get(row, "Renovation Status") || null;
      const stage     = rawStatus ? (STAGE_MAP[rawStatus.trim()] || null) : null;
      return {
        plk:      get(row, "Store Number"),
        division: get(row, "Division") || null,
        bu:       get(row, "BU"),
        rawStatus,
        stage,
      };
    }).filter(function(r) {
      if (!r.division || !r.stage) return false;
      const buNum = parseFloat(r.bu);
      return !isNaN(buNum) && buNum >= 1 && buNum <= 12;
    });

    res.status(200).json({ rows, lastUpdated: new Date().toISOString() });

  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
