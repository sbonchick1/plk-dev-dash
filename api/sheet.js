const fetch = require("node-fetch");

module.exports = async function(req, res) {
  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  const TOKEN = process.env.SMARTSHEET_TOKEN;
  const SHEET_ID = "6457304820961156";

  try {
    const response = await fetch(
      "https://api.smartsheet.com/2.0/sheets/" + SHEET_ID,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const text = await response.text();
    let data;
    
    try { 
      data = JSON.parse(text); 
    } catch(e) { 
      return res.status(500).json({ 
        error: "Failed to parse Smartsheet response", 
        raw: text.slice(0, 300) 
      }); 
    }
    
    if (!data.columns || !data.rows) {
      return res.status(500).json({ error: "Unexpected response", detail: data });
    }

    const colMap = {};
    data.columns.forEach(function(col, i) { colMap[col.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.displayValue !== undefined) return cell.displayValue;
      if (cell.value !== undefined) return cell.value;
      return null;
    }

    function parseNum(v) {
      if (!v) return null;
      const s = String(v);
      if (s === "NA" || s === "-" || s === "Missing" || s === "N/A") return null;
      if (s.indexOf("INVALID") !== -1 || s.indexOf("NO MATCH") !== -1) return null;
      const n = parseFloat(s.replace(/[$,%]/g, "").replace(/,/g, "").trim());
      return isNaN(n) ? null : n;
    }

    const rows = data.rows.map(function(row) {
      return {
        restNum:     get(row, "Rest #"),
        openYear:    parseNum(get(row, "Restaurant Opening Year")),
        state:       get(row, "State"),
        div:         get(row, "Division"),
        urb:         get(row, "Urbanicity"),
        deal:        get(row, "Deal Type"),
        arch:        (function() {
          var a = get(row, "Architecture Type");
          if (!a) return null;
          if (a.indexOf("Inline") !== -1) return "Inline";
          return a;
        })(),
        annARS:      parseNum(get(row, "Annualized ARS")),
        annEB:       parseNum(get(row, "Annualized EBITDA $")),
        ebPct:       parseNum(get(row, "Annualized EBITDA %")),
        fssScore:    parseNum(get(row, "Average of Past 2 Rounds Score")),
        fssGrade:    get(row, "Average of Past 2 Rounds Grade"),
        arsBench:    parseNum(get(row, "ARS Benchmark for Analysis")),
        ebitdaBench: parseNum(get(row, "EBITDA Benchmark for Analysis")),
      };
    }).filter(function(r) {
      return r.openYear && r.openYear >= 2020;
    });

    res.status(200).json({ rows: rows, lastUpdated: new Date().toISOString() });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
