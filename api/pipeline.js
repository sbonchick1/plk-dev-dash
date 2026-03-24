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
  const PIPELINE_SHEET_ID = "2173912308726660"; // 02. Opening sheet
  const BUDGET_SHEET_ID = "3699973825901444"; // Budget sheet

  try {
    // Fetch pipeline data
    const pipelineResponse = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${PIPELINE_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const pipelineText = await pipelineResponse.text();
    let pipelineData;
    
    try { 
      pipelineData = JSON.parse(pipelineText); 
    } catch(e) { 
      return res.status(500).json({ 
        error: "Failed to parse pipeline Smartsheet response", 
        raw: pipelineText.slice(0, 300) 
      }); 
    }
    
    if (!pipelineData.columns || !pipelineData.rows) {
      return res.status(500).json({ error: "Unexpected pipeline response", detail: pipelineData });
    }

    // Fetch budget data
    const budgetResponse = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${BUDGET_SHEET_ID}`,
      { headers: { "Authorization": "Bearer " + TOKEN } }
    );
    const budgetText = await budgetResponse.text();
    let budgetData;
    
    try { 
      budgetData = JSON.parse(budgetText); 
    } catch(e) { 
      return res.status(500).json({ 
        error: "Failed to parse budget Smartsheet response", 
        raw: budgetText.slice(0, 300) 
      }); 
    }

    // Build column map for pipeline
    const colMap = {};
    pipelineData.columns.forEach(function(col, i) { colMap[col.title] = i; });

    function get(row, title) {
      const idx = colMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.displayValue !== undefined) return cell.displayValue;
      if (cell.value !== undefined) return cell.value;
      return null;
    }

    // Parse pipeline rows
    const rows = pipelineData.rows.map(function(row) {
      return {
        division: get(row, "Division"),
        status: get(row, "Status"),
        riskLevel: get(row, "Risk Level"),
      };
    }).filter(function(r) {
      // Only include rows with valid division and status
      return r.division && r.status;
    });

    // Build budget map from budget sheet
    const budgetColMap = {};
    budgetData.columns.forEach(function(col, i) { budgetColMap[col.title] = i; });

    function getBudget(row, title) {
      const idx = budgetColMap[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      if (cell.value !== undefined) return cell.value;
      return null;
    }

    const budgets = {};
    budgetData.rows.forEach(function(row) {
      const div = getBudget(row, "Division");
      const budget = getBudget(row, "Openings Budget");
      if (div && budget && typeof budget === 'number') {
        budgets[div] = budget;
      }
    });

    res.status(200).json({ 
      rows: rows, 
      budgets: budgets,
      lastUpdated: new Date().toISOString() 
    });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
