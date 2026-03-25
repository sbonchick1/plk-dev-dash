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
  const PIPELINE_SHEET_ID = "8717733601798020";
  const BUDGET_SHEET_ID = "7864670409936772";

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
      return res.status(500).json({ error: "Failed to parse pipeline response", raw: pipelineText.slice(0, 300) }); 
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
      return res.status(500).json({ error: "Failed to parse budget response", raw: budgetText.slice(0, 300) }); 
    }

    // Return the RAW budget sheet so we can see exactly what Smartsheet gives us
    return res.status(200).json({
      budgetSheetName: budgetData.name,
      budgetColumns: budgetData.columns,
      budgetRows: budgetData.rows.slice(0, 15)
    });

  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
