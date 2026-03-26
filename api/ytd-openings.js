const fetch = require("node-fetch");

module.exports = async function(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  const TOKEN         = process.env.SMARTSHEET_TOKEN;
  const PIPELINE_ID   = "8717733601798020";
  const MASTER_FZ_ID  = "5821658321342340";
  const MASTER_REC_ID = "3277302519517060";

  try {
    const [pr, fr, rr] = await Promise.all([
      fetch(`https://api.smartsheet.com/2.0/sheets/${PIPELINE_ID}`,   { headers: { Authorization: "Bearer " + TOKEN } }),
      fetch(`https://api.smartsheet.com/2.0/sheets/${MASTER_FZ_ID}`,  { headers: { Authorization: "Bearer " + TOKEN } }),
      fetch(`https://api.smartsheet.com/2.0/sheets/${MASTER_REC_ID}`, { headers: { Authorization: "Bearer " + TOKEN } }),
    ]);
    const [pipeline, fzSheet, recSheet] = await Promise.all([
      pr.text().then(t => JSON.parse(t)),
      fr.text().then(t => JSON.parse(t)),
      rr.text().then(t => JSON.parse(t)),
    ]);

    function colMap(sheet) {
      const m = {};
      (sheet.columns || []).forEach((c, i) => { m[c.title] = i; });
      return m;
    }
    function get(row, map, title) {
      const idx = map[title];
      if (idx === undefined) return null;
      const cell = row.cells[idx];
      if (!cell) return null;
      return cell.displayValue !== undefined && cell.displayValue !== null
        ? cell.displayValue : cell.value !== undefined ? cell.value : null;
    }
    function num(v) {
      if (v === null || v === undefined) return null;
      const s = String(v).replace(/[$,%]/g, "").replace(/,/g, "").trim();
      if (!s || s === "-" || s === "NA" || s === "N/A") return null;
      const n = parseFloat(s);
      return isNaN(n) ? null : n;
    }

    const pMap = colMap(pipeline);
    const opened = (pipeline.rows || [])
      .map(row => ({
        restNum:  get(row, pMap, "Rest No"),
        fz:       get(row, pMap, "FZ"),
        division: get(row, pMap, "Division"),
        address:  get(row, pMap, "Address"),
        city:     get(row, pMap, "City"),
        state:    get(row, pMap, "ST"),
        status:   get(row, pMap, "Status"),
        fzPOD:    get(row, pMap, "FZ Projected Open Date"),
      }))
      .filter(r => r.status === "Open" && r.restNum);

    const fzColMap = colMap(fzSheet);
    const fzLookup = {};
    (fzSheet.rows || []).forEach(row => {
      const k = String(get(row, fzColMap, "PLK#") || "").trim();
      if (!k) return;
      fzLookup[k] = {
        fzName:       get(row, fzColMap, "FZ Name"),
        ttmARS:       num(get(row, fzColMap, "Trailing 12M ARS")),
        annARS:       num(get(row, fzColMap, "Annualized ARS")),
        ttmEBITDA:    num(get(row, fzColMap, "Trailing 12M EBITDA")),
        annEBITDA:    num(get(row, fzColMap, "Annualized EBITDA")),
        annEBITDAPct: num(get(row, fzColMap, "Annualized EBITDA %")),
        fssGrade:     get(row, fzColMap, "2026 Round 1 FSS Grade"),
        driveThru:    get(row, fzColMap, "Drive-Thru Type"),
        archType:     get(row, fzColMap, "Architecture Type"),
      };
    });

    const recColMap = colMap(recSheet);
    const recLookup = {};
    (recSheet.rows || []).forEach(row => {
      const k = String(get(row, recColMap, "Store#") || "").trim();
      if (!k) return;
      recLookup[k] = {
        salesForecast: num(get(row, recColMap, "Sales Forecast")),
        capEx:         num(get(row, recColMap, "CapEx Estimate")),
        estEBITDA:     num(get(row, recColMap, "Estimated EBITDA $")),
        estEBITDAPct:  num(get(row, recColMap, "Estimated EBITDA %")),
        roi:           num(get(row, recColMap, "ROI")),
      };
    });

    const today = Date.now();
    const rows = opened.map(r => {
      const key = String(r.restNum || "").trim();
      const fz  = fzLookup[key]  || {};
      const rec = recLookup[key] || {};

      let daysOpen = null, openDate = null;
      if (r.fzPOD) {
        const d = new Date(r.fzPOD);
        if (!isNaN(d.getTime())) {
          openDate = r.fzPOD;
          daysOpen = Math.max(1, Math.floor((today - d.getTime()) / 86400000));
        }
      }

      const ttmARS = fz.ttmARS;
      let avgWeekly = null, avgMonthly = null, avgAnnual = null;
      if (ttmARS != null && daysOpen != null && daysOpen > 0) {
        const dailyRate = ttmARS / daysOpen;
        avgWeekly  = dailyRate * 7;
        avgMonthly = dailyRate * 30.44;
        avgAnnual  = dailyRate * 365;
      }

      const salesForecast  = rec.salesForecast;
      const forecastVarD   = (avgAnnual != null && salesForecast) ? avgAnnual - salesForecast : null;
      const forecastVarPct = (forecastVarD != null && salesForecast) ? forecastVarD / salesForecast : null;

      return {
        restNum: r.restNum, fzName: fz.fzName || r.fz || null,
        division: r.division, address: r.address, city: r.city, state: r.state,
        openDate, daysOpen,
        ttmARS, annARS: fz.annARS, ttmEBITDA: fz.ttmEBITDA,
        annEBITDA: fz.annEBITDA, annEBITDAPct: fz.annEBITDAPct,
        fssGrade: fz.fssGrade, driveThru: fz.driveThru, archType: fz.archType,
        avgWeekly, avgMonthly, avgAnnual,
        salesForecast, forecastVarD, forecastVarPct,
        capEx: rec.capEx, estEBITDA: rec.estEBITDA,
        estEBITDAPct: rec.estEBITDAPct, roi: rec.roi,
      };
    });

    rows.sort((a, b) => {
      if (!a.openDate && !b.openDate) return 0;
      if (!a.openDate) return 1; if (!b.openDate) return -1;
      return new Date(a.openDate) - new Date(b.openDate);
    });

    res.status(200).json({ rows, lastUpdated: new Date().toISOString() });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
};
