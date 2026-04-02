/**
 * kpi-fix.js
 * Drop this script tag into any Popeyes dashboard page to lock
 * the KPI banner to always one row regardless of zoom level.
 * 
 * Usage: Add to each HTML file, right before </body>:
 *   <script src="/kpi-fix.js"></script>
 */
(function() {
  const style = document.createElement('style');
  style.textContent = `
    /* ── KPI BANNER: always one row, never wraps ── */
    .kpi-banner {
      display: flex !important;
      flex-wrap: nowrap !important;
      grid-template-columns: unset !important;
      gap: 12px !important;
    }
    .kpi-banner > .kpi-card {
      flex: 1 1 0 !important;
      min-width: 0 !important;
      padding: 14px 10px !important;
    }
    .kpi-banner .kpi-value {
      font-size: clamp(14px, 1.8vw, 26px) !important;
      line-height: 1 !important;
    }
    .kpi-banner .kpi-label {
      font-size: clamp(7px, 0.65vw, 10px) !important;
      white-space: nowrap !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
    }
    .kpi-banner .kpi-sub {
      font-size: clamp(7px, 0.62vw, 11px) !important;
      white-space: nowrap !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
    }
    /* On very small screens allow wrap as fallback */
    @media (max-width: 600px) {
      .kpi-banner { flex-wrap: wrap !important; }
    }
  `;
  document.head.appendChild(style);
})();
