/**
 * Splice a freshly-composed DASHBOARD_DATA blob into the dashboard HTML file
 * by replacing the content between the <!-- DASHBOARD_DATA:START --> and
 * <!-- DASHBOARD_DATA:END --> markers.
 *
 * The HTML must already have the markers in place — see the refactored
 * index.html which uses a <script type="application/json" id="dashboard-data">
 * block parsed by the in-page render code on load.
 */

const fs = require('fs');

const START = '<!-- DASHBOARD_DATA:START -->';
const END = '<!-- DASHBOARD_DATA:END -->';

// Sentinel + CSS block re-injected on every refresh if missing. Keeps the
// dashboard usable on phones even if a future external edit strips the
// existing rules. Bump the version suffix in the marker if you change the
// CSS so older injections get replaced — current logic only injects when
// the marker is absent.
const MOBILE_CSS_MARKER = '/* corecast-mobile-css:v1 */';
const MOBILE_CSS = `
${MOBILE_CSS_MARKER}
@media (max-width: 640px) {
  .stat-grid { grid-template-columns: repeat(2, 1fr); }
  .job-metrics { grid-template-columns: repeat(2, 1fr); }
  .alerts-row { grid-template-columns: 1fr; }
  .tab-bar { overflow-x: auto; flex-wrap: nowrap; -webkit-overflow-scrolling: touch; }
  .tab { flex-shrink: 0; white-space: nowrap; }
  table { font-size: 12px; }
  .quote-row { flex-direction: column; align-items: flex-start; gap: 6px; }
  .completed-row { flex-direction: column; align-items: flex-start; gap: 6px; }
  .job-card-header { flex-direction: column; gap: 8px; }
  .header-sub-text { display: none; }
  .tab-icon { display: none; }
}
`;

function ensureMobileCss(html) {
  if (html.includes(MOBILE_CSS_MARKER)) return html;
  // Inject just before the LAST </style> close tag so it sits at the end of
  // the stylesheet and overrides any earlier rules with the same specificity.
  const idx = html.lastIndexOf('</style>');
  if (idx === -1) return html; // no style block — bail rather than corrupt
  return html.slice(0, idx) + MOBILE_CSS + '\n' + html.slice(idx);
}

function renderToHtml(htmlPath, data) {
  if (!fs.existsSync(htmlPath)) {
    throw new Error(`Dashboard HTML not found: ${htmlPath}`);
  }
  const html = ensureMobileCss(fs.readFileSync(htmlPath, 'utf8'));
  const startIdx = html.indexOf(START);
  const endIdx = html.indexOf(END);
  if (startIdx === -1 || endIdx === -1) {
    throw new Error(
      `Missing data markers in ${htmlPath}. Expected ${START} … ${END}. ` +
      `Has the HTML been refactored to the data-driven layout?`,
    );
  }
  if (endIdx < startIdx) {
    throw new Error(`END marker before START in ${htmlPath}`);
  }

  // Escape `<` to keep `</script>` (or anything HTML-ish) inside the JSON
  // payload safe to embed in a script tag.
  const safeJson = JSON.stringify(data, null, 2).replace(/</g, '\\u003c');

  const replacement =
    START + '\n' +
    `<script id="dashboard-data" type="application/json">\n` +
    safeJson + '\n' +
    `</script>\n` +
    END;

  const newHtml = html.slice(0, startIdx) + replacement + html.slice(endIdx + END.length);

  // Sanity: the post-splice HTML must still end with </html>. If it doesn't,
  // something went wrong (truncation, marker collision, etc.) and we'd rather
  // bail than write a broken dashboard the user has to debug visually.
  if (!newHtml.trimEnd().endsWith('</html>')) {
    throw new Error(
      'render.js refusing to write: post-splice HTML does not end with </html>. ' +
      'Last 120 chars: ' + JSON.stringify(newHtml.slice(-120))
    );
  }

  // Atomic-ish write
  const tmp = htmlPath + '.tmp';
  fs.writeFileSync(tmp, newHtml);
  fs.renameSync(tmp, htmlPath);
}

module.exports = { renderToHtml, ensureMobileCss, MOBILE_CSS_MARKER, START, END };
