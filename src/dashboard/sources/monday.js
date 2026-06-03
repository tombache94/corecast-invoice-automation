/**
 * Fetch active jobs from the CoreCast Monday board. Returns a normalized
 * job list with stage, timeline, crew, contract value, notes, and the
 * resolved Xero job key (via manualAllocations.jobAliases or 1:1 fallback).
 */

const https = require('https');

const BOARD_ID = 5027051539;

const COLUMN_IDS = {
  stage:        'color_mm17qxm2',
  timeline:     'timerange_mm17qx0m',
  notes:        'long_text_mm17a895',
  crew:         'dropdown_mm1fx1zt',
  assignedSubs: 'text_mm17k3m8',
  contractValue:'numeric_mm2zp1ek',
  sharePointLink:'link_mm36d0ea',
};

async function fetchJobs(manualAllocations) {
  const apiKey = process.env.MONDAY_API_KEY;
  if (!apiKey) throw new Error('Missing MONDAY_API_KEY in .env');

  const aliases = (manualAllocations && manualAllocations.jobAliases) || {};
  const meta = (manualAllocations && manualAllocations.jobMeta) || {};

  const data = await gql(apiKey, `
    query {
      boards(ids: [${BOARD_ID}]) {
        items_page(limit: 100) {
          items {
            id
            name
            group { title id }
            column_values { id text value }
          }
        }
      }
    }
  `);

  if (data.errors) {
    throw new Error('Monday API error: ' + JSON.stringify(data.errors));
  }

  const items = data.data.boards[0].items_page.items;
  const jobs = items.map((it) => {
    const cv = (id) => it.column_values.find((c) => c.id === id);
    const text = (id) => cv(id)?.text || null;
    const numericText = text(COLUMN_IDS.contractValue);
    const contractValue = numericText ? Number(numericText) : null;
    const stage = text(COLUMN_IDS.stage); // "In Progress" | "Not begun" | "Completed"
    const timeline = text(COLUMN_IDS.timeline); // "2026-03-30 - 2026-04-07"
    const [startDate, endDate] = (timeline || '').split(' - ').map((s) => s.trim() || null);
    const crewText = text(COLUMN_IDS.crew); // "Tom, Josh, Zolboo"
    const notes = text(COLUMN_IDS.notes);
    const assignedSubs = text(COLUMN_IDS.assignedSubs);
    const linkVal = cv(COLUMN_IDS.sharePointLink)?.value;
    let sharePointUrl = null;
    if (linkVal) {
      try {
        const parsed = JSON.parse(linkVal);
        sharePointUrl = parsed.url || null;
      } catch (_) {}
    }

    const jobKey = aliases[it.name] || it.name;
    const m = meta[jobKey] || {};

    return {
      id: it.id,
      name: it.name,
      jobKey,                  // canonical key used to join with Xero tracking
      group: it.group.title,   // "Active Jobs" | "Completed"
      stage,
      timeline,
      startDate,
      endDate,
      crew: crewText ? crewText.split(',').map((s) => s.trim()).filter(Boolean) : [],
      assignedSubs,
      contractValue,
      notes,
      sharePointUrl,
      panels: m.panels ?? null,
      panelsComplete: m.panelsComplete ?? null,
      panelsPct: (m.panels && m.panelsComplete != null)
        ? Math.round((m.panelsComplete / m.panels) * 100)
        : null,
      areaSqm: m.areaSqm ?? null,
    };
  });

  return { jobs, fetchedAt: new Date().toISOString() };
}

function gql(apiKey, query) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ query });
    const req = https.request(
      {
        hostname: 'api.monday.com',
        path: '/v2',
        method: 'POST',
        headers: {
          Authorization: apiKey,
          'Content-Type': 'application/json',
          'API-Version': '2024-10',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => {
          try { resolve(JSON.parse(data)); }
          catch (e) { reject(new Error(`Monday parse error: ${e.message} body=${data.slice(0, 200)}`)); }
        });
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

module.exports = { fetchJobs };
