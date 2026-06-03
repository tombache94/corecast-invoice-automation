/**
 * One-off probe — verifies the MONDAY_API_KEY works and shows what columns +
 * sample items are on the Active Jobs board. Used to design the field mapping
 * for src/dashboard/sources/monday.js. Read-only.
 */

require('dotenv').config();
const https = require('https');

const API_KEY = process.env.MONDAY_API_KEY;
const BOARD_ID = 5027051539;

if (!API_KEY) {
  console.error('Missing MONDAY_API_KEY in .env');
  process.exit(1);
}

function gql(query, variables = {}) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ query, variables });
    const req = https.request(
      {
        hostname: 'api.monday.com',
        path: '/v2',
        method: 'POST',
        headers: {
          Authorization: API_KEY,
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(body),
          'API-Version': '2024-10',
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => {
          try {
            resolve({ status: res.statusCode, body: JSON.parse(data) });
          } catch {
            resolve({ status: res.statusCode, body: data });
          }
        });
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function main() {
  // 1. Identity / token sanity
  const me = await gql(`{ me { id name email account { id name } } }`);
  if (me.body.errors) {
    console.error('❌ Token rejected:', JSON.stringify(me.body.errors, null, 2));
    process.exit(2);
  }
  console.log('✅ Token works');
  console.log('   User    :', me.body.data.me.name, `<${me.body.data.me.email}>`);
  console.log('   Account :', me.body.data.me.account.name, `(id=${me.body.data.me.account.id})`);
  console.log();

  // 2. Board metadata + columns
  const board = await gql(
    `query($id: [ID!]) {
      boards(ids: $id) {
        id
        name
        description
        items_count
        columns { id title type settings_str }
        groups { id title }
      }
    }`,
    { id: [String(BOARD_ID)] },
  );
  if (board.body.errors || !board.body.data.boards?.[0]) {
    console.error('❌ Board fetch failed:', JSON.stringify(board.body, null, 2));
    process.exit(3);
  }
  const b = board.body.data.boards[0];
  console.log(`✅ Board: ${b.name} (id=${b.id})`);
  console.log(`   Items: ${b.items_count}`);
  console.log(`   Groups: ${b.groups.map((g) => `${g.title}[${g.id}]`).join(', ')}`);
  console.log();
  console.log('   Columns:');
  for (const c of b.columns) {
    console.log(`     - ${c.title.padEnd(30)} ${c.type.padEnd(15)} id=${c.id}`);
  }
  console.log();

  // 3. Sample items with all column values
  const items = await gql(
    `query($id: ID!) {
      boards(ids: [$id]) {
        items_page(limit: 50) {
          cursor
          items {
            id
            name
            group { title }
            column_values {
              id
              text
              value
              column { title type }
            }
          }
        }
      }
    }`,
    { id: String(BOARD_ID) },
  );
  if (items.body.errors) {
    console.error('❌ Items fetch failed:', JSON.stringify(items.body.errors, null, 2));
    process.exit(4);
  }
  const list = items.body.data.boards[0].items_page.items;
  console.log(`✅ Fetched ${list.length} items. Showing field shape per item:\n`);
  for (const it of list) {
    console.log(`── ${it.name}  [group: ${it.group.title}]`);
    for (const cv of it.column_values) {
      const text = cv.text === null ? '∅' : (cv.text.length > 80 ? cv.text.slice(0, 80) + '…' : cv.text);
      console.log(`     ${cv.column.title.padEnd(28)} (${cv.column.type.padEnd(12)}) → ${text}`);
    }
    console.log();
  }
}

main().catch((err) => {
  console.error('Fatal:', err.message);
  process.exit(99);
});
