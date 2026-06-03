/**
 * One-time Xero OAuth2 authentication script.
 * Run this once to get a refresh token, which is saved to data/xero-token.json
 * and used automatically by the Xero pipeline from then on.
 *
 * Usage: node scripts/xero-auth.js
 */

require('dotenv').config();
const http = require('http');
const https = require('https');
const url = require('url');
const fs = require('fs');
const path = require('path');

const CLIENT_ID = process.env.XERO_CLIENT_ID;
const CLIENT_SECRET = process.env.XERO_CLIENT_SECRET;
const REDIRECT_URI = process.env.XERO_REDIRECT_URI || 'http://localhost:5000/callback';
const TOKEN_FILE = path.resolve(__dirname, '../data/xero-token.json');
// Scopes we actually use:
//   offline_access            — required to get a refresh token
//   accounting.transactions   — read/write bills (ACCPAY invoices)
//   accounting.contacts       — find/create supplier contacts
//   accounting.settings.read  — read TrackingCategories (needed by tag-tracking-categories.js)
// Override via XERO_SCOPES in .env if your Developer Portal app exposes a different set.
// Note: we deliberately do NOT request `openid profile` — those require the app to be
// configured for OpenID Connect, which causes invalid_scope errors on accounting-only apps.
const SCOPES = process.env.XERO_SCOPES ||
  'offline_access accounting.transactions accounting.contacts accounting.attachments';

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('Missing XERO_CLIENT_ID or XERO_CLIENT_SECRET in .env');
  process.exit(1);
}

// Build the Xero authorisation URL
const authUrl = `https://login.xero.com/identity/connect/authorize?` +
  `response_type=code` +
  `&client_id=${CLIENT_ID}` +
  `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
  `&scope=${encodeURIComponent(SCOPES)}` +
  `&state=corecast`;

console.log('\n=== Xero OAuth2 Setup ===\n');
console.log('Opening your browser to log into Xero...');
console.log('If it does not open automatically, paste this URL into your browser:\n');
console.log(authUrl);
console.log('\nWaiting for Xero to redirect back...\n');

// Try to open browser automatically
const { exec } = require('child_process');
exec(`start "" "${authUrl}"`);

// Start a local server to catch the OAuth callback
const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  if (parsed.pathname !== '/callback' && parsed.pathname !== '/callback/') {
    res.end('Not found');
    return;
  }

  const code = parsed.query.code;
  if (!code) {
    res.end('<h2>Error: no code returned from Xero</h2>');
    console.error('No code in callback. Try again.');
    server.close();
    return;
  }

  console.log('Got authorisation code — exchanging for tokens...');

  // Exchange code for tokens
  const body = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
  }).toString();

  const credentials = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');

  const tokenReq = https.request({
    hostname: 'identity.xero.com',
    path: '/connect/token',
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Authorization': `Basic ${credentials}`,
      'Content-Length': Buffer.byteLength(body),
    },
  }, (tokenRes) => {
    let data = '';
    tokenRes.on('data', chunk => data += chunk);
    tokenRes.on('end', async () => {
      const tokens = JSON.parse(data);

      if (tokens.error) {
        res.end(`<h2>Error: ${tokens.error_description || tokens.error}</h2>`);
        console.error('Token error:', tokens);
        server.close();
        return;
      }

      // Get tenant ID
      const tenantsReq = https.request({
        hostname: 'api.xero.com',
        path: '/connections',
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${tokens.access_token}`,
          'Content-Type': 'application/json',
        },
      }, (tenantsRes) => {
        let tData = '';
        tenantsRes.on('data', c => tData += c);
        tenantsRes.on('end', () => {
          const tenants = JSON.parse(tData);
          const tenant = tenants[0];

          const tokenData = {
            access_token: tokens.access_token,
            refresh_token: tokens.refresh_token,
            expires_in: tokens.expires_in,
            token_type: tokens.token_type,
            scope: tokens.scope,
            tenant_id: tenant?.tenantId,
            tenant_name: tenant?.tenantName,
            saved_at: new Date().toISOString(),
          };

          // Ensure data directory exists
          fs.mkdirSync(path.dirname(TOKEN_FILE), { recursive: true });
          fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokenData, null, 2));

          console.log('\n✅ Success! Xero connected.');
          console.log(`   Tenant: ${tenant?.tenantName} (${tenant?.tenantId})`);
          console.log(`   Token saved to: ${TOKEN_FILE}`);
          console.log('\nYou can now run:');
          console.log('   node scripts/test-xero-pipeline.js --live\n');

          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <h2>✅ Xero connected successfully!</h2>
            <p>Tenant: <strong>${tenant?.tenantName}</strong></p>
            <p>You can close this window and return to the terminal.</p>
          `);

          server.close();
        });
      });
      tenantsReq.end();
    });
  });

  tokenReq.write(body);
  tokenReq.end();
});

const PORT = parseInt((process.env.XERO_REDIRECT_URI || 'http://localhost:5000').split(':')[2]) || 5000;
server.listen(PORT, () => {
  console.log(`Listening on port ${PORT} for Xero callback...`);
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error('Port 3000 is already in use. Close whatever is using it and try again.');
  } else {
    console.error('Server error:', err.message);
  }
  process.exit(1);
});
