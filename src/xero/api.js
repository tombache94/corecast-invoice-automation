/**
 * Shared Xero REST helper. Loads the cached refresh token, refreshes the
 * access token, and exposes a small `xeroRequest` for any path.
 *
 * Used by both the live tagging script and the dashboard sources so we don't
 * reimplement token rotation in every consumer.
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

const TOKEN_FILE = path.resolve(__dirname, '../../data/xero-token.json');

let cachedAccessToken = null;
let cachedTenantId = null;

function loadTokens() {
  if (!fs.existsSync(TOKEN_FILE)) {
    throw new Error(
      `No token file at ${TOKEN_FILE}. Run: node scripts/xero-auth.js`,
    );
  }
  return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
}

function saveTokens(tokens) {
  fs.writeFileSync(
    TOKEN_FILE,
    JSON.stringify({ ...tokens, saved_at: new Date().toISOString() }, null, 2),
  );
}

function refreshAccessToken(tokens) {
  const clientId = process.env.XERO_CLIENT_ID;
  const clientSecret = process.env.XERO_CLIENT_SECRET;
  if (!clientId || !clientSecret) {
    throw new Error('Missing XERO_CLIENT_ID / XERO_CLIENT_SECRET in .env');
  }
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: tokens.refresh_token,
  }).toString();
  const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        hostname: 'identity.xero.com',
        path: '/connect/token',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Authorization: `Basic ${credentials}`,
          'Content-Length': Buffer.byteLength(body),
        },
      },
      (res) => {
        let data = '';
        res.on('data', (c) => (data += c));
        res.on('end', () => resolve(JSON.parse(data)));
      },
    );
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

/**
 * Ensure we have a fresh access token. Caches in-process for the duration of
 * a single run so multiple requests don't each refresh.
 */
async function getAccessToken() {
  if (cachedAccessToken) return { accessToken: cachedAccessToken, tenantId: cachedTenantId };
  let tokens = loadTokens();
  const refreshed = await refreshAccessToken(tokens);
  if (refreshed.error) {
    throw new Error(`Xero token refresh failed: ${refreshed.error_description || refreshed.error}`);
  }
  tokens = { ...tokens, ...refreshed };
  saveTokens(tokens);
  cachedAccessToken = tokens.access_token;
  cachedTenantId = tokens.tenant_id || process.env.XERO_TENANT_ID;
  return { accessToken: cachedAccessToken, tenantId: cachedTenantId };
}

/**
 * Make an authenticated Xero REST call. Returns { status, body }.
 */
async function xeroRequest({ method = 'GET', path: apiPath, body }) {
  const { accessToken, tenantId } = await getAccessToken();
  const bodyStr = body ? JSON.stringify(body) : null;
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        hostname: 'api.xero.com',
        path: apiPath,
        method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'xero-tenant-id': tenantId,
          'Content-Type': 'application/json',
          Accept: 'application/json',
          ...(bodyStr ? { 'Content-Length': Buffer.byteLength(bodyStr) } : {}),
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
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

module.exports = { xeroRequest, getAccessToken };
