/**
 * Push the freshly-spliced dashboard HTML to Cloudflare Pages via the
 * Direct Upload REST API. No `wrangler` dependency — works on any platform
 * (including Windows ARM64 where workerd isn't supported).
 *
 * Flow:
 *   1. GET an upload JWT from /pages/projects/{name}/upload-token
 *   2. SHA-256 the file (truncated to 32 hex chars — Pages convention)
 *   3. POST /pages/assets/check-missing to see if it needs uploading
 *   4. POST /pages/assets/upload (base64) for missing assets
 *   5. POST /pages/projects/{name}/deployments with the manifest
 *
 * Requires CLOUDFLARE_API_TOKEN + CLOUDFLARE_ACCOUNT_ID in .env. The token
 * needs Account → Cloudflare Pages → Edit permission.
 *
 * Best-effort: returns { ok, message, url? } and never throws — caller
 * decides whether to surface the failure.
 */

const fs = require('fs');
const crypto = require('crypto');

const CF_API = 'https://api.cloudflare.com/client/v4';

function pagesAssetHash(buf) {
  // Cloudflare Pages keys assets by SHA-256(content) truncated to 32 hex
  // characters (matches wrangler's internal behaviour).
  return crypto.createHash('sha256').update(buf).digest('hex').slice(0, 32);
}

async function cfFetch(url, opts) {
  const res = await fetch(url, opts);
  const text = await res.text();
  let body;
  try { body = JSON.parse(text); }
  catch { return { httpStatus: res.status, raw: text, success: false }; }
  body.httpStatus = res.status;
  return body;
}

async function getUploadJWT(token, accountId, projectName) {
  const j = await cfFetch(
    `${CF_API}/accounts/${accountId}/pages/projects/${projectName}/upload-token`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (!j.success) throw new Error(`upload-token failed (${j.httpStatus}): ${JSON.stringify(j.errors || j.raw)}`);
  return j.result.jwt;
}

async function checkMissing(jwt, hashes) {
  const j = await cfFetch(`${CF_API}/pages/assets/check-missing`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${jwt}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ hashes }),
  });
  if (!j.success) throw new Error(`check-missing failed (${j.httpStatus}): ${JSON.stringify(j.errors || j.raw)}`);
  return j.result;
}

async function uploadAssets(jwt, payload) {
  const j = await cfFetch(`${CF_API}/pages/assets/upload`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${jwt}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!j.success) throw new Error(`upload failed (${j.httpStatus}): ${JSON.stringify(j.errors || j.raw)}`);
  return j.result;
}

async function createDeployment(token, accountId, projectName, manifest) {
  // Multipart form using the global FormData/Blob (Node 18+).
  const form = new FormData();
  form.append('manifest', JSON.stringify(manifest));
  form.append('branch', 'main');
  const j = await cfFetch(
    `${CF_API}/accounts/${accountId}/pages/projects/${projectName}/deployments`,
    { method: 'POST', headers: { Authorization: `Bearer ${token}` }, body: form },
  );
  if (!j.success) throw new Error(`deployment failed (${j.httpStatus}): ${JSON.stringify(j.errors || j.raw)}`);
  return j.result;
}

async function deployToCloudflare({ htmlPath } = {}) {
  const token       = process.env.CLOUDFLARE_API_TOKEN;
  const accountId   = process.env.CLOUDFLARE_ACCOUNT_ID;
  const projectName = process.env.CLOUDFLARE_PAGES_PROJECT || 'corecast-dashboard';

  if (!token || !accountId) {
    return { ok: false, message: 'Missing CLOUDFLARE_API_TOKEN or CLOUDFLARE_ACCOUNT_ID in .env' };
  }
  if (!fs.existsSync(htmlPath)) {
    return { ok: false, message: `Source HTML not found: ${htmlPath}` };
  }

  try {
    const buf = fs.readFileSync(htmlPath);
    // Refuse to publish a truncated/broken file. render.js has the same
    // guard for the splice path; we mirror it here so direct deploys
    // (without a refresh) can't push a broken page either.
    const tail = buf.toString('utf8', Math.max(0, buf.length - 200));
    if (!tail.trimEnd().endsWith('</html>')) {
      return {
        ok: false,
        message: 'refusing to deploy: source HTML does not end with </html> (likely truncated). Last 120 chars: ' + JSON.stringify(tail.slice(-120)),
      };
    }
    const hash = pagesAssetHash(buf);
    const manifest = { '/index.html': hash };

    const jwt = await getUploadJWT(token, accountId, projectName);
    const missing = await checkMissing(jwt, [hash]);

    if (missing.includes(hash)) {
      await uploadAssets(jwt, [{
        key: hash,
        value: buf.toString('base64'),
        metadata: { contentType: 'text/html' },
        base64: true,
      }]);
    }

    const dep = await createDeployment(token, accountId, projectName, manifest);
    // Pages returns a per-deployment preview URL plus the project's stable URL.
    const stableUrl = `https://${projectName}.pages.dev`;
    return {
      ok: true,
      message: missing.includes(hash) ? 'uploaded + deployed' : 'unchanged content, redeployed',
      url: stableUrl,
      previewUrl: dep.url || null,
      deploymentId: dep.id || null,
    };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

module.exports = { deployToCloudflare };
