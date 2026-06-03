require('isomorphic-fetch');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

// 4 MB — Graph's hard cap for a single PUT to /content. Above this we
// must use an upload session (chunked).
const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;
// Chunk size for upload sessions. Must be a multiple of 320 KiB per Graph docs.
const CHUNK_SIZE = 5 * 320 * 1024; // 1.6 MB

function encodePath(p) {
  return p
    .split('/')
    .map((seg) => encodeURIComponent(seg))
    .join('/');
}

class SharePointClient {
  constructor({ tenantId, clientId, clientSecret, driveId }) {
    if (!driveId) throw new Error('SharePointClient: driveId required');
    this.driveId = driveId;
    this.msal = new ConfidentialClientApplication({
      auth: {
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    });
    this._graph = Client.init({
      authProvider: async (done) => {
        try {
          const result = await this.msal.acquireTokenByClientCredential({
            scopes: ['https://graph.microsoft.com/.default'],
          });
          done(null, result.accessToken);
        } catch (err) {
          done(err, null);
        }
      },
    });
  }

  async _accessToken() {
    const result = await this.msal.acquireTokenByClientCredential({
      scopes: ['https://graph.microsoft.com/.default'],
    });
    return result.accessToken;
  }

  async listChildren(folderPath) {
    return this._graph
      .api(`/drives/${this.driveId}/root:/${encodePath(folderPath)}:/children`)
      .select('name,folder,file,size,lastModifiedDateTime,webUrl')
      .top(200)
      .get();
  }

  async folderExists(folderPath) {
    try {
      await this._graph
        .api(`/drives/${this.driveId}/root:/${encodePath(folderPath)}`)
        .select('id,name,folder')
        .get();
      return true;
    } catch (err) {
      if (err.statusCode === 404) return false;
      throw err;
    }
  }

  async ensureFolder(parentPath, name) {
    const fullPath = `${parentPath}/${name}`;
    if (await this.folderExists(fullPath)) return fullPath;

    await this._graph
      .api(`/drives/${this.driveId}/root:/${encodePath(parentPath)}:/children`)
      .post({
        name,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      });
    return fullPath;
  }

  async uploadFile(folderPath, filename, content) {
    const buf = Buffer.isBuffer(content) ? content : Buffer.from(content);
    return buf.byteLength <= SIMPLE_UPLOAD_LIMIT
      ? this._uploadSmall(folderPath, filename, buf)
      : this._uploadLarge(folderPath, filename, buf);
  }

  async _uploadSmall(folderPath, filename, buf) {
    const target = `${folderPath}/${filename}`;
    return this._graph
      .api(`/drives/${this.driveId}/root:/${encodePath(target)}:/content`)
      .header('Content-Type', 'application/octet-stream')
      .put(buf);
  }

  async _uploadLarge(folderPath, filename, buf) {
    const target = `${folderPath}/${filename}`;
    const session = await this._graph
      .api(`/drives/${this.driveId}/root:/${encodePath(target)}:/createUploadSession`)
      .post({
        item: {
          '@microsoft.graph.conflictBehavior': 'rename',
          name: filename,
        },
      });

    const uploadUrl = session.uploadUrl;
    const total = buf.byteLength;
    let offset = 0;
    let lastResponse = null;

    while (offset < total) {
      const end = Math.min(offset + CHUNK_SIZE, total);
      const chunk = buf.subarray(offset, end);
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(chunk.byteLength),
          'Content-Range': `bytes ${offset}-${end - 1}/${total}`,
        },
        body: chunk,
      });
      if (!res.ok && res.status !== 202) {
        const body = await res.text().catch(() => '');
        throw new Error(`Upload session chunk failed: ${res.status} ${res.statusText} ${body}`);
      }
      lastResponse = res.status === 200 || res.status === 201 ? await res.json() : null;
      offset = end;
    }

    return lastResponse;
  }
}

module.exports = { SharePointClient };
