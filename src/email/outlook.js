require('isomorphic-fetch');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

class OutlookClient {
  constructor({ tenantId, clientId, clientSecret }) {
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

  async listRecentMessages(mailbox, { sinceISO, top = 50 } = {}) {
    let req = this._graph
      .api(`/users/${mailbox}/mailFolders/Inbox/messages`)
      .top(top)
      .orderby('receivedDateTime DESC')
      .select('id,subject,from,receivedDateTime,hasAttachments,bodyPreview,internetMessageId');

    if (sinceISO) req = req.filter(`receivedDateTime ge ${sinceISO}`);

    const res = await req.get();
    return res.value.map((m) => ({
      id: m.id,
      mailbox,
      internetMessageId: m.internetMessageId,
      receivedAt: m.receivedDateTime,
      from: m.from?.emailAddress?.address || null,
      fromName: m.from?.emailAddress?.name || null,
      subject: m.subject || '',
      hasAttachments: !!m.hasAttachments,
      preview: m.bodyPreview || '',
    }));
  }

  async getMessage(mailbox, messageId, { includeAttachments = false } = {}) {
    const message = await this._graph.api(`/users/${mailbox}/messages/${messageId}`).get();
    if (includeAttachments && message.hasAttachments) {
      const att = await this._graph.api(`/users/${mailbox}/messages/${messageId}/attachments`).get();
      message.attachments = att.value;
    }
    return message;
  }

  async sendMail(sender, { to, subject, text, html }) {
    const recipients = (Array.isArray(to) ? to : [to]).filter(Boolean).map((address) => ({
      emailAddress: { address },
    }));
    if (!recipients.length) throw new Error('sendMail: at least one recipient required');

    const message = {
      subject: subject || '(no subject)',
      body: html
        ? { contentType: 'HTML', content: html }
        : { contentType: 'Text', content: text || '' },
      toRecipients: recipients,
    };

    await this._graph.api(`/users/${sender}/sendMail`).post({
      message,
      saveToSentItems: true,
    });
  }
}

async function listAllInboxes(outlook, mailboxes, opts) {
  const results = await Promise.all(
    mailboxes.map(async (mb) => {
      try {
        const msgs = await outlook.listRecentMessages(mb, opts);
        return { mailbox: mb, ok: true, messages: msgs };
      } catch (err) {
        return { mailbox: mb, ok: false, error: err.message, messages: [] };
      }
    }),
  );
  return results;
}

module.exports = { OutlookClient, listAllInboxes };
