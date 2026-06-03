// Integration smoke test: proves the Outlook → parser → storage wiring works
// without hitting Microsoft Graph. Same code path as `node src/index.js ingest`,
// but `outlook` is replaced with a stub that returns canned messages.

const fs = require('fs/promises');
const path = require('path');
const { ingest } = require('../src/index');
const { InvoiceStore } = require('../src/storage/invoices');

const MAILBOX = 'tom@corecastconcrete.com.au';

const messages = [
  {
    id: 'msg-1',
    receivedDateTime: '2026-04-29T08:00:00Z',
    from: { emailAddress: { address: 'bob@smithformwork.com.au', name: 'Bob Smith' } },
    subject: 'Invoice 0123 — Smith Formwork',
    body: {
      contentType: 'html',
      content: `
        <p>Labour for week ending 29/04/2026:</p>
        <table>
          <tr><td>Hours worked</td><td>40 hours</td></tr>
          <tr><td>Subtotal</td><td>$3,000.00</td></tr>
          <tr><td>GST</td><td>$300.00</td></tr>
          <tr><td><strong>Total</strong></td><td><strong>AU$3,300.00</strong></td></tr>
        </table>`,
    },
    hasAttachments: false,
  },
  {
    id: 'msg-2',
    receivedDateTime: '2026-05-03T19:30:00Z',
    from: { emailAddress: { address: 'dave@daveselectrical.com.au', name: '' } },
    subject: 'Inv 17 — week ending 3 May',
    body: {
      contentType: 'text',
      content:
        'Hi mate,\n\nWeek ending 3 May 2026.\n\nHours: 38h\nRate: $80/hour\nTotal: $3,040.00\n\nCheers,\nDave',
    },
    hasAttachments: false,
  },
  {
    id: 'msg-3',
    receivedDateTime: '2026-05-02T11:00:00Z',
    from: { emailAddress: { address: 'accounts@acmeconcrete.com.au', name: 'Acme Concrete Pumping' } },
    subject: 'Tax Invoice 0042',
    body: { contentType: 'text', content: 'Hi, please see attached invoice. Cheers.' },
    hasAttachments: true,
    attachments: [
      {
        name: 'INV-0042.pdf',
        contentType: 'application/pdf',
        contentBytes: Buffer.from(
          'TAX INVOICE\nAcme Concrete Pumping Pty Ltd\nABN: 12 345 678 901\n\n' +
            'Invoice Date: 1/05/2026\n18 hours @ $120/hr\n' +
            'Subtotal: $2,160.00\nGST: $216.00\nTotal Amount Due: $2,376.00\n',
          'utf8',
        ).toString('base64'),
      },
    ],
  },
  {
    id: 'msg-4',
    receivedDateTime: '2026-05-01T07:00:00Z',
    from: { emailAddress: { address: 'newsletter@buildersnews.com.au', name: 'Builders News' } },
    subject: 'This week in concrete',
    body: { contentType: 'text', content: 'Top stories this week. No invoice content here at all.' },
    hasAttachments: false,
  },
  {
    id: 'msg-5',
    receivedDateTime: '2026-04-30T14:00:00Z',
    from: { emailAddress: { address: 'sketchy@subbie.com.au', name: 'Sketchy Subbie' } },
    subject: 'Invoice attached',
    body: {
      contentType: 'text',
      content: 'Hi Tom, $1,500 for jobs done this week. Cheers.',
    },
    hasAttachments: false,
  },
];

const stubOutlook = {
  async listRecentMessages(mailbox, { top }) {
    return messages.slice(0, top).map((m) => ({
      id: m.id,
      mailbox,
      receivedAt: m.receivedDateTime,
      from: m.from?.emailAddress?.address || null,
      fromName: m.from?.emailAddress?.name || null,
      subject: m.subject,
      hasAttachments: !!m.hasAttachments,
      preview: '',
    }));
  },
  async getMessage(mailbox, id) {
    const m = messages.find((x) => x.id === id);
    if (!m) throw new Error(`message not found: ${id}`);
    return m;
  },
};

(async () => {
  const tmpFile = path.resolve('./data/invoices.test-ingest.json');
  await fs.rm(tmpFile, { force: true });
  await fs.rm(tmpFile + '.lock', { recursive: true, force: true });

  const store = new InvoiceStore(tmpFile);
  await ingest({ mailbox: MAILBOX, limit: 5 }, { outlook: stubOutlook, store });

  console.log('\n=== Storage verification ===');
  const data = JSON.parse(await fs.readFile(tmpFile, 'utf8'));
  console.log(`  Records persisted: ${data.invoices.length}`);
  for (const inv of data.invoices) {
    console.log(
      `    [${inv.status.padEnd(13)}] ${inv.sender}  $${inv.amount}  ${inv.hoursClaimed}h  ` +
        `week ${inv.weekEnding}  (${inv.id})`,
    );
  }

  await fs.rm(tmpFile, { force: true });
  await fs.rm(tmpFile + '.lock', { recursive: true, force: true });
  console.log('\nOK');
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
