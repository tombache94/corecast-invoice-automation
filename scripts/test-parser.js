const { parseInvoiceMessage } = require('../src/email/parser');

// In tests, "PDF text" is just utf8 in contentBytes — production uses pdf-parse on real PDF bytes.
const mockExtractPdf = async (buf) => buf.toString('utf8');

const samples = [
  {
    label: 'Sample 1 — HTML body, all inline (Bob Smith Formwork)',
    message: {
      from: { emailAddress: { address: 'bob@smithformwork.com.au', name: 'Bob Smith' } },
      subject: 'Invoice 0123 — Smith Formwork',
      receivedDateTime: '2026-04-29T08:00:00Z',
      body: {
        contentType: 'html',
        content: `
          <p>Hi Tom,</p>
          <p>Labour for week ending 29/04/2026:</p>
          <table>
            <tr><td>Hours worked</td><td>40 hours</td></tr>
            <tr><td>Rate</td><td>$75/hr</td></tr>
            <tr><td>Subtotal</td><td>$3,000.00</td></tr>
            <tr><td>GST</td><td>$300.00</td></tr>
            <tr><td><strong>Total</strong></td><td><strong>AU$3,300.00</strong></td></tr>
          </table>
          <p>Cheers,<br>Bob</p>
        `,
      },
    },
    expect: {
      senderEmail: 'bob@smithformwork.com.au',
      amount: 3300,
      hoursClaimed: 40,
      // 29/04/2026 is a Wednesday → nearest Friday is 2026-05-01
      weekEnding: '2026-05-01',
    },
  },
  {
    label: 'Sample 2 — Plain text, "38h" bare suffix (Dave Wilson Electrical)',
    message: {
      from: { emailAddress: { address: 'dave@daveselectrical.com.au', name: '' } },
      subject: 'Inv 17 — week ending 3 May',
      receivedDateTime: '2026-05-03T19:30:00Z',
      body: {
        contentType: 'text',
        content:
          'Hi mate,\n\n' +
          'Week ending 3 May 2026.\n\n' +
          'Hours: 38h\n' +
          'Rate: $80/hour\n' +
          'Total: $3,040.00\n\n' +
          'Cheers,\nDave Wilson',
      },
    },
    expect: {
      senderEmail: 'dave@daveselectrical.com.au',
      amount: 3040,
      hoursClaimed: 38,
      date: '2026-05-03',
      // 3 May 2026 is a Sunday → nearest Friday is 2026-05-01
      weekEnding: '2026-05-01',
    },
  },
  {
    label: 'Sample 3 — Body says "see attached", PDF attachment carries the invoice',
    message: {
      from: { emailAddress: { address: 'accounts@acmeconcrete.com.au', name: 'Acme Concrete Pumping' } },
      subject: 'Tax Invoice 0042',
      receivedDateTime: '2026-05-02T11:00:00Z',
      body: { contentType: 'text', content: 'Hi, please see attached invoice. Cheers.' },
      attachments: [
        {
          name: 'INV-0042.pdf',
          contentType: 'application/pdf',
          contentBytes: Buffer.from(
            'TAX INVOICE\n' +
              'Acme Concrete Pumping Pty Ltd\n' +
              'ABN: 12 345 678 901\n\n' +
              'Invoice #: 0042\n' +
              'Invoice Date: 1/05/2026\n\n' +
              'Description: Concrete pumping labour\n' +
              '18 hours @ $120/hr\n' +
              'Subtotal: $2,160.00\n' +
              'GST: $216.00\n' +
              'Total Amount Due: $2,376.00\n',
            'utf8',
          ).toString('base64'),
        },
      ],
    },
    expect: {
      sender: 'Acme Concrete Pumping',
      senderEmail: 'accounts@acmeconcrete.com.au',
      amount: 2376,
      hoursClaimed: 18,
      date: '2026-05-01',
      // 1 May 2026 is Friday → weekEnding is the same day
      weekEnding: '2026-05-01',
    },
  },
];

function fmt(v) {
  if (v === null || v === undefined) return '(null)';
  return String(v);
}

(async () => {
  let failed = 0;
  for (const s of samples) {
    console.log('\n' + '='.repeat(70));
    console.log(s.label);
    console.log('='.repeat(70));
    const result = await parseInvoiceMessage(s.message, { extractPdf: mockExtractPdf });
    console.log('  sender         :', fmt(result.sender), `[conf ${result.confidence.sender}]`);
    console.log('  senderEmail    :', fmt(result.senderEmail), `[conf ${result.confidence.senderEmail}]`);
    console.log('  amount         :', fmt(result.amount), `[conf ${result.confidence.amount} via ${result.sources.amount}]`);
    console.log('  date           :', fmt(result.date), `[conf ${result.confidence.date} via ${result.sources.date}]`);
    console.log('  hoursClaimed   :', fmt(result.hoursClaimed), `[conf ${result.confidence.hoursClaimed} via ${result.sources.hoursClaimed}]`);
    console.log('  weekEnding     :', fmt(result.weekEnding), `[conf ${result.confidence.weekEnding}]`);
    console.log('  rawContent[..160]:', result.rawContent.slice(0, 160).replace(/\n/g, ' '));

    for (const [k, exp] of Object.entries(s.expect)) {
      const actual = result[k];
      const ok = actual === exp;
      if (!ok) {
        failed++;
        console.log(`  ✗ FAIL ${k}: expected ${fmt(exp)}, got ${fmt(actual)}`);
      } else {
        console.log(`  ✓ ${k} = ${fmt(actual)}`);
      }
    }
  }
  console.log('\n' + '='.repeat(70));
  if (failed) {
    console.log(`RESULT: ${failed} assertion(s) failed`);
    process.exit(1);
  } else {
    console.log('RESULT: all assertions passed');
  }
})();
