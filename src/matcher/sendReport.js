const STATUS_LABELS = {
  match: { emoji: '✅', label: 'Match' },
  mismatch: { emoji: '⚠️', label: 'MISMATCH' },
  not_found: { emoji: '❓', label: 'Not found in timesheet' },
  no_hours_in_timesheet: { emoji: '⏸️', label: 'No hours clocked' },
  invoice_hours_missing: { emoji: '❓', label: 'Invoice hours missing' },
};

function fmtHours(h) {
  if (h == null) return '—';
  return `${h.toFixed(2)}h`;
}

function fmtMoney(n) {
  if (n == null) return '—';
  return `$${Number(n).toFixed(2)}`;
}

function fmtDiff(d) {
  if (d == null) return '';
  const sign = d > 0 ? '+' : '';
  return `${sign}${d.toFixed(2)}h`;
}

function escapeHtml(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildSubject({ weekStart }) {
  return `[CoreCast] Timesheet Validation — Week of ${weekStart}`;
}

function buildPlainText(matchResult) {
  const { weekStart, weekEnd, results, counts, tolerance, timesheetPath, bills } = matchResult;
  const lines = [];

  lines.push(`CoreCast Timesheet Validation`);
  lines.push(`Week: ${weekStart} to ${weekEnd}`);
  lines.push(`Tolerance: ±${tolerance}h`);
  lines.push('');

  lines.push(`=== LABOUR INVOICES (${results.length}) ===`);
  if (!results.length) {
    lines.push('No labour invoices found for this week.');
  } else {
    for (const r of results) {
      const lbl = STATUS_LABELS[r.classification] || { emoji: '?', label: r.classification };
      lines.push(`${lbl.emoji} ${r.sender || '(unknown sender)'} — ${lbl.label}`);
      lines.push(`    Invoice claimed: ${fmtHours(r.hoursClaimed)} / ${fmtMoney(r.amount)}`);
      lines.push(`    Connecteam:      ${fmtHours(r.timesheetHours)}`);
      if (r.classification === 'mismatch') {
        lines.push(`    Difference:      ${fmtDiff(r.diff)}`);
      }
      lines.push(`    Week ending:     ${r.weekEnding}`);
      lines.push('');
    }
  }

  if (bills) {
    const created = bills.results.filter((b) => b.result.created);
    const duplicates = bills.results.filter((b) => b.result.reason === 'duplicate');
    const flagged = bills.results.filter((b) => b.result.flag === 'manual_review');
    const errors = bills.results.filter((b) => b.result.error);
    const skippedNoXero = bills.results.filter((b) => b.result.reason === 'no_xero_client');

    lines.push(`=== BILLS CREATED IN XERO (${created.length}) ===`);
    if (!bills.results.length) {
      lines.push('No bills this week.');
    } else if (skippedNoXero.length === bills.results.length) {
      lines.push('Xero client not configured — bills not pushed:');
      for (const b of skippedNoXero) {
        const acct = b.result.category?.accountName || '?';
        lines.push(`  ${pad(b.sender || '(unknown)', 22)} ${pad(fmtMoney(b.amount), 12)} → ${acct}`);
      }
    } else {
      for (const b of created) {
        lines.push(
          `  ${pad(b.sender || '(unknown)', 22)} ${pad(fmtMoney(b.amount), 12)} → ${pad(b.result.accountName, 26)} [DRAFT in Xero]`,
        );
      }
      if (duplicates.length) {
        lines.push('');
        lines.push(`⚠️  ${duplicates.length} bill(s) skipped (already in Xero):`);
        for (const b of duplicates) {
          lines.push(
            `  ${pad(b.sender || '(unknown)', 22)} ${pad(fmtMoney(b.amount), 12)} → existing: ${b.result.xeroNumber || b.result.xeroId}`,
          );
        }
      }
      if (flagged.length) {
        lines.push('');
        lines.push(`⚠️  ${flagged.length} bill(s) flagged for manual review:`);
        for (const b of flagged) {
          const detail = b.result.detail || b.result.reason;
          lines.push(`  ${pad(b.sender || '(unknown)', 22)} ${pad(fmtMoney(b.amount), 12)} → ${detail}`);
        }
      }
      if (errors.length) {
        lines.push('');
        lines.push(`❌ ${errors.length} bill(s) errored:`);
        for (const b of errors) {
          lines.push(`  ${pad(b.sender || '(unknown)', 22)} ${pad(fmtMoney(b.amount), 12)} → ${b.result.error}`);
        }
      }
    }
    lines.push('');
  }

  lines.push('--- Summary ---');
  lines.push(`Labour invoices:    ${counts.total}`);
  lines.push(`  Matches:          ${counts.match}`);
  lines.push(`  Mismatches:       ${counts.mismatch}`);
  lines.push(`  Not in timesheet: ${counts.not_found}`);
  lines.push(`  No hours clocked: ${counts.no_hours_in_timesheet}`);
  lines.push(`  Hrs not parsed:   ${counts.invoice_hours_missing}`);
  if (bills) {
    lines.push(`Bills:              ${bills.counts.total}`);
    lines.push(`  Created in Xero:  ${bills.counts.created}`);
    lines.push(`  Duplicates:       ${bills.counts.duplicates}`);
    lines.push(`  Flagged:          ${bills.counts.flagged}`);
    lines.push(`  Errors:           ${bills.counts.errors}`);
  }
  lines.push('');
  lines.push(`Source: ${timesheetPath}`);

  return lines.join('\n');
}

function pad(s, n) {
  const str = String(s ?? '');
  if (str.length >= n) return str;
  return str + ' '.repeat(n - str.length);
}

function buildHtml(matchResult) {
  const { weekStart, weekEnd, results, counts, tolerance, timesheetPath, bills } = matchResult;

  const rowColor = (cls) =>
    cls === 'match'
      ? '#e8f5e9'
      : cls === 'mismatch'
        ? '#fff3e0'
        : '#f5f5f5';

  const rowsHtml = results.length
    ? results
        .map((r) => {
          const lbl = STATUS_LABELS[r.classification] || { emoji: '?', label: r.classification };
          const diffCell =
            r.classification === 'mismatch'
              ? `<strong>${escapeHtml(fmtDiff(r.diff))}</strong>`
              : '';
          return `
<tr style="background:${rowColor(r.classification)};">
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(r.sender || '(unknown)')}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(fmtHours(r.hoursClaimed))}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(fmtMoney(r.amount))}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(fmtHours(r.timesheetHours))}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${diffCell}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${lbl.emoji} ${escapeHtml(lbl.label)}</td>
</tr>`;
        })
        .join('')
    : `<tr><td colspan="6" style="padding:12px;text-align:center;color:#666;">No invoices found for this week.</td></tr>`;

  const billsSection = bills ? buildBillsHtml(bills) : '';
  const billsSummary = bills
    ? `
    <li>Bills total: <strong>${bills.counts.total}</strong></li>
    <li>Created in Xero: <strong>${bills.counts.created}</strong></li>
    <li>Duplicates: <strong>${bills.counts.duplicates}</strong></li>
    <li>Flagged: <strong>${bills.counts.flagged}</strong></li>
    <li>Errors: <strong>${bills.counts.errors}</strong></li>`
    : '';

  return `<!DOCTYPE html>
<html>
<body style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#222;max-width:780px;margin:0 auto;padding:20px;">
  <h2 style="margin:0 0 4px 0;">CoreCast Timesheet Validation</h2>
  <p style="margin:0 0 4px 0;color:#555;">Week: <strong>${escapeHtml(weekStart)}</strong> to <strong>${escapeHtml(weekEnd)}</strong></p>
  <p style="margin:0 0 20px 0;color:#888;font-size:13px;">Tolerance: ±${escapeHtml(tolerance)}h</p>

  <h3 style="margin:0 0 8px 0;">Labour invoices (${results.length})</h3>
  <table style="border-collapse:collapse;width:100%;font-size:14px;">
    <thead>
      <tr style="background:#f0f0f0;text-align:left;">
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Subcontractor</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Claimed</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Amount</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Connecteam</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Diff</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Status</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>

  ${billsSection}

  <h3 style="margin-top:28px;margin-bottom:8px;">Summary</h3>
  <ul style="line-height:1.6;margin:0;padding-left:18px;">
    <li>Labour total: <strong>${counts.total}</strong></li>
    <li>Matches: <strong>${counts.match}</strong></li>
    <li>Mismatches: <strong>${counts.mismatch}</strong></li>
    <li>Not in timesheet: <strong>${counts.not_found}</strong></li>
    <li>No hours clocked: <strong>${counts.no_hours_in_timesheet}</strong></li>
    <li>Invoice hours missing: <strong>${counts.invoice_hours_missing}</strong></li>${billsSummary}
  </ul>

  <p style="margin-top:24px;color:#888;font-size:12px;">
    Source: ${escapeHtml(timesheetPath)}
  </p>
</body>
</html>`;
}

function buildBillsHtml(bills) {
  if (!bills.results.length) {
    return `<h3 style="margin-top:28px;margin-bottom:8px;">Bills</h3><p style="color:#666;">No bills this week.</p>`;
  }

  const rowColor = (r) => {
    if (r.error) return '#ffebee';
    if (r.created) return '#e8f5e9';
    if (r.reason === 'duplicate') return '#fffde7';
    if (r.flag === 'manual_review') return '#fff3e0';
    return '#f5f5f5';
  };

  const statusLabel = (r) => {
    if (r.error) return `❌ Error`;
    if (r.created) return `✅ Draft in Xero`;
    if (r.reason === 'duplicate') return `↩️ Duplicate`;
    if (r.flag === 'manual_review') return `⚠️ Manual review`;
    if (r.reason === 'no_xero_client') return `⏸️ Not pushed`;
    if (r.dryRun) return `🔍 Dry-run`;
    return r.reason || '—';
  };

  const rows = bills.results
    .map((b) => {
      const r = b.result;
      const account = r.accountName || r.category?.accountName || r.wouldCreate?.accountName || '—';
      const detail = r.error || r.detail || r.xeroNumber || r.xeroId || '';
      return `
<tr style="background:${rowColor(r)};">
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(b.sender || '(unknown)')}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(fmtMoney(b.amount))}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(account)}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;">${escapeHtml(statusLabel(r))}</td>
  <td style="padding:8px 12px;border-bottom:1px solid #ddd;color:#666;font-size:12px;">${escapeHtml(detail)}</td>
</tr>`;
    })
    .join('');

  return `
  <h3 style="margin-top:28px;margin-bottom:8px;">Bills (${bills.counts.total})</h3>
  <table style="border-collapse:collapse;width:100%;font-size:14px;">
    <thead>
      <tr style="background:#f0f0f0;text-align:left;">
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Supplier</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Amount</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Account</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Status</th>
        <th style="padding:8px 12px;border-bottom:2px solid #aaa;">Detail</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

/**
 * Send the weekly validation report email via Microsoft Graph.
 *
 * @param {object} matchResult - Result from matchWeek()
 * @param {object} opts
 * @param {object} opts.outlook   - OutlookClient instance
 * @param {string} opts.sender    - From mailbox UPN
 * @param {string|string[]} opts.recipient - To address(es)
 * @returns {Promise<{subject:string,text:string,html:string}>}
 */
async function sendReport(matchResult, { outlook, sender, recipient }) {
  if (!outlook) throw new Error('sendReport: outlook is required');
  if (!sender) throw new Error('sendReport: sender is required');
  if (!recipient) throw new Error('sendReport: recipient is required');

  const subject = buildSubject(matchResult);
  const text = buildPlainText(matchResult);
  const html = buildHtml(matchResult);

  await outlook.sendMail(sender, { to: recipient, subject, text, html });

  return { subject, text, html };
}

module.exports = { sendReport, buildSubject, buildPlainText, buildHtml };
