/**
 * Test the Connecteam XLSX timesheet parser.
 * Usage: node scripts/test-timesheet-parser.js <path-to-xlsx>
 *
 * Example:
 *   node scripts/test-timesheet-parser.js "C:\Users\tomba\Downloads\timeclock-timesheet_overview_2026-04-27_2026-05-03.xlsx"
 */

const path = require('path');
const { parseTimesheetExport, lookupHours } = require('../src/connecteam/timesheetParser');

const filePath = process.argv[2];
if (!filePath) {
  console.error('Usage: node scripts/test-timesheet-parser.js <path-to-xlsx>');
  process.exit(1);
}

try {
  const result = parseTimesheetExport(path.resolve(filePath));

  console.log(`\nTimesheet parsed successfully`);
  console.log(`Week: ${result.weekStart} → ${result.weekEnd}`);
  console.log(`\nEmployees found (${result.employees.length}):\n`);

  result.employees.forEach(e => {
    const bar = '█'.repeat(Math.round(e.totalHours / 2));
    console.log(`  ${e.name.padEnd(28)} ${String(e.rawWeeklyHours || '—').padStart(6)}  (${String(e.totalHours).padStart(5)}h)  ${bar}`);
  });

  // Simulate matching two invoice senders
  console.log('\nSample lookups:');
  ['Aaron Norris', 'Jack Henderson', 'Unknown Person'].forEach(name => {
    const hours = lookupHours(name, result.byName);
    console.log(`  "${name}" → ${hours !== null ? hours + 'h' : 'NOT FOUND'}`);
  });

} catch (err) {
  console.error('Failed:', err.message);
  process.exit(1);
}
