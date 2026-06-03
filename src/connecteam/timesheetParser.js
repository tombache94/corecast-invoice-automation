/**
 * Parse a Connecteam weekly timesheet export (.xlsx) into a map of
 * { "First Last" => totalHours (decimal) }
 *
 * Export format (from "All Employees" sheet):
 *   - One header row, then one or more rows per employee
 *   - First name + Last name only appear on the employee's FIRST row
 *   - "Weekly total hours" (col O) only appears on the employee's first row
 *   - Hours are strings like "19:14" (HH:MM) — converted to decimal (19.233...)
 *   - Employees with no shifts have no time value in the weekly column
 */

const XLSX = require('xlsx');

/**
 * Convert a "HH:MM" string to decimal hours.
 * Returns 0 if the value is missing or unparseable.
 */
function hhmmToDecimal(value) {
  if (!value || typeof value !== 'string') return 0;
  const parts = value.trim().split(':');
  if (parts.length < 2) return 0;
  const hours = parseInt(parts[0], 10) || 0;
  const mins  = parseInt(parts[1], 10) || 0;
  return Math.round((hours + mins / 60) * 100) / 100;
}

/**
 * Parse the Connecteam timesheet XLSX export.
 *
 * @param {string} filePath - Absolute path to the .xlsx file
 * @returns {object} - {
 *     weekStart: "2026-04-27",
 *     weekEnd:   "2026-05-03",
 *     employees: [
 *       { name: "Aaron Norris", firstName: "Aaron", lastName: "Norris", totalHours: 19.23 },
 *       ...
 *     ],
 *     byName: { "aaron norris": 19.23, ... }   // lowercase keys for matching
 *   }
 */
function parseTimesheetExport(filePath) {
  const wb = XLSX.readFile(filePath);

  // Use "All Employees" sheet if present, otherwise use the first sheet
  const sheetName = wb.SheetNames.includes('All Employees')
    ? 'All Employees'
    : wb.SheetNames[0];

  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rows.length < 2) {
    throw new Error(`Timesheet "${sheetName}" sheet has no data rows`);
  }

  // Map column names to indices from header row
  const header = rows[0].map(h => (h || '').toString().trim());
  const col = name => header.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const COL_FIRST  = col('First name');
  const COL_LAST   = col('Last name');
  const COL_WEEKLY = col('Weekly total hours');

  if (COL_FIRST === -1 || COL_LAST === -1) {
    throw new Error('Could not find "First name" / "Last name" columns in timesheet export');
  }
  if (COL_WEEKLY === -1) {
    throw new Error('Could not find "Weekly total hours" column in timesheet export');
  }

  // Extract week range from filename e.g. timeclock-timesheet_overview_2026-04-27_2026-05-03.xlsx
  const fileBase = filePath.replace(/\\/g, '/').split('/').pop();
  const dateMatch = fileBase.match(/(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})/);
  const weekStart = dateMatch ? dateMatch[1] : null;
  const weekEnd   = dateMatch ? dateMatch[2] : null;

  const employees = [];
  const byName = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const firstName = (row[COL_FIRST] || '').toString().trim();
    const lastName  = (row[COL_LAST]  || '').toString().trim();

    // Only process rows where a first name is present (first row for each employee)
    if (!firstName) continue;

    const weeklyRaw  = (row[COL_WEEKLY] || '').toString().trim();
    const totalHours = hhmmToDecimal(weeklyRaw);
    const name       = `${firstName} ${lastName}`.trim();

    const entry = {
      name,
      firstName,
      lastName,
      totalHours,
      rawWeeklyHours: weeklyRaw || null,
    };

    employees.push(entry);
    byName[name.toLowerCase()] = totalHours;
  }

  return { weekStart, weekEnd, employees, byName };
}

/**
 * Match an invoice sender name against the timesheet byName map.
 * Tolerant: handles case, partial matches, and last-name-only invoices.
 *
 * @param {string} senderName  - Name from the invoice (e.g. "Aaron Norris")
 * @param {object} byName      - { "aaron norris": 19.23, ... }
 * @returns {number|null}      - Hours from timesheet, or null if not found
 */
function lookupHours(senderName, byName) {
  const needle = senderName.toLowerCase().trim();

  // 1. Exact match
  if (byName[needle] !== undefined) return byName[needle];

  // 2. Substring match either way
  for (const [key, hours] of Object.entries(byName)) {
    if (key.includes(needle) || needle.includes(key)) return hours;
  }

  // 3. Last-name only match (e.g. invoice says "Norris")
  const lastWord = needle.split(/\s+/).pop();
  for (const [key, hours] of Object.entries(byName)) {
    if (key.split(/\s+/).pop() === lastWord) return hours;
  }

  return null;
}

module.exports = { parseTimesheetExport, lookupHours, hhmmToDecimal };
