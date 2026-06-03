/**
 * Test script to validate Connecteam API integration.
 * Run with: node scripts/test-connecteam.js
 */

const { config } = require('../src/config');
const { ConnecteamClient } = require('../src/connecteam/client');

async function test() {
  console.log('Testing Connecteam API integration...\n');

  const client = new ConnecteamClient(
    config.connecteam.apiKey,
    config.connecteam.baseUrl,
    config.connecteam.mainTimeClockId
  );

  // Test 1: Validate Aaron Norris invoice for week of Apr 27 – May 3
  // Connecteam UI shows: Thu 30/4 (10:00h) + Fri 1/5 (09:14h) = 19:14 total
  console.log('Test 1: Validating Aaron Norris invoice (week ending 2026-05-01)...');
  const result = await client.validateInvoiceHours('Aaron Norris', '2026-05-01');
  console.log('\nResult:');
  console.log(`  Name:        ${result.name}`);
  console.log(`  User ID:     ${result.userId}`);
  console.log(`  Total hours: ${result.totalHours}`);
  console.log(`  Note:        ${result.note}`);
  if (result.shifts && result.shifts.length > 0) {
    console.log(`  Shifts:`);
    result.shifts.forEach(s => {
      console.log(`    ${s.date}  ${s.startTime} – ${s.endTime}  (${s.durationHours}h)  [${s.status}]`);
    });
  }

  console.log('\nTest complete.');
}

(async () => {
  try {
    await test();
  } catch (err) {
    console.error('Test failed:', err.message);
    process.exit(1);
  }
})();
