/**
 * Connecteam API client for fetching timesheet data and validating invoices.
 * Uses the API key from environment variables.
 *
 * Key facts about the Connecteam API:
 * - Time clock API ID (16246267) differs from the punch-clock UI ID (13723871)
 * - The userId query param on time-activities returns the user's entry with empty shifts
 *   (it appears to be ignored or broken server-side) — so we fetch all users and filter client-side
 * - Rate limit is fairly tight; minimise calls by caching users + time-activities per run
 */

const fetch = require('node-fetch');

class ConnecteamClient {
  constructor(apiKey, baseUrl = 'https://api.connecteam.com', mainTimeClockId = '16246267') {
    this.apiKey = apiKey;
    this.baseUrl = baseUrl;
    this.mainTimeClockId = mainTimeClockId;
    // Per-instance caches to avoid redundant API calls within one run
    this._usersCache = null;
    this._timeClockIdCache = null;
  }

  /**
   * Make an authenticated request to the Connecteam API.
   */
  async request(endpoint, method = 'GET', body = null) {
    const url = `${this.baseUrl}${endpoint}`;
    const options = {
      method,
      headers: {
        'X-API-KEY': this.apiKey,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      },
    };

    if (body) options.body = JSON.stringify(body);

    const res = await fetch(url, options);
    const data = await res.json();

    if (!res.ok) {
      console.error(`  [API ERROR] ${res.status} ${url}:`, JSON.stringify(data));
      throw new Error(`Connecteam API error: ${res.status} ${data.detail || data.message || data.error || JSON.stringify(data)}`);
    }

    return data;
  }

  /**
   * Return all users in the account (cached for the lifetime of this client instance).
   */
  async getUsers() {
    if (this._usersCache) return this._usersCache;
    const data = await this.request('/users/v1/users');
    this._usersCache = data.data?.users || data.users || [];
    return this._usersCache;
  }

  /**
   * Search for a user by name or email.
   */
  async searchUsers(query) {
    try {
      const allUsers = await this.getUsers();
      const q = query.toLowerCase();
      return allUsers.filter(user => {
        const full = `${user.firstName || ''} ${user.lastName || ''}`.toLowerCase();
        return (
          full.includes(q) ||
          (user.email || '').toLowerCase().includes(q)
        );
      });
    } catch (err) {
      console.error(`Error searching users for "${query}":`, err.message);
      return [];
    }
  }

  /**
   * Resolve the correct time clock ID to use for the time-activities API.
   * The UI punch-clock ID (13723871) is different from the API time clock ID.
   * Cached after first call.
   */
  async resolveTimeClockId() {
    if (this._timeClockIdCache) return this._timeClockIdCache;

    const data = await this.request('/time-clock/v1/time-clocks');
    const clocks = data.data?.timeClocks || data.timeClocks || [];
    console.log(`  [DEBUG] Time clocks in account:`,
      clocks.map(c => `${c.id} "${c.name}" (archived: ${c.isArchived})`).join(' | '));

    // Prefer configured ID if it exists in the list
    if (this.mainTimeClockId) {
      const match = clocks.find(c => String(c.id) === String(this.mainTimeClockId));
      if (match) {
        this._timeClockIdCache = match.id;
        return this._timeClockIdCache;
      }
      console.warn(`  [WARN] mainTimeClockId ${this.mainTimeClockId} not found — using first active clock`);
    }

    const active = clocks.find(c => !c.isArchived) || clocks[0];
    if (!active) throw new Error('No time clocks found in Connecteam account');
    this._timeClockIdCache = active.id;
    return this._timeClockIdCache;
  }

  /**
   * Fetch all time activities for a date range, then extract shifts for a specific user.
   *
   * NOTE: The userId query param appears to be ignored by the API (returns empty shifts).
   * We fetch the full response and filter client-side instead.
   *
   * @param {number|string} userId - Connecteam user ID
   * @param {string} startDate - YYYY-MM-DD
   * @param {string} endDate   - YYYY-MM-DD
   * @returns {object[]} - Array of normalised shift records: { date, startMs, endMs, durationHours, status }
   */
  async getShiftsForUser(userId, startDate, endDate) {
    const timeClockId = await this.resolveTimeClockId();

    // Fetch WITHOUT userId filter — the param returns empty shifts server-side
    const endpoint = `/time-clock/v1/time-clocks/${timeClockId}/time-activities?startDate=${startDate}&endDate=${endDate}`;
    const data = await this.request(endpoint);

    const allUserActivities =
      data.data?.timeActivities ||
      data.timeActivities ||
      (Array.isArray(data.data) ? data.data : []);

    // Find this user's entry in the response
    const userEntry = Array.isArray(allUserActivities)
      ? allUserActivities.find(a => String(a.userId) === String(userId))
      : null;

    if (!userEntry) {
      console.log(`  [DEBUG] userId ${userId} not found in time-activities response. ` +
        `Users present: ${allUserActivities.map(a => a.userId).join(', ')}`);
      return [];
    }

    console.log(`  [DEBUG] Found ${(userEntry.shifts || []).length} shift(s) for userId ${userId}`);

    // Normalise each shift into a simple record
    return (userEntry.shifts || []).map(shift => {
      // Timestamps can be epoch seconds or milliseconds — normalise to ms
      const toMs = t => t > 1e12 ? t : t * 1000;
      const startMs = shift.start?.timestamp ? toMs(shift.start.timestamp) : null;
      const endMs   = shift.end?.timestamp   ? toMs(shift.end.timestamp)   : null;
      const durationHours = (startMs && endMs) ? (endMs - startMs) / 3_600_000 : 0;

      return {
        date: new Date(startMs).toISOString().split('T')[0],
        startMs,
        endMs,
        startTime: shift.start?.timestamp ? new Date(startMs).toLocaleTimeString('en-AU', { hour: '2-digit', minute: '2-digit', hour12: false }) : null,
        endTime:   shift.end?.timestamp   ? new Date(endMs).toLocaleTimeString('en-AU', { hour: '2-digit', minute: '2-digit', hour12: false }) : null,
        durationHours: Math.round(durationHours * 100) / 100,
        status: shift.status || 'unknown',
        raw: shift,
      };
    }).filter(s => s.durationHours > 0);
  }

  /**
   * Validate a subcontractor's invoice hours against their Connecteam time clock.
   *
   * @param {string} senderName - Name from the invoice (e.g. "Aaron Norris")
   * @param {string} weekEnding - The invoice's week-ending date, ISO format (e.g. "2026-05-01")
   * @returns {object} - { userId, name, totalHours, shifts[], note }
   */
  async validateInvoiceHours(senderName, weekEnding) {
    // Derive the Monday–Sunday week that contains the weekEnding date
    const invoiceDate = new Date(weekEnding);
    const dow = invoiceDate.getDay(); // 0=Sun … 6=Sat
    const daysBackToMonday = dow === 0 ? 6 : dow - 1;

    const monday = new Date(invoiceDate);
    monday.setDate(invoiceDate.getDate() - daysBackToMonday);

    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    const startDate = monday.toISOString().split('T')[0];
    const endDate   = sunday.toISOString().split('T')[0];

    console.log(`  Validating "${senderName}" hours for week ${startDate} → ${endDate}...`);

    // Step 1: Find user (uses cached users list — no extra API call if already fetched)
    const matches = await this.searchUsers(senderName);
    if (matches.length === 0) {
      return {
        userId: null,
        name: senderName,
        totalHours: null,
        shifts: [],
        note: `User "${senderName}" not found in Connecteam`,
      };
    }

    const user = matches[0];
    const userId = user.id || user.userId;
    const displayName = `${user.firstName || ''} ${user.lastName || ''}`.trim() || senderName;

    // Step 2: Fetch shifts (single API call, filtered client-side)
    const shifts = await this.getShiftsForUser(userId, startDate, endDate);

    // Step 3: Sum hours
    const totalHours = Math.round(shifts.reduce((sum, s) => sum + s.durationHours, 0) * 100) / 100;
    const pendingCount  = shifts.filter(s => s.status !== 'approved').length;
    const approvedCount = shifts.filter(s => s.status === 'approved').length;

    const shiftSummary = shifts.map(s =>
      `  ${s.date} ${s.startTime}–${s.endTime} = ${s.durationHours}h [${s.status}]`
    ).join('\n');

    if (shifts.length > 0) {
      console.log(`  Shifts found:\n${shiftSummary}`);
    }

    return {
      userId,
      name: displayName,
      totalHours: totalHours > 0 ? totalHours : null,
      shifts,
      note: shifts.length > 0
        ? `${shifts.length} shift(s), ${totalHours}h total (${approvedCount} approved, ${pendingCount} pending)`
        : 'No shifts found in Connecteam for this week',
    };
  }
}

module.exports = { ConnecteamClient };
