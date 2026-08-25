/* ═══════════════════════════════════════════════════════════
   STEM PoC Meetings — shared data loader
   Comeback Patch 2

   ONE SOURCE OF TRUTH: /data/stem-poc-meetings.json
   Every page that needs "what's the next STEM PoC meeting"
   (profile.html, index.html, workgroups.html) loads through this
   file instead of hardcoding a date. Edit the JSON, not the HTML.

   Local development:
     cd "/Users/greg_macer/Projects/04-PA-Statewide/STEM_POC_Tool"
     python3 -m http.server 8000
     open http://localhost:8000
   (fetch() of the JSON file is blocked under file:// by the browser —
   this loader detects that and fails gracefully, see loadStemPocMeetings.)

   FUTURE:
   Current source: local JSON file (/data/stem-poc-meetings.json).
   Future source: SharePoint List / Power Automate / other API.
   Everything below only depends on the normalized meeting shape
   returned by normalizeMeeting() — swapping the fetch in
   loadStemPocMeetings() for a different data source should not
   require touching any rendering code on any page.
═══════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  var MEETINGS_URL = '/data/stem-poc-meetings.json';
  var VALID_STATUSES = ['scheduled', 'tentative', 'completed', 'cancelled'];
  var NEXT_MEETING_EXCLUDED_STATUSES = ['cancelled', 'completed'];

  // ── Calendar-safe date parsing ──────────────────────────────────
  // Parses "YYYY-MM-DD" as local calendar components (not `new Date(str)`,
  // which parses as UTC midnight and can silently shift a day depending on
  // the viewer's timezone).
  function parseCalendarDate(dateString) {
    if (typeof dateString !== 'string') return null;
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateString.trim());
    if (!m) return null;
    var year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
    var d = new Date(year, month - 1, day);
    // Reject JS's auto-rollover of invalid dates (e.g. 2026-02-30 -> Mar 2)
    if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return null;
    return d;
  }

  function startOfDay(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  // ── Validation (Part 22) — never let a malformed record crash a page ──
  function normalizeMeeting(raw) {
    if (!raw || typeof raw !== 'object') return null;
    if (!raw.id || typeof raw.id !== 'string') {
      console.warn('[stem-poc-meetings] Skipping meeting with no id.', raw);
      return null;
    }
    var dateObj = parseCalendarDate(raw.date);
    if (!dateObj) {
      console.warn('[stem-poc-meetings] Skipping meeting "' + raw.id + '" — invalid or missing date.');
      return null;
    }
    var status = VALID_STATUSES.indexOf(raw.status) !== -1 ? raw.status : 'scheduled';
    var links = (raw.links && typeof raw.links === 'object') ? raw.links : {};

    return {
      id: raw.id,
      title: (typeof raw.title === 'string' && raw.title.trim()) ? raw.title.trim() : '',
      date: raw.date,
      dateObj: dateObj,
      startTime: typeof raw.startTime === 'string' ? raw.startTime.trim() : '',
      endTime: typeof raw.endTime === 'string' ? raw.endTime.trim() : '',
      location: typeof raw.location === 'string' ? raw.location.trim() : '',
      format: typeof raw.format === 'string' ? raw.format.trim() : '',
      status: status,
      summary: typeof raw.summary === 'string' ? raw.summary.trim() : '',
      priorities: Array.isArray(raw.priorities) ? raw.priorities.filter(isNonEmptyString) : [],
      preparation: Array.isArray(raw.preparation) ? raw.preparation.filter(isNonEmptyString) : [],
      agendaItems: Array.isArray(raw.agendaItems) ? raw.agendaItems.filter(isUsableAgendaItem) : [],
      links: {
        meeting: typeof links.meeting === 'string' ? links.meeting.trim() : '',
        agenda: typeof links.agenda === 'string' ? links.agenda.trim() : '',
        resources: typeof links.resources === 'string' ? links.resources.trim() : ''
      },
      notes: typeof raw.notes === 'string' ? raw.notes : ''
    };
  }

  function isNonEmptyString(x) { return typeof x === 'string' && x.trim().length > 0; }
  function isUsableAgendaItem(x) {
    if (typeof x === 'string') return x.trim().length > 0;
    return !!(x && typeof x === 'object' && (x.title || x.description));
  }

  // ── Loader ────────────────────────────────────────────────────
  async function loadStemPocMeetings() {
    if (window.location.protocol === 'file:') {
      console.warn('[stem-poc-meetings] Meeting data requires the local web server (fetch() is blocked under file://). Run: python3 -m http.server 8000');
      return [];
    }
    try {
      var url = MEETINGS_URL + '?v=' + Date.now(); // cache-bust (Part 17)
      var res = await fetch(url, { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      var json = await res.json();
      var rawMeetings = Array.isArray(json.meetings) ? json.meetings : [];
      return rawMeetings.map(normalizeMeeting).filter(Boolean);
    } catch (e) {
      console.warn('[stem-poc-meetings] Could not load meeting data — the rest of the site continues normally.', e.message || e);
      return [];
    }
  }

  // ── Selection (Part 4/5/20/21) ───────────────────────────────────
  function getNextStemPocMeeting(meetings, now) {
    now = now instanceof Date ? now : new Date();
    var today = startOfDay(now);
    var eligible = (meetings || []).filter(function (m) {
      return NEXT_MEETING_EXCLUDED_STATUSES.indexOf(m.status) === -1 && m.dateObj.getTime() >= today.getTime();
    });
    eligible.sort(function (a, b) { return a.dateObj - b.dateObj; }); // chronological, not JSON order
    var next = eligible[0];
    if (!next) return null;
    var result = Object.assign({}, next);
    result.isToday = next.dateObj.getTime() === today.getTime();
    return result;
  }

  function getRecentStemPocMeeting(meetings, now) {
    now = now instanceof Date ? now : new Date();
    var today = startOfDay(now);
    var past = (meetings || []).filter(function (m) {
      return m.status !== 'cancelled' && m.dateObj.getTime() < today.getTime();
    });
    past.sort(function (a, b) { return b.dateObj - a.dateObj; }); // most recent first
    return past[0] || null;
  }

  // ── Formatting ────────────────────────────────────────────────
  function formatStemPocMeetingDate(dateString) {
    var d = parseCalendarDate(dateString);
    if (!d) return '';
    return d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  }

  function formatStemPocMeetingTime(startTime, endTime) {
    var s = (startTime || '').trim();
    var e = (endTime || '').trim();
    if (s && e) return s + ' – ' + e;
    return s;
  }

  window.STEM_POC_MEETINGS = {
    load: loadStemPocMeetings,
    getNext: getNextStemPocMeeting,
    getRecent: getRecentStemPocMeeting,
    formatDate: formatStemPocMeetingDate,
    formatTime: formatStemPocMeetingTime
  };
})();
