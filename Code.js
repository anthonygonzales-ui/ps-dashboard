// ============================================================
// PS Lifecycle Dashboard — Google Apps Script Backend
// Spreadsheet: PS Lifecycle Dashboard
// Spreadsheet ID: 1XVgzOgoQvEN1r1lRJemBVJx5xGB6FrruiKfiT7hhqWs
// ============================================================

const SPREADSHEET_ID = '1XVgzOgoQvEN1r1lRJemBVJx5xGB6FrruiKfiT7hhqWs';

/**
 * Serves the dashboard when accessed via the web app URL.
 */
function doGet(e) {
  const output = HtmlService.createHtmlOutputFromFile('index');
  output.setTitle('PS Lifecycle Dashboard');
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

/**
 * Returns all sheet data as a JSON string.
 * Called from the frontend via google.script.run.getAllData()
 */
function getAllData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const result = {};
    const tz = Session.getScriptTimeZone();

    sheets.forEach(function (sheet) {
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 1 || lastCol < 1) return;

      const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

      // Hours Data is large and lazy-loaded — exclude from initial payload
      const fRow = (values[0] || []).map(String);
      if (fRow.indexOf('Work: Owner Name') > -1 || fRow.indexOf('Hours (Number)') > -1) return;

      result[name] = values.map(function (row) {
        return row.map(function (cell) {
          if (cell instanceof Date) {
            if (cell.getFullYear() < 1970) return '';
            return Utilities.formatDate(cell, tz, 'M/d/yyyy');
          }
          if (typeof cell === 'number') return cell;
          return String(cell);
        });
      });
    });

    return JSON.stringify({
      success: true,
      data: result,
      timestamp: new Date().toISOString()
    });

  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Returns just the Hours Data sheet rows (used by lazy-load on Utilization tab).
 * Scans sheets for the one containing 'Work: Owner Name' or 'Hours (Number)'.
 */
function getHoursData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var tz = Session.getScriptTimeZone();
    for (var si = 0; si < sheets.length; si++) {
      var sheet = sheets[si];
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) continue;
      // Read just the header row to identify the hours sheet cheaply
      var hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      if (hdrs.indexOf('Work: Owner Name') < 0 && hdrs.indexOf('Hours (Number)') < 0) continue;
      // Found it — read full data
      var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var rows = values.map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date) {
            if (cell.getFullYear() < 1970) return '';
            return Utilities.formatDate(cell, tz, 'M/d/yyyy');
          }
          if (typeof cell === 'number') return cell;
          return String(cell);
        });
      });
      return JSON.stringify({ success: true, rows: rows });
    }
    return JSON.stringify({ success: false, error: 'Hours Data sheet not found' });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Saves a single Bookings Forecast entry (key → value) to the
 * "Bookings Forecast Data" sheet.  Creates the sheet on first call.
 */
function saveBooking(key, value) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Bookings Forecast Data');

    if (!sheet) {
      sheet = ss.insertSheet('Bookings Forecast Data');
      sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
          sheet.getRange(i + 2, 2).setValue(value);
          return;
        }
      }
    }

    sheet.appendRow([key, value]);

  } catch (e) {
    Logger.log('saveBooking error: ' + e.toString());
  }
}

/**
 * Saves a single Backlog Analysis entry (key → value) to the
 * "Backlog Analysis Data" sheet.
 */
function saveBacklogEntry(key, value) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Backlog Analysis Data');

    if (!sheet) {
      sheet = ss.insertSheet('Backlog Analysis Data');
      sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) {
          sheet.getRange(i + 2, 2).setValue(value);
          return;
        }
      }
    }

    sheet.appendRow([key, value]);

  } catch (e) {
    Logger.log('saveBacklogEntry error: ' + e.toString());
  }
}


// ════════════════════════════════════════════════════════════
// TEAM AVAILABILITY — CALENDAR API
//
// Parallelization strategy:
//   The frontend calls getUserList() once to get the user roster +
//   week metadata, then fires N simultaneous getCalendarDataForBatch()
//   calls (one per batch of ~10 users).  Each call is an independent
//   Apps Script execution that runs in parallel, reducing wall-clock
//   time from ~40 s to ~10 s for a 40-person team.
// ════════════════════════════════════════════════════════════

// ── Shared constants ─────────────────────────────────────────
var CAL_OOO_KW    = ['ooo','out of office','pto','vacation','holiday','time off',
                     'leave','personal day','off work','annual leave','sick'];
var CAL_DAY_NAMES = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
var CAL_DAY_SHORT = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// ── OOO helpers ──────────────────────────────────────────────

function calTitleIsOOO(title) {
  if (!title) return false;
  var t = title.toLowerCase();
  return CAL_OOO_KW.some(function(kw) { return t.indexOf(kw) > -1; });
}

function calEventIsOOO(ev) {
  if (calTitleIsOOO(ev.getTitle() || '')) return true;
  if (ev.isAllDayEvent()) {
    var dur   = ev.getEndTime() - ev.getStartTime();
    var title = ev.getTitle() || '';
    if (dur >= 2 * 86400000 && (title === '' || title === 'Busy')) return true;
  }
  return false;
}

function calWorkingDaysInRange(evStart, evEnd, wkStart, wkEnd) {
  var s = (evStart < wkStart) ? wkStart : evStart;
  var e = (evEnd   > wkEnd)   ? wkEnd   : evEnd;
  if (s >= e) return 0;
  var days = 0;
  var d = new Date(s); d.setHours(0,0,0,0);
  while (d < e) {
    var wd = d.getDay();
    if (wd !== 0 && wd !== 6) days++;
    d.setDate(d.getDate() + 1);
  }
  return days;
}

function calMergeIntervalsHrs(intervals) {
  if (!intervals.length) return 0;
  intervals.sort(function(a, b) { return a[0] - b[0]; });
  var merged = [intervals[0].slice()];
  for (var i = 1; i < intervals.length; i++) {
    var last = merged[merged.length - 1];
    if (intervals[i][0] < last[1]) {
      last[1] = Math.max(last[1], intervals[i][1]);
    } else {
      merged.push(intervals[i].slice());
    }
  }
  var ms = 0;
  merged.forEach(function(iv) { ms += iv[1] - iv[0]; });
  return ms / 3600000;
}

function calBuildOOONote(ev, wkStart, wkEnd) {
  var title = ev.getTitle() || '';
  var label = (title && title !== 'Busy') ? title : 'OOO';
  if (ev.isAllDayEvent()) {
    var s = new Date(ev.getStartTime()); s.setHours(0,0,0,0);
    var e = new Date(ev.getEndTime());   e.setDate(e.getDate() - 1); e.setHours(0,0,0,0);
    if (s.toDateString() === e.toDateString()) {
      return label + ' on ' + CAL_DAY_NAMES[s.getDay()];
    }
    var affected = [];
    var d   = new Date(Math.max(s.getTime(), wkStart.getTime())); d.setHours(0,0,0,0);
    var cap = new Date(Math.min(e.getTime(), wkEnd.getTime()));   cap.setHours(0,0,0,0);
    while (d <= cap) {
      var wd = d.getDay();
      if (wd >= 1 && wd <= 5) affected.push(CAL_DAY_SHORT[wd]);
      d.setDate(d.getDate() + 1);
    }
    if (affected.length === 0) return label;
    if (affected.length >= 5) return label + ' (full week)';
    if (affected.length === 1) return label + ' on ' + affected[0];
    return label + ' (' + affected.join(', ') + ')';
  } else {
    return label + ' on ' + CAL_DAY_NAMES[ev.getStartTime().getDay()];
  }
}

// Returns true if the calendar owner has DECLINED this event.
// Guest list is PRIMARY check; getMyStatus() is FALLBACK only.
// Critical: declined status is CalendarApp.GuestStatus.NO — not DECLINED (which is undefined).
function calendarOwnerDeclined(ev, ownerEmail) {
  try {
    var lower     = ownerEmail.toLowerCase();
    var localPart = lower.indexOf('@') > -1 ? lower.split('@')[0] : lower;
    var guests = ev.getGuestList(true);
    for (var gi = 0; gi < guests.length; gi++) {
      var gEmail = String(guests[gi].getEmail() || '').toLowerCase();
      var gLocal = gEmail.indexOf('@') > -1 ? gEmail.split('@')[0] : gEmail;
      if (gEmail === lower || (localPart && gLocal === localPart)) {
        var gStatus = guests[gi].getGuestStatus();
        if (gStatus === CalendarApp.GuestStatus.OWNER) return false;
        return gStatus === CalendarApp.GuestStatus.NO; // NO = declined in Apps Script
      }
    }
    return ev.getMyStatus() === CalendarApp.GuestStatus.NO;
  } catch(err) { return false; }
}

// Returns true if the event has at least one attendee outside @redis.com.
function hasExternalAttendee(ev) {
  try {
    var guests = ev.getGuestList();
    if (!guests || guests.length === 0) return false;
    for (var gi = 0; gi < guests.length; gi++) {
      var gEmail = String(guests[gi].getEmail() || '').toLowerCase();
      if (gEmail && gEmail.indexOf('@redis.com') < 0) return true;
    }
    return false;
  } catch(err) { return false; }
}

// Analyzes one week of events for a single calendar owner.
// userTz — IANA timezone string for the person (e.g. "America/New_York").
// Only time within their local 9am–5pm window is counted; events entirely
// outside that window are skipped; those that straddle a boundary are clipped.
function analyzeWeek(wkStart, wkEnd, calObj, ownerEmail, userTz) {
  var workTz = userTz || Session.getScriptTimeZone();

  // Returns the UTC ms for local midnight on the calendar day that contains `ms`.
  // Uses Utilities.formatDate so DST transitions are handled automatically.
  function toLocalMidnightMs(ms) {
    var d         = new Date(ms);
    var dayStr    = Utilities.formatDate(d, workTz, 'yyyy-MM-dd');   // "2024-07-15"
    var offsetStr = Utilities.formatDate(d, workTz, 'Z');             // "+0530" or "-0700"
    var sign      = offsetStr.charAt(0) === '+' ? 1 : -1;
    var offHr     = parseInt(offsetStr.slice(1, 3), 10);
    var offMin    = parseInt(offsetStr.slice(3, 5), 10);
    var parts     = dayStr.split('-');
    var utcMidnight = Date.UTC(parseInt(parts[0], 10),
                               parseInt(parts[1], 10) - 1,
                               parseInt(parts[2], 10));
    // local midnight in UTC = UTC midnight of that date MINUS the UTC offset
    return utcMidnight - sign * (offHr * 60 + offMin) * 60000;
  }

  // Clips [startMs, endMs] to the 9am–5pm window of the event's local day.
  // Returns null if the event falls entirely outside working hours.
  function clipToWorkday(startMs, endMs) {
    var midnight  = toLocalMidnightMs(startMs);
    var workStart = midnight + 9  * 3600000;   // 09:00 local
    var workEnd   = midnight + 17 * 3600000;   // 17:00 local
    var cs = Math.max(startMs, workStart);
    var ce = Math.min(endMs,   workEnd);
    return cs < ce ? [cs, ce] : null;
  }

  var events        = calObj.getEvents(wkStart, wkEnd);
  var allIntervals  = [];
  var custIntervals = [];
  var intIntervals  = [];
  var oooHrs        = 0;
  var oooNotes      = [];
  var meetingsList  = [];

  events.forEach(function(ev) {
    if (calendarOwnerDeclined(ev, ownerEmail)) return;

    var ooo = calEventIsOOO(ev);
    if (ev.isAllDayEvent()) {
      // All-day OOO → 8h per affected working day (timezone-agnostic; unchanged)
      var hrs = calWorkingDaysInRange(ev.getStartTime(), ev.getEndTime(), wkStart, wkEnd) * 8;
      if (ooo) { oooHrs += hrs; oooNotes.push(calBuildOOONote(ev, wkStart, wkEnd)); }
    } else {
      var evStartMs = ev.getStartTime().getTime();
      var evEndMs   = ev.getEndTime().getTime();

      if (ooo) {
        // Timed OOO — clip to 9–5 before counting
        var clipped = clipToWorkday(evStartMs, evEndMs);
        if (clipped) {
          oooHrs += (clipped[1] - clipped[0]) / 3600000;
          oooNotes.push(calBuildOOONote(ev, wkStart, wkEnd));
        }
      } else {
        // Regular meeting — clip to 9–5; skip if entirely outside
        var iv = clipToWorkday(evStartMs, evEndMs);
        if (!iv) return;
        var isCust = hasExternalAttendee(ev);
        allIntervals.push(iv);
        if (isCust) { custIntervals.push(iv); }
        else        { intIntervals.push(iv); }
        meetingsList.push({
          title:   ev.getTitle() || '(no title)',
          startMs: evStartMs,                   // original time — for display in drill-down
          endMs:   evEndMs,
          rawHrs:  Math.round((iv[1] - iv[0]) / 360000) / 10,  // clipped hours
          type:    isCust ? 'cust' : 'int'
        });
      }
    }
  });

  meetingsList.sort(function(a, b) { return a.startMs - b.startMs; });

  var plannedHrs  = calMergeIntervalsHrs(allIntervals);
  var customerHrs = calMergeIntervalsHrs(custIntervals);
  var internalHrs = calMergeIntervalsHrs(intIntervals);
  oooHrs = Math.min(oooHrs, 40);
  var cap      = Math.max(0, 40 - oooHrs);
  var custUtil = (cap > 0) ? (customerHrs / cap) : null;
  var totUtil  = (cap > 0) ? (plannedHrs  / cap) : null;

  return {
    customerHrs:  Math.round(customerHrs * 10) / 10,
    internalHrs:  Math.round(internalHrs * 10) / 10,
    plannedHrs:   Math.round(plannedHrs  * 10) / 10,
    oooHrs:       Math.round(oooHrs      * 10) / 10,
    capacity:     Math.round(cap         * 10) / 10,
    customerUtil: custUtil !== null ? Math.round(custUtil * 1000) / 10 : null,
    utilization:  totUtil  !== null ? Math.round(totUtil  * 1000) / 10 : null,
    oooNotes:     oooNotes,
    meetings:     meetingsList
  };
}

// ── Read PS User Records sheet ────────────────────────────────

function readPSUsers() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var userRows = null;
  for (var si = 0; si < sheets.length; si++) {
    var shData = sheets[si].getDataRange().getValues();
    var hdrs0  = shData[0] || [];
    if ((hdrs0.indexOf('Full Name') > -1 || hdrs0.indexOf('Name') > -1) &&
         hdrs0.indexOf('Title') > -1) {
      userRows = shData;
      break;
    }
  }

  var users = [];
  if (userRows && userRows.length > 1) {
    var hdrs     = userRows[0];
    var iName    = hdrs.indexOf('Full Name'); if (iName    < 0) iName    = hdrs.indexOf('Name');
    var iRegion  = hdrs.indexOf('Region');
    var iManager = hdrs.indexOf('Manager Name'); if (iManager < 0) iManager = hdrs.indexOf('Manager');
    var EMAIL_COLS = ['Email','Email Address','Work Email','Google Email',
                      'Google Account','Username','User Email','Login'];
    var iEmail = -1;
    for (var ec = 0; ec < EMAIL_COLS.length && iEmail < 0; ec++) {
      iEmail = hdrs.indexOf(EMAIL_COLS[ec]);
    }
    if (iEmail < 0) {
      for (var hc = 0; hc < hdrs.length && iEmail < 0; hc++) {
        var h = String(hdrs[hc]).toLowerCase();
        if (h.indexOf('email') > -1 || h.indexOf('gmail') > -1 || h === 'username') {
          iEmail = hc;
        }
      }
    }

    for (var ri = 1; ri < userRows.length; ri++) {
      var row    = userRows[ri];
      var name   = iName   >= 0 ? String(row[iName]   || '').trim() : '';
      if (!name) continue;
      var region = iRegion >= 0 ? String(row[iRegion] || '').trim() : '';
      if (!region) region = 'Unknown';
      var email  = iEmail  >= 0 ? String(row[iEmail]  || '').trim() : '';
      if (!email) {
        var parts = name.toLowerCase().replace(/[',]/g, '').split(/\s+/);
        var SFXS = ['jr', 'jr.', 'sr', 'sr.', 'ii', 'iii', 'iv'];
        while (parts.length > 1 && SFXS.indexOf(parts[parts.length - 1]) > -1) parts.pop();
        if (parts.length >= 2) {
          email = parts[0] + '.' + parts[parts.length - 1] + '@redis.com';
        }
      }
      if (email) {
        var manager = iManager >= 0 ? String(row[iManager] || '').trim() : '';
        users.push({ name: name, email: email, region: region, manager: manager });
      }
    }
  }

  if (users.length === 0) {
    users = [{ name: 'Anil Kondapaneni', email: 'anil.kondapaneni@redis.com',
               region: 'Global', manager: '' }];
  }

  return users;
}

// ── Compute week windows ──────────────────────────────────────

function getCalWeekWindows() {
  var now    = new Date();
  var dow    = now.getDay();
  var dToMon = (dow === 0) ? -6 : 1 - dow;
  var thisMon = new Date(now); thisMon.setDate(now.getDate() + dToMon); thisMon.setHours(0,0,0,0);
  var thisFri = new Date(thisMon); thisFri.setDate(thisMon.getDate() + 4); thisFri.setHours(23,59,59,0);
  var nextMon = new Date(thisMon); nextMon.setDate(thisMon.getDate() + 7);
  var nextFri = new Date(nextMon); nextFri.setDate(nextMon.getDate() + 4); nextFri.setHours(23,59,59,0);
  return { thisMon: thisMon, thisFri: thisFri, nextMon: nextMon, nextFri: nextFri };
}


// ════════════════════════════════════════════════════════════
// PUBLIC API — Team Availability (parallel)
// ════════════════════════════════════════════════════════════

/**
 * Step 1 of 2 in the parallel calendar load.
 * Returns the PS user roster + week metadata so the frontend can
 * split users into batches and fire them all simultaneously.
 * No Calendar API calls — fast (~0.5 s).
 */
function getUserList() {
  try {
    var weeks = getCalWeekWindows();
    var users = readPSUsers();

    function fmt(d) { return Utilities.formatDate(d, 'America/New_York', 'MMM d'); }

    return JSON.stringify({
      success:       true,
      users:         users,
      thisMondayMs:  weeks.thisMon.getTime(),
      nextMondayMs:  weeks.nextMon.getTime(),
      thisWeekLabel: fmt(weeks.thisMon) + ' – ' + fmt(weeks.thisFri),
      nextWeekLabel: fmt(weeks.nextMon) + ' – ' + fmt(weeks.nextFri)
    });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Step 2 of 2 in the parallel calendar load.
 * Called once per batch of users; multiple calls run simultaneously.
 *
 * @param {string} batchJson — JSON: { users, thisMondayMs, nextMondayMs }
 *   users: array of { name, email, region, manager }
 *   thisMondayMs / nextMondayMs: epoch ms for week start (from getUserList)
 * @returns {string} JSON: { success, people }
 *   people: array of person objects with thisWeek / nextWeek calendar metrics
 */
function getCalendarDataForBatch(batchJson) {
  try {
    var batch   = JSON.parse(batchJson);
    var users   = batch.users;
    var thisMon = new Date(batch.thisMondayMs); thisMon.setHours(0,0,0,0);
    var thisFri = new Date(thisMon); thisFri.setDate(thisMon.getDate() + 4); thisFri.setHours(23,59,59,0);
    var nextMon = new Date(batch.nextMondayMs); nextMon.setHours(0,0,0,0);
    var nextFri = new Date(nextMon); nextFri.setDate(nextMon.getDate() + 4); nextFri.setHours(23,59,59,0);

    var people = [];
    users.forEach(function(user) {
      var calObj = null;
      try { calObj = CalendarApp.getCalendarById(user.email); } catch(err) {}

      var personData = {
        name:       user.name,
        email:      user.email,
        region:     user.region  || 'Unknown',
        manager:    user.manager || '',
        accessible: !!calObj
      };

      if (calObj) {
        var userTz = '';
        try { userTz = calObj.getTimeZone(); } catch(e) {}
        if (!userTz) userTz = Session.getScriptTimeZone();
        personData.timezone = userTz;
        personData.thisWeek = analyzeWeek(thisMon, thisFri, calObj, user.email, userTz);
        personData.nextWeek = analyzeWeek(nextMon, nextFri, calObj, user.email, userTz);
      }

      people.push(personData);
    });

    return JSON.stringify({ success: true, people: people });

  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * Legacy single-call entry point — calls getUserList + getCalendarDataForBatch
 * internally.  Kept for backward compatibility; not used by the parallelized frontend.
 */
function getCalendarData() {
  try {
    var ulJson = getUserList();
    var ul     = JSON.parse(ulJson);
    if (!ul.success) return ulJson;

    var batchJson = JSON.stringify({
      users:        ul.users,
      thisMondayMs: ul.thisMondayMs,
      nextMondayMs: ul.nextMondayMs
    });

    var brJson = getCalendarDataForBatch(batchJson);
    var br     = JSON.parse(brJson);
    if (!br.success) return brJson;

    var REGION_ORDER = ['Global','AMER','EMEA','APAC'];
    var regionMap    = {};
    var totalAccessible = 0;

    br.people.forEach(function(person) {
      if (person.accessible) totalAccessible++;
      var reg = person.region || 'Unknown';
      if (!regionMap[reg]) regionMap[reg] = [];
      regionMap[reg].push(person);
    });

    var allKeys     = Object.keys(regionMap);
    var orderedKeys = REGION_ORDER.filter(function(r) { return !!regionMap[r]; });
    allKeys.forEach(function(r) { if (REGION_ORDER.indexOf(r) < 0) orderedKeys.push(r); });

    var regionsArr = orderedKeys.map(function(r) {
      return { name: r, people: regionMap[r] };
    });

    return JSON.stringify({
      success:         true,
      thisWeekLabel:   ul.thisWeekLabel,
      nextWeekLabel:   ul.nextWeekLabel,
      thisMondayMs:    ul.thisMondayMs,
      nextMondayMs:    ul.nextMondayMs,
      regions:         regionsArr,
      totalUsers:      ul.users.length,
      totalAccessible: totalAccessible,
      timestamp:       new Date().toISOString()
    });

  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


// ════════════════════════════════════════════════════════════
// DEBUG — run manually from the Apps Script editor
// ════════════════════════════════════════════════════════════

/**
 * Inspect why a declined meeting is still being counted for a given person.
 * 1. Set PERSON_EMAIL and DATE_TO_CHECK below.
 * 2. Click Run → debugDeclined in the editor.
 * 3. Open View → Logs.
 */
function debugDeclined() {
  var PERSON_EMAIL  = 'koumaran.bergen@redis.com';
  var DATE_TO_CHECK = new Date(2026, 5, 25); // Jun 25, 2026

  var cal = CalendarApp.getCalendarById(PERSON_EMAIL);
  if (!cal) {
    Logger.log('❌ Calendar not accessible for: ' + PERSON_EMAIL);
    return;
  }
  Logger.log('✅ Calendar found: ' + cal.getName() + ' (' + PERSON_EMAIL + ')');

  var dayStart = new Date(DATE_TO_CHECK); dayStart.setHours(0, 0, 0, 0);
  var dayEnd   = new Date(DATE_TO_CHECK); dayEnd.setHours(23, 59, 59, 0);
  var events   = cal.getEvents(dayStart, dayEnd);
  Logger.log('Events on ' + DATE_TO_CHECK.toDateString() + ': ' + events.length);

  events.forEach(function(ev) {
    Logger.log('\n──────────────────────────────────────');
    Logger.log('Title    : ' + ev.getTitle());
    Logger.log('Start    : ' + ev.getStartTime());
    Logger.log('End      : ' + ev.getEndTime());
    Logger.log('AllDay   : ' + ev.isAllDayEvent());

    try { Logger.log('getMyStatus() → ' + ev.getMyStatus()); }
    catch(e) { Logger.log('getMyStatus() → ERROR: ' + e); }

    try {
      var guestsWithOwner = ev.getGuestList(true);
      Logger.log('getGuestList(true) — ' + guestsWithOwner.length + ' entries:');
      guestsWithOwner.forEach(function(g) {
        Logger.log('  email="' + g.getEmail() + '"  status=' + g.getGuestStatus() +
                   (g.getName ? '  name="' + g.getName() + '"' : ''));
      });
    } catch(e) { Logger.log('getGuestList(true) → ERROR: ' + e); }

    Logger.log('→ FINAL: this event would be ' +
               (calendarOwnerDeclined(ev, PERSON_EMAIL) ? '🚫 EXCLUDED (declined)' : '✅ COUNTED'));
  });
}

// ============================================================
// GENERATE SLIDES DECK  (v2 — screenshots + redesigned slides)
// ============================================================
function generateSlidesDeck(payloadJson) {
  var _deckSection = 'parse';
  try {
    var p = JSON.parse(payloadJson);

    // ── Color constants (Redis brand — white theme) ─────────
    // Palette per Redis 2026 brand: Hyper red, Midnight, Dusk, Dusk 30%.
    var BG     = '#FFFFFF';   // White
    var CARD   = '#F4F6F7';   // subtle Dusk-tinted surface
    var CARD2  = '#E7ECEE';
    var RED    = '#FF4438';   // Hyper (Redis red)
    var TEXT   = '#091A23';   // Midnight
    var MUTED  = '#163341';   // Dusk
    var DIM    = '#B9C2C6';   // Dusk 30%
    var GREEN  = '#107C41';   // functional status (good) — kept for legibility
    var YELLOW = '#C27C0E';   // functional status (warn)
    var BLUE   = '#1565C0';   // functional status (info)
    // Paste your base64-encoded Redis logo PNG here (no data: prefix, just the raw base64 string).
    // Leave empty to render footers without the logo image.
    var REDIS_LOGO_B64 = '';


    // ── Slide geometry (points) ──────────────────────────────
    var W = 720, H = 405, PAD = 28;
    var CW = W - PAD * 2; // 664

    // ── Helper: set slide background (white) ────────────────
    function setBg(slide) {
      slide.getBackground().setSolidFill('#FFFFFF');
    }

    // ── Helper: thin horizontal bar ─────────────────────────
    function bar(slide, top, color) {
      var s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, top, W, 2);
      s.getFill().setSolidFill(color || RED);
      s.getBorder().setTransparent();
    }

    // ── Helper: insert styled text box ──────────────────────
    function txt(slide, text, x, y, w, h, opts) {
      opts = opts || {};
      var tb = slide.insertTextBox(String(text || ''), x, y, w, h);
      // getObjectId() is server-assigned — calling it forces GAS to flush the
      // entire pending write batch (including any rect()/bar() shapes that were
      // inserted just before this textbox).  Without this flush, a pending shape
      // and this textbox share the same batch; GAS can mis-resolve the textbox's
      // provisional handle to the shape, causing getText() to hit the wrong
      // element and throw "has no text."
      tb.getObjectId();
      tb.setContentAlignment(
        opts.vAlign === 'top' ? SlidesApp.ContentAlignment.TOP : SlidesApp.ContentAlignment.MIDDLE
      );
      tb.getFill().setTransparent();
      tb.getBorder().setTransparent();
      var ts = tb.getText().getTextStyle();
      ts.setFontFamily('Space Grotesk');
      ts.setFontSize(opts.size || 11);
      ts.setForegroundColor(opts.color || TEXT);
      if (opts.bold)   ts.setBold(true);
      if (opts.italic) ts.setItalic(true);
      if (tb.getText().getParagraphs().length > 0) {
        var ps = tb.getText().getParagraphs()[0].getRange().getParagraphStyle();
        if (opts.align === 'center') ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        else if (opts.align === 'right') ps.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      }
      return tb;
    }

    // ── Helper: filled rectangle ────────────────────────────
    function rect(slide, x, y, w, h, fillColor) {
      var s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
      s.getFill().setSolidFill(fillColor || CARD);
      s.getBorder().setTransparent();
      return s;
    }

    // ── Helper: standard slide title + accent bar ────────────
    function titleBar(slide, mainTitle, sub) {
      txt(slide, mainTitle, PAD, 10, CW, 34, { size: 18, bold: true });
      if (sub) txt(slide, sub, PAD, 44, CW, 16, { size: 9, color: MUTED });
      bar(slide, 62);
    }

    // ── Helper: insert data table ────────────────────────────
    function makeTable(slide, top, hdrs, rows, opts) {
      opts = opts || {};
      var maxR = opts.maxRows || 22;
      var data = rows.slice(0, maxR);
      var nc   = hdrs.length;
      var nr   = data.length + 1; // +1 for header
      if (nr < 2) { nr = 2; data = [hdrs.map(function() { return '—'; })]; }
      var avail = H - top - 10;
      var rh    = Math.max(13, Math.min(20, Math.floor(avail / nr)));
      var th    = rh * nr;
      if (th > avail) th = avail;

      var tbl = slide.insertTable(nr, nc, PAD, top, CW, th);

      // Header row
      for (var c = 0; c < nc; c++) {
        var hc = tbl.getCell(0, c);
        hc.getFill().setSolidFill(RED);
        var ht = hc.getText();
        ht.setText(String(hdrs[c]));
        ht.getTextStyle().setFontFamily('Space Grotesk').setFontSize(7.5).setForegroundColor('#FFFFFF').setBold(true);
        var rightAlign  = opts.rightCols  && opts.rightCols.indexOf(c)  > -1;
        var centerAlign = opts.centerCols && opts.centerCols.indexOf(c) > -1;
        if (ht.getParagraphs().length > 0)
          ht.getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(
            rightAlign ? SlidesApp.ParagraphAlignment.END : centerAlign ? SlidesApp.ParagraphAlignment.CENTER : SlidesApp.ParagraphAlignment.START
          );
      }

      // Data rows
      for (var r = 0; r < data.length; r++) {
        var bg = r % 2 === 0 ? CARD : CARD2;
        for (var c2 = 0; c2 < nc; c2++) {
          var dc  = tbl.getCell(r + 1, c2);
          dc.getFill().setSolidFill(bg);
          var val = data[r][c2];
          var cellStr = (val === null || val === undefined) ? '—' : String(val);
          var dt = dc.getText();
          dt.setText(cellStr);
          dt.getTextStyle().setFontFamily('Space Grotesk').setFontSize(opts.fontSize || 7.5).setForegroundColor(TEXT);
          var rightAlign2  = opts.rightCols  && opts.rightCols.indexOf(c2)  > -1;
          var centerAlign2 = opts.centerCols && opts.centerCols.indexOf(c2) > -1;
          if (dt.getParagraphs().length > 0)
            dt.getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(
              rightAlign2 ? SlidesApp.ParagraphAlignment.END : centerAlign2 ? SlidesApp.ParagraphAlignment.CENTER : SlidesApp.ParagraphAlignment.START
            );
        }
      }

      if (rows.length > maxR) {
        txt(slide, '+ ' + (rows.length - maxR) + ' more rows (truncated)',
            PAD, top + th + 3, CW, 11, { size: 7, color: MUTED, italic: true });
      }
      return tbl;
    }

    // ════════════════════════════════════════════════════════
    // CREATE PRESENTATION
    // ════════════════════════════════════════════════════════
    var pres = SlidesApp.create('PS Monthly Review — ' + p.generatedDate);
    var COMMENT_H = 70; // reserved height at bottom for commentary lines

    // ── SLIDE 1: TITLE ───────────────────────────────────────
    _deckSection = 'slide1-title';
    var s1 = pres.getSlides()[0];
    // Remove default title/subtitle placeholder shapes
    s1.getPageElements().forEach(function(el) { el.remove(); });
    setBg(s1);
    bar(s1, H * 0.38);
    bar(s1, H * 0.62);

    var logo = s1.insertShape(SlidesApp.ShapeType.RECTANGLE, PAD, H * 0.38 + 12, 56, 56);
    logo.getFill().setSolidFill(RED);
    logo.getBorder().setTransparent();
    logo.getText().setText('PS');
    logo.getText().getTextStyle().setFontFamily('Space Grotesk').setFontSize(22).setForegroundColor('#FFFFFF').setBold(true);
    logo.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    if (logo.getText().getParagraphs().length > 0)
      logo.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    txt(s1, 'PS Monthly Review', PAD + 66, H * 0.38 + 10, CW - 66, 38, { size: 28, bold: true });
    txt(s1, 'Redis Professional Services  •  ' + p.generatedDate, PAD + 66, H * 0.38 + 50, CW - 66, 18, { size: 11, color: MUTED });
    footer(s1);


    // ── Helper: insert screenshot image ─────────────────────
    function insertShot(slide, base64Data, yTop, bottomPad, opts) {
      opts = opts || {};
      var ix = opts.x !== undefined ? opts.x : PAD;
      var iw = opts.w !== undefined ? opts.w : CW;
      var y  = yTop      || 68;
      var bp = bottomPad || 0;
      var imgH = H - y - 8 - bp;
      if (!base64Data) {
        var ph = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, PAD, y, CW, imgH);
        ph.getFill().setSolidFill(CARD);
        ph.getBorder().setTransparent();
        ph.getText().setText('⚠  Visit this tab in the dashboard before generating the deck.');
        ph.getText().getTextStyle().setFontFamily('Space Grotesk').setFontSize(11).setForegroundColor(YELLOW);
        ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        if (ph.getText().getParagraphs().length > 0)
          ph.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        return;
      }
      try {
        var imgData = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(imgData, 'image/png', 'screenshot.png');
        slide.insertImage(blob, ix, y, iw, imgH);
      } catch(imgErr) {
        Logger.log('insertShot failed: ' + imgErr);
        var ph2 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, ix, y, iw, imgH);
        ph2.getFill().setSolidFill(CARD);
        ph2.getBorder().setTransparent();
        ph2.getText().setText('⚠  Screenshot data is invalid. Re-open this tab in the dashboard and regenerate.');
        ph2.getText().getTextStyle().setFontFamily('Space Grotesk').setFontSize(11).setForegroundColor(YELLOW);
        ph2.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        if (ph2.getText().getParagraphs().length > 0)
          ph2.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
    }

    // ── Helper: commentary area (red dotted line + single editable text box) ──
    function commentaryArea(slide, yTop) {
      var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, PAD, yTop, PAD + CW, yTop);
      line.getLineFill().setSolidFill(RED);
      line.setWeight(1);
      line.setDashStyle(SlidesApp.DashStyle.DOT);
      var tb = slide.insertTextBox('1:\n2:\n3:', PAD, yTop + 5, CW, 44);
      tb.getText().getTextStyle().setFontFamily('Space Grotesk').setFontSize(9).setForegroundColor(MUTED);
    }

    // ── Helper: slide footer (red rule + Redis logo) ─────────────────────────
    function footer(slide) {
      // 1-px red rule across the very bottom of the slide
      var rule = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, H - 14, W, 1);
      rule.getFill().setSolidFill(RED);
      rule.getBorder().setTransparent();
      // CRITICAL flush: commit the rule shape in its own batch before any
      // subsequent insertTextBox (on this slide or the next) runs.  If the
      // rule's INSERT shares a batch with a later insertTextBox, GAS can
      // mis-resolve the textbox's provisional handle to this Rectangle,
      // causing its RED fill to be applied to the textbox instead of the
      // rule — leaving the rule with a transparent/default fill (invisible).
      rule.getObjectId();

      if (REDIS_LOGO_B64) {
        var imgData = Utilities.base64Decode(REDIS_LOGO_B64);
        var blob    = Utilities.newBlob(imgData, 'image/png', 'redis-logo.png');
        var img     = slide.insertImage(blob, W - PAD - 56, H - 13, 56, 11);
        // Same flush requirement for the image: keeps it out of the next
        // slide's titleBar batch, preventing "has no text" on the Image element.
        img.getObjectId();
      }
    }





    // ── SLIDE 3: PROJECTS BY TIER & OWNER (horizontal stacked bar chart) ──
    _deckSection = 'slide3-tier-owner';
    var s3 = pres.appendSlide();
    setBg(s3);
    var to = p.tierOwner || {};
    titleBar(s3, 'Projects by Owner & Tier', p.generatedDate + '  •  Excludes Sophi Operations');

    var toHdrs3  = to.headers || [];
    var toRows3  = to.rows   || [];

    // Data rows only — strip the TOTAL summary row
    var dataR3  = toRows3.filter(function(r){ return r[0] !== 'TOTAL'; });
    // Tier column names: headers[1 .. length-2]  (drop 'Owner' and 'Total')
    var tierC3  = toHdrs3.slice(1, toHdrs3.length - 1);
    var nTiers3 = tierC3.length;

    // Chart area geometry — right margin reserved for bar totals, extra bottom gap for legend
    var LBL_W3  = 100;                             // owner label column width
    var TOT_W3  = 30;                              // total-label zone right of bars
    var chartL3 = PAD + LBL_W3 + 5;               // x where bars start
    var chartR3 = W - PAD - TOT_W3 - 4;           // x where bars end (leaves room for totals)
    var chartW3 = chartR3 - chartL3;
    var chartT3 = 76;                              // y top of first bar
    var LEG_H3  = 18;                              // legend strip height
    var AXIS_H3 = 0;                               // x-axis labels removed
    var chartB3 = H - 14 - LEG_H3 - AXIS_H3 - 14; // chart bottom — extra 14pt gap for legend

    // Cap owners to what fits at ≥9pt bar height
    var nOwners3 = Math.min(dataR3.length, Math.floor((chartB3 - chartT3) / 10));
    var rows3    = dataR3.slice(0, nOwners3);

    // Bar dimensions
    var totalH3  = chartB3 - chartT3;
    var barGap3  = Math.max(2, Math.floor(totalH3 / nOwners3 * 0.12));
    var barH3    = Math.max(9, Math.floor((totalH3 - barGap3 * (nOwners3 - 1)) / nOwners3));
    var lblSz3   = barH3 >= 14 ? 8 : 7;

    // X-axis scale
    var maxTot3 = 1;
    rows3.forEach(function(r){ var t = r[r.length-1]; if (t > maxTot3) maxTot3 = t; });
    var xMax3  = Math.ceil(maxTot3 / 5) * 5;
    var xStep3 = xMax3 <= 50 ? 5 : 10;

    // Tier colour palette — Redis brand accents (Sky Blue, Purple, Hyper, Dusk, Midnight)
    var TCOL3 = ['#80DBFF', '#C795E3', '#FF4438', '#163341', '#091A23', '#B9C2C6'];
    // Data-label text colour per tier (Midnight on light fills, white on dark fills)
    var TLAB3 = ['#091A23', '#091A23', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#091A23'];

    // ── Vertical grid lines + x-axis labels ──
    for (var xi3 = 0; xi3 <= xMax3; xi3 += xStep3) {
      var gx3 = chartL3 + Math.round((xi3 / xMax3) * chartW3);
      var gl3 = s3.insertShape(SlidesApp.ShapeType.RECTANGLE, gx3, chartT3, 0.5, chartB3 - chartT3);
      gl3.getObjectId();
      gl3.getFill().setSolidFill(xi3 === 0 ? '#9CA3AF' : '#E0E0E0');
      gl3.getBorder().setTransparent();
      // x-axis number labels removed per user request
    }

    // ── Owner bars with data labels + row total ──
    rows3.forEach(function(row3, ri3) {
      var by3     = chartT3 + ri3 * (barH3 + barGap3);
      var rowTot3 = row3[row3.length - 1] || 0;

      // Owner name — right-aligned, vertically centred in bar height
      txt(s3, row3[0], PAD, by3, LBL_W3, barH3,
          { size:lblSz3, color:TEXT, align:'right' });

      // Stacked tier segments
      var segX3 = chartL3;
      for (var ti3 = 0; ti3 < nTiers3; ti3++) {
        var cnt3 = row3[ti3 + 1] || 0;
        if (cnt3 <= 0) continue;
        var segW3 = Math.max(1, Math.round((cnt3 / xMax3) * chartW3));

        // Segment rectangle
        var seg3 = s3.insertShape(SlidesApp.ShapeType.RECTANGLE, segX3, by3, segW3, barH3);
        seg3.getObjectId();
        seg3.getFill().setSolidFill(TCOL3[ti3 % TCOL3.length]);
        seg3.getBorder().setTransparent();

        // Data label (only when segment is wide enough to hold a number)
        if (segW3 >= 14) {
          txt(s3, String(cnt3), segX3, by3, segW3, barH3,
              { size:7, color:TLAB3[ti3 % TLAB3.length], align:'center' });
        }

        segX3 += segW3;
      }

      // Row total to the right of the bar — bold, vertically centred
      txt(s3, String(rowTot3), chartR3 + 4, by3, TOT_W3 - 2, barH3,
          { size:8, bold:true, color:TEXT });
    });

    // ── Legend (centred, with clear gap below axis labels) ──
    var legItemW3  = 65;
    var legY3      = chartB3 + AXIS_H3 + 8;   // +8 breathing room above legend
    var legStartX3 = chartL3 + Math.floor((chartW3 - nTiers3 * legItemW3) / 2);
    tierC3.forEach(function(tier3, ti3) {
      var lx3  = legStartX3 + ti3 * legItemW3;
      var dot3 = s3.insertShape(SlidesApp.ShapeType.RECTANGLE, lx3, legY3 + 3, 9, 9);
      dot3.getObjectId();
      dot3.getFill().setSolidFill(TCOL3[ti3 % TCOL3.length]);
      dot3.getBorder().setTransparent();
      txt(s3, tier3, lx3 + 11, legY3, legItemW3 - 13, LEG_H3 - 2,
          { size:8, color:TEXT, vAlign:'top' });
    });

    footer(s3);


    // ── SLIDE 4: GO LIVES (screenshot) ───────────────────────
    _deckSection = 'slide4-go-lives';
    var s4 = pres.appendSlide();
    setBg(s4);
    titleBar(s4, 'Go Live Summary', p.generatedDate);
    insertShot(s4, (p.screenshots || {}).slide4, 110, COMMENT_H);
    commentaryArea(s4, H - COMMENT_H);
    footer(s4);


    // ── SLIDE 5: ANNUAL PLAN & RE SUMMARY (scorecard cards) ──
    _deckSection = 'slide5-annual-plan';
    var s5 = pres.appendSlide();
    setBg(s5);
    titleBar(s5, 'Annual Plan & RE Summary', 'Project count by region and utilization ratio');

    var aps5     = p.annualPlanSummary || {};
    var hdrs5    = aps5.headers || ['Region','<50%','50–89%','≥90%','Total'];
    var allR5    = aps5.rows || [];
    var regR5    = allR5.filter(function(r){ return r[0] !== 'TOTAL'; });
    var totR5    = allR5.filter(function(r){ return r[0] === 'TOTAL'; })[0] || ['TOTAL',0,0,0,0];
    var cards5   = regR5.concat([totR5]);

    // Card geometry — fills from just below title bar to commentary strip
    var nC5   = cards5.length || 4;
    var cGap5 = 12;
    var cW5   = Math.floor((W - 2*PAD - (nC5-1)*cGap5) / nC5);
    var cTop5 = 74;
    var cH5   = H - COMMENT_H - cTop5 - 10;  // e.g. 405-70-74-10 = 251

    // Bucket row layout (positions relative to card top)
    var cPad5    = 8;   // inner horizontal padding
    var bkStart5 = 96;  // y where first bucket row begins
    var bkH5     = Math.floor((cH5 - bkStart5 - 4) / 3);  // height of each bucket row

    // Bucket definitions: [ header-index, bg-on-white, bg-on-dark ]
    var bkDefs5 = [
      [1, '#FFCDD6', '#2D2A6B'],   // <50%   — pink / dark indigo
      [2, '#FFF3C4', '#2D2A6B'],   // 50-89% — yellow / dark indigo
      [3, '#D4F0D4', '#2D2A6B']    //  ≥90%  — green / dark indigo
    ];

    // Theme colours
    var INDIGO5   = '#1E1B4B';   // dark indigo — text on white cards
    var GREY5     = '#6B7280';   // subdued label on white cards
    var TOTBG5    = '#1E1B4B';   // TOTAL card background
    var TOTSUB5   = '#9B8FD4';   // "across all regions" label on dark card

    cards5.forEach(function(row5, ci5) {
      var cx5    = PAD + ci5 * (cW5 + cGap5);
      var cy5    = cTop5;
      var isTot5 = row5[0] === 'TOTAL';
      var cardBg5 = isTot5 ? TOTBG5  : '#FFFFFF';
      var txtC5   = isTot5 ? '#FFFFFF': INDIGO5;
      var subC5   = isTot5 ? TOTSUB5 : GREY5;

      // ── Card background rectangle ──
      var cRect5 = s5.insertShape(SlidesApp.ShapeType.RECTANGLE, cx5, cy5, cW5, cH5);
      cRect5.getObjectId();   // CRITICAL flush
      cRect5.getFill().setSolidFill(cardBg5);
      if (isTot5) {
        cRect5.getBorder().setTransparent();
      } else {
        cRect5.getBorder().getLineFill().setSolidFill('#E0E0E0');
        cRect5.getBorder().setWeight(0.5);
      }

      // ── Region / TOTAL label ──
      txt(s5, row5[0],
          cx5+cPad5, cy5+10, cW5-2*cPad5, 18,
          { bold:true, size:10, color:txtC5, vAlign:'top' });

      // ── Big count number ──
      txt(s5, String(row5[4] != null ? row5[4] : 0),
          cx5+cPad5, cy5+28, cW5-2*cPad5, 50,
          { bold:true, size:36, color:txtC5, vAlign:'top' });

      // ── Sub-label ("total assessed" / "across all regions") ──
      txt(s5, isTot5 ? 'across all regions' : 'total assessed',
          cx5+cPad5, cy5+78, cW5-2*cPad5, 12,
          { size:7, color:subC5, italic:true, vAlign:'top' });

      // ── Three bucket rows ──
      bkDefs5.forEach(function(bd5, bi5) {
        var bx5  = cx5 + cPad5;
        var by5  = cy5 + bkStart5 + bi5 * (bkH5 + 2);
        var bw5  = cW5 - 2*cPad5;
        var bBg5 = isTot5 ? bd5[2] : bd5[1];
        var bTxt5= isTot5 ? '#FFFFFF' : '#374151';

        // Bucket background rect
        var bkR5 = s5.insertShape(SlidesApp.ShapeType.RECTANGLE, bx5, by5, bw5, bkH5);
        bkR5.getObjectId();   // CRITICAL flush
        bkR5.getFill().setSolidFill(bBg5);
        bkR5.getBorder().setTransparent();

        // Label (left ~60%)
        txt(s5, hdrs5[bd5[0]],
            bx5+4, by5+4, bw5*0.6-4, bkH5-8,
            { size:9, color:bTxt5, vAlign:'top' });

        // Count (right ~40%, right-aligned bold)
        txt(s5, String(row5[bd5[0]] != null ? row5[bd5[0]] : 0),
            bx5+bw5*0.6, by5+4, bw5*0.4, bkH5-8,
            { size:11, bold:true, color:bTxt5, align:'right', vAlign:'top' });
      });
    });

    commentaryArea(s5, H - COMMENT_H);
    footer(s5);


    // ── SLIDE 5b: UTILIZATION SUMMARY ────────────────────────
    _deckSection = 'slide5b-utilization-summary';
    var sU = pres.appendSlide();
    setBg(sU);
    var ut = p.utilization || {};
    var utRes     = ut.resources      || [];
    var utRegions = ut.regionSummary  || [];
    titleBar(sU, 'Utilization Summary', (ut.quarter || '') + '  •  Excludes ramping resources');

    // ── Helper: colour-code by billable % ──
    function utColor(n) {
      if (n >= 85) return '#16A34A';   // green
      if (n >= 70) return '#CA8A04';   // amber
      return '#DC2626';                // red
    }

    // ── Helper: draw a utilization panel (title + header + rows) ──
    function drawUtilPanel(slide, px, py, pw, ph, title, panelRows) {
      var TITLE_H = 15;
      var HDR_H   = 14;
      var areaH   = ph - TITLE_H - HDR_H - 2; // 2pt gap between header and first data row
      var nVis    = Math.min(panelRows.length, Math.floor(areaH / 10));
      var rowH    = nVis > 0 ? Math.min(12, Math.floor(areaH / nVis)) : 11;
      rowH        = Math.max(10, rowH);
      var emptyPanel = (panelRows.length === 0);
      var c1W     = Math.floor(pw * 0.56);
      var c2W     = Math.floor(pw * 0.22);
      var c3W     = pw - c1W - c2W;

      // Title bar (RED)
      var tR = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, px, py, pw, TITLE_H);
      tR.getObjectId();
      tR.getFill().setSolidFill(RED);
      tR.getBorder().setTransparent();
      txt(slide, title, px+6, py, pw-8, TITLE_H,
          { bold:true, size:8, color:'#FFFFFF' });

      // Header row (dark charcoal)
      var hY = py + TITLE_H;
      var hR = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, px, hY, pw, HDR_H);
      hR.getObjectId();
      hR.getFill().setSolidFill('#374151');
      hR.getBorder().setTransparent();
      txt(slide, 'Name',     px+4,      hY, c1W,    HDR_H, { bold:true, size:7, color:'#FFFFFF' });
      txt(slide, 'Billable', px+c1W,    hY, c2W,    HDR_H, { bold:true, size:7, color:'#FFFFFF', align:'center' });
      txt(slide, 'PTO-Adj',  px+c1W+c2W,hY, c3W,   HDR_H, { bold:true, size:7, color:'#FFFFFF', align:'center' });

      // Empty-state: no resource data in this payload version
      if (emptyPanel) {
        var emR = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, px, hY + HDR_H, pw, areaH);
        emR.getObjectId();
        emR.getFill().setSolidFill(CARD);
        emR.getBorder().setTransparent();
        txt(slide, 'Resource data not found in payload.',
            px+6, hY + HDR_H + 10, pw-12, 14,
            { size:8, color:MUTED, italic:true, vAlign:'top' });
        txt(slide, 'Regenerate after deploying the latest dashboard code.',
            px+6, hY + HDR_H + 24, pw-12, 14,
            { size:7.5, color:MUTED, italic:true, vAlign:'top' });
        return;
      }

      // Data rows (2pt clear gap after header so header text isn't covered)
      var rY = hY + HDR_H + 2;
      panelRows.slice(0, nVis).forEach(function(row, ri) {
        var bg  = ri % 2 === 0 ? CARD : CARD2;
        var rR  = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, px, rY, pw, rowH);
        rR.getObjectId();
        rR.getFill().setSolidFill(bg);
        rR.getBorder().setTransparent();
        txt(slide, row.name,            px+4,       rY, c1W-4, rowH, { size:7, color:TEXT });
        txt(slide, row.billNum + '%',   px+c1W,     rY, c2W,   rowH, { size:7, bold:true, color:utColor(row.billNum), align:'center' });
        txt(slide, row.ptoNum  + '%',   px+c1W+c2W, rY, c3W,   rowH, { size:7, color:TEXT, align:'center' });
        rY += rowH;
      });

      if (panelRows.length > nVis) {
        txt(slide, '+' + (panelRows.length - nVis) + ' more (truncated)',
            px+4, rY+2, pw-8, 10, { size:6.5, color:MUTED, italic:true, vAlign:'top' });
      }
    }

    if (utRes.length === 0 && utRegions.length === 0) {
      // No data guard
      txt(sU, 'No utilization data available for this period.',
          PAD, 120, CW, 20, { size:11, color:MUTED, align:'center' });
    } else {
      // ── Region + Global cards ──
      // Global card: use pre-computed hour-weighted summary from payload (matches dashboard exactly).
      // Falls back to simple region average only if globalSummary wasn't in the payload.
      var utGlobal    = ut.globalSummary || {};
      var globalBillNum = parseInt(utGlobal.billable) || (utRegions.length
        ? Math.round(utRegions.reduce(function(s,rg){ return s+(parseInt(rg.billable)||0); },0)/utRegions.length) : 0);
      var globalPtoNum  = parseInt(utGlobal.ptoAdj)   || (utRegions.length
        ? Math.round(utRegions.reduce(function(s,rg){ return s+(parseInt(rg.ptoAdj)||0);   },0)/utRegions.length) : 0);

      var cardDataU = utRegions.map(function(rg) {
        return { label: rg.region, bill: rg.billable, pto: rg.ptoAdj,
                 billN: parseInt(rg.billable) || 0 };
      });
      cardDataU.push({ label: 'Global', bill: globalBillNum + '%', pto: globalPtoNum + '%',
                       billN: globalBillNum });

      var nCU   = cardDataU.length;
      var cGapU = 10;
      var cWU   = Math.floor((W - 2*PAD - (nCU - 1) * cGapU) / nCU);
      var cTopU = 74;
      var cHU   = 64;
      var halfW_U = Math.floor((cWU - 14) / 2);

      cardDataU.forEach(function(cd, ci) {
        var cx = PAD + ci * (cWU + cGapU);
        var cy = cTopU;
        var ac = utColor(cd.billN);

        // Card background
        var cRect = sU.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cy, cWU, cHU);
        cRect.getObjectId();
        cRect.getFill().setSolidFill('#FFFFFF');
        cRect.getBorder().getLineFill().setSolidFill('#E0E0E0');
        cRect.getBorder().setWeight(0.5);

        // Bottom accent bar (colour-coded)
        var acBar = sU.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cy + cHU - 4, cWU, 4);
        acBar.getObjectId();
        acBar.getFill().setSolidFill(ac);
        acBar.getBorder().setTransparent();

        // Region label
        txt(sU, cd.label, cx+6, cy+4, cWU-12, 11,
            { bold:true, size:8, color:'#1E1B4B', vAlign:'top' });

        // BILLABLE column (left half)
        txt(sU, 'BILLABLE', cx+7, cy+20, halfW_U, 9,
            { size:6, color:'#9CA3AF', vAlign:'top' });
        txt(sU, cd.bill, cx+7, cy+29, halfW_U, 26,
            { bold:true, size:17, color:ac, vAlign:'top' });

        // PTO-ADJ column (right half)
        var rx = cx + 7 + halfW_U + 4;
        txt(sU, 'PTO-ADJ', rx, cy+20, halfW_U, 9,
            { size:6, color:'#9CA3AF', vAlign:'top' });
        txt(sU, cd.pto, rx, cy+29, halfW_U, 26,
            { bold:true, size:17, color:'#374151', vAlign:'top' });
      });

      // ── Two resource tables ──
      var consultants = utRes.filter(function(r){ return r.role === 'Consultant'; });
      var ems         = utRes.filter(function(r){ return r.role !== 'Consultant'; });
      consultants.sort(function(a,b){ return b.billNum - a.billNum; });
      ems.sort(function(a,b){ return b.billNum - a.billNum; });

      var tblTop  = cTopU + cHU + 8;
      var tblH    = H - 14 - tblTop;
      var tblGap  = 8;
      var tblW    = Math.floor((W - 2*PAD - tblGap) / 2);

      drawUtilPanel(sU, PAD,           tblTop, tblW, tblH, 'CONSULTANTS',          consultants);
      drawUtilPanel(sU, PAD+tblW+tblGap, tblTop, tblW, tblH, 'ENGAGEMENT MANAGERS', ems);
    }

    footer(sU);


    // ── SLIDE 6a: ANNUAL PLAN < 50% UTILIZATION (screenshot) ───
    _deckSection = 'slide6a-under50';
    var s6a = pres.appendSlide();
    setBg(s6a);
    titleBar(s6a, 'Annual Plan & RE  —  Under 50% Utilization', p.generatedDate);
    insertShot(s6a, (p.screenshots || {}).slide6a, 68, COMMENT_H);
    commentaryArea(s6a, H - COMMENT_H);
    footer(s6a);

    // ── SLIDE 6b: ANNUAL PLAN 50–75% UTILIZATION (screenshot) ──
    _deckSection = 'slide6b-50to75';
    var s6b = pres.appendSlide();
    setBg(s6b);
    titleBar(s6b, 'Annual Plan & RE  —  50–75% Utilization', p.generatedDate);
    insertShot(s6b, (p.screenshots || {}).slide6b, 68, COMMENT_H);
    commentaryArea(s6b, H - COMMENT_H);
    footer(s6b);




    // ── SLIDE 8: BOOKINGS FORECAST (screenshot) ──────────
    _deckSection = 'slide8-bookings';
    var s8 = pres.appendSlide();
    setBg(s8);
    titleBar(s8, 'Quarterly Bookings Forecast', 'FY' + (p.currentFY || ''));
    insertShot(s8, (p.screenshots || {}).slide8, 115, COMMENT_H);
    commentaryArea(s8, H - COMMENT_H);
    footer(s8);


    // ── SLIDE 9: BACKLOG ANALYSIS (screenshot) ───────────
    _deckSection = 'slide9-backlog';
    var s9 = pres.appendSlide();
    setBg(s9);
    titleBar(s9, 'Backlog Analysis', 'FY' + (p.currentFY || ''));
    insertShot(s9, (p.screenshots || {}).slide9, 100, COMMENT_H, {x: 14, w: 692});
    commentaryArea(s9, H - COMMENT_H);
    footer(s9);


    // ── SLIDE 10: P&L STATUS ────────────────────────────
    _deckSection = 'slide10-pl';
    var s10 = pres.appendSlide();
    setBg(s10);
    titleBar(s10, 'P&L Status', p.generatedDate);
    var ph = s10.insertShape(SlidesApp.ShapeType.RECTANGLE, PAD, 74, CW, H - 84 - COMMENT_H);
    ph.getFill().setSolidFill(CARD);
    ph.getBorder().setTransparent();
    ph.getText().setText('[ Paste P&L screenshot here ]');
    ph.getText().getTextStyle().setFontFamily('Space Grotesk').setFontSize(14).setForegroundColor(DIM).setItalic(true);
    ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    if (ph.getText().getParagraphs().length > 0)
      ph.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    commentaryArea(s10, H - COMMENT_H);
    footer(s10);


    return JSON.stringify({ success: true, url: pres.getUrl(), title: pres.getName() });
  } catch(e) {
    return JSON.stringify({ success: false, error: '[' + _deckSection + '] ' + e.toString() });
  }
}

function writeKPIsToSheet(kpiArray) {
  try {
    var TARGET_SS_ID = '1HNr9QtJ_De_ZLQcHx6-rDBgHze3P7InQfFVQGax8Vjg';
    var targetSheet = SpreadsheetApp.openById(TARGET_SS_ID).getSheetByName('PS');
    if (!targetSheet) return JSON.stringify({ success: false, error: 'Sheet "PS" not found in target spreadsheet' });
    var parsed = (typeof kpiArray === 'string') ? JSON.parse(kpiArray) : kpiArray;
    parsed.forEach(function(item) {
      if (item.value !== null && item.value !== undefined) {
        targetSheet.getRange(item.cell).setValue(item.value);
      }
    });
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ============================================================
// MISSING TIMECARD EMAIL REMINDER
// Runs every Monday at 9am via a time-driven trigger.
// Sends one email per resource with < 40 hrs logged last week,
// CC'd to their manager.
//
// SETUP: Run createMissingTimecardTrigger() once from the
// GAS editor (Run menu) to install the weekly trigger.
// ============================================================

/**
 * Sends missing timecard reminder emails for the previous Mon–Sun week.
 * Called automatically by the Monday 9am trigger.
 */
// Resources on an Israel (Sun–Thu) work week. Their 40-hour lookback uses a
// Sunday–Saturday window instead of the standard Monday–Sunday. Emails must be
// lowercase. Add more Israel-based resources here as needed.
var ISRAEL_USERS = [
  'asaf.hirshberg@redis.com'
];

function sendMissingTimecardEmails() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Locate the two required sheets ──────────────────────────
  var userSheet  = ss.getSheetByName('PS User records');
  var hoursSheet = ss.getSheetByName('Hours Data');

  if (!userSheet || !hoursSheet) {
    Logger.log('sendMissingTimecardEmails: required sheets not found. Expected "PS User records" and "Hours Data".');
    return;
  }

  // ── Previous completed week: Sunday–Saturday ───────────────
  // The timecard week runs Sunday–Saturday for ALL resources. Finds the last
  // COMPLETED Sun–Sat week regardless of which day the script runs.
  // e.g. for a Monday Jul 27 run → Jul 19–25.
  var today = new Date(); today.setHours(0,0,0,0);
  var dow = today.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  var daysToLastSat = ((dow + 1) % 7) || 7; // days back to the most recent completed Saturday
  var wkEnd = new Date(today); wkEnd.setDate(today.getDate() - daysToLastSat);
  var wkStart = new Date(wkEnd); wkStart.setDate(wkEnd.getDate() - 6);
  wkStart.setHours(0,0,0,0);
  wkEnd.setHours(23,59,59,999);

  // Israel-based resources (ISRAEL_USERS) also use Sunday–Saturday, so their
  // window is currently identical to everyone else's. Kept as its own pair so
  // the two can diverge again later without a refactor.
  var ilStart = new Date(wkStart), ilEnd = new Date(wkEnd);

  function fmtDate(d) {
    return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  }
  function rangeStr(a, b) {
    return fmtDate(a) + ' – ' + fmtDate(new Date(b.getFullYear(), b.getMonth(), b.getDate()));
  }
  var stdWeekRange = rangeStr(wkStart, wkEnd);
  var ilWeekRange  = rangeStr(ilStart, ilEnd);

  // ── Build hours-per-person maps for both windows ────────────
  var hData = hoursSheet.getDataRange().getValues();
  var hHdrs = hData[0].map(String);
  var iHN = hHdrs.indexOf('Work: Owner Name');
  if (iHN < 0) iHN = hHdrs.indexOf('Work Owner Name');
  var iHD = hHdrs.indexOf('Date');
  var iHH = hHdrs.indexOf('Hours (Number)');

  if (iHN < 0 || iHD < 0 || iHH < 0) {
    Logger.log('sendMissingTimecardEmails: hours data columns missing.');
    return;
  }

  var stdHoursMap = {}, ilHoursMap = {};
  for (var hi = 1; hi < hData.length; hi++) {
    var hr = hData[hi];
    var hName = String(hr[iHN] || '').trim().toLowerCase();
    if (!hName) continue;
    var rawD = hr[iHD];
    if (!rawD) continue;
    // Handle both Date objects (from Sheets date cells) and strings robustly.
    // Avoid String(dateObj) → new Date(string) round-trip which causes timezone shifts.
    var hd = (rawD instanceof Date) ? new Date(rawD) : new Date(String(rawD).trim());
    if (isNaN(hd)) continue;
    hd.setHours(0,0,0,0);
    var hrs = parseFloat(hr[iHH]) || 0;
    if (hd >= wkStart && hd <= wkEnd) stdHoursMap[hName] = (stdHoursMap[hName] || 0) + hrs;
    if (hd >= ilStart && hd <= ilEnd) ilHoursMap[hName]  = (ilHoursMap[hName]  || 0) + hrs;
  }

  // ── Read user records and send emails ───────────────────────
  var uData = userSheet.getDataRange().getValues();
  var uHdrs = uData[0].map(String);
  var iUN  = uHdrs.indexOf('Full Name');  if (iUN  < 0) iUN  = uHdrs.indexOf('Name');
  var iUE  = uHdrs.indexOf('Email');
  var iUME = uHdrs.indexOf('Manager Email');
  var iUMN = uHdrs.indexOf('Manager Name'); if (iUMN < 0) iUMN = uHdrs.indexOf('Manager');
  var iUSD = uHdrs.indexOf('User Start Day'); if (iUSD < 0) iUSD = uHdrs.indexOf('Start Date');

  if (iUN < 0 || iUE < 0 || iUME < 0) {
    Logger.log('sendMissingTimecardEmails: user record columns missing (need Full Name, Email, Manager Email).');
    return;
  }

  var sent = 0, skipped = 0, errors = [];
  var today0 = new Date(); today0.setHours(0,0,0,0);

  for (var ui = 1; ui < uData.length; ui++) {
    var ur    = uData[ui];
    var name  = String(ur[iUN]  || '').trim();  if (!name)  continue;
    var email = String(ur[iUE]  || '').trim();  if (!email) { skipped++; continue; }
    var managerEmail = String(ur[iUME] || '').trim();
    var managerName  = iUMN >= 0 ? String(ur[iUMN] || '').trim() : '';

    // Skip resources still within their 90-day ramp window
    if (iUSD >= 0) {
      var startStr = String(ur[iUSD] || '').trim();
      if (startStr) {
        var startD = new Date(startStr);
        if (!isNaN(startD) && today0 < new Date(startD.getTime() + 90 * 86400000)) {
          skipped++;
          continue; // still ramping — no email
        }
      }
    }

    // Israel-based resources use the Sunday–Saturday window; everyone else Mon–Sun.
    var isIsrael    = ISRAEL_USERS.indexOf(email.toLowerCase()) >= 0;
    var weekRange   = isIsrael ? ilWeekRange : stdWeekRange;
    var windowStart = isIsrael ? ilStart : wkStart;

    var logged = (isIsrael ? ilHoursMap : stdHoursMap)[name.toLowerCase()] || 0;
    if (logged >= 40) continue; // timecard complete — no email needed

    var loggedDisplay = (logged % 1 === 0) ? String(logged) : logged.toFixed(1);
    var firstName = name.split(' ')[0];

    var missing = (40 - logged);
    var missingDisplay = (missing % 1 === 0) ? String(missing) : missing.toFixed(1);

    var subject = 'Reminder: Incomplete Timecard for the Week of ' + fmtDate(windowStart);

    // Plain-text fallback (for email clients that don't render HTML)
    var body =
      'Hi ' + firstName + ',\n\n' +
      'This is a friendly reminder that your timecard for the week of ' + weekRange + ' appears to be incomplete.\n\n' +
      '  Hours logged:   ' + loggedDisplay + ' hrs\n' +
      '  Hours expected: 40 hrs\n' +
      '  Still missing:  ' + missingDisplay + ' hrs\n\n' +
      'Please log your remaining ' + missingDisplay + ' hour(s) in the timecard system at your earliest convenience.\n\n' +
      'If you have an approved exception or believe this message is in error, please reach out to ' +
      (managerName ? managerName : 'your manager') + ' directly.\n\n' +
      'Note: The data used for this email could be up to 2 hours old.\n\n' +
      'Thank you,\nPS Operations';

    // HTML version — table-based layout required for Gmail dark background support
    // (Gmail strips background-color from div wrappers; bgcolor attr on td is reliable)
    var htmlBody =
      '<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#0f172a" style="background-color:#0f172a;">' +
      '<tr><td align="center" style="padding:24px 16px;background-color:#0f172a;">' +
      '<table width="560" cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,sans-serif;max-width:560px;">' +

        // Header
        '<tr><td bgcolor="#CC1A14" style="background-color:#CC1A14;padding:18px 24px;border-radius:8px 8px 0 0;">' +
          '<span style="color:#ffffff;font-size:16px;font-weight:600;">PS Operations</span>' +
        '</td></tr>' +

        // Body
        '<tr><td bgcolor="#1e293b" style="background-color:#1e293b;padding:28px 24px;border-radius:0 0 8px 8px;border:1px solid #334155;border-top:none;">' +
          '<p style="margin:0 0 16px;color:#f1f5f9;font-size:14px;font-family:Arial,sans-serif;">Hi ' + firstName + ',</p>' +
          '<p style="margin:0 0 20px;color:#cbd5e1;font-size:14px;line-height:1.6;font-family:Arial,sans-serif;">This is a friendly reminder that your timecard for the week of <strong style="color:#f1f5f9;">' + weekRange + '</strong> appears to be incomplete.</p>' +

          // Hours summary card (nested table)
          '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 20px;">' +
          '<tr><td bgcolor="#0f172a" style="background-color:#0f172a;border:1px solid #334155;border-radius:8px;padding:20px;">' +
            '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
            '<tr>' +
              '<td width="33%" align="center" style="padding:8px;">' +
                '<p style="margin:0 0 6px;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;font-family:Arial,sans-serif;">Hours logged</p>' +
                '<p style="margin:0;font-size:24px;font-weight:700;color:#f87171;font-family:Arial,sans-serif;">' + loggedDisplay + ' hrs</p>' +
              '</td>' +
              '<td width="33%" align="center" style="padding:8px;border-left:1px solid #334155;border-right:1px solid #334155;">' +
                '<p style="margin:0 0 6px;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;font-family:Arial,sans-serif;">Hours expected</p>' +
                '<p style="margin:0;font-size:24px;font-weight:700;color:#f1f5f9;font-family:Arial,sans-serif;">40 hrs</p>' +
              '</td>' +
              '<td width="33%" align="center" style="padding:8px;">' +
                '<p style="margin:0 0 6px;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:0.06em;font-family:Arial,sans-serif;">Still missing</p>' +
                '<p style="margin:0;font-size:24px;font-weight:700;color:#fbbf24;font-family:Arial,sans-serif;">' + missingDisplay + ' hrs</p>' +
              '</td>' +
            '</tr>' +
            '</table>' +
          '</td></tr>' +
          '</table>' +

          '<p style="margin:0 0 16px;color:#cbd5e1;font-size:14px;line-height:1.6;font-family:Arial,sans-serif;">Please log your remaining <strong style="color:#f1f5f9;">' + missingDisplay + ' hour(s)</strong> in the timecard system at your earliest convenience.</p>' +
          '<p style="margin:0 0 28px;color:#cbd5e1;font-size:14px;line-height:1.6;font-family:Arial,sans-serif;">If you have an approved exception or believe this message is in error, please reach out to <strong style="color:#f1f5f9;">' + (managerName ? managerName : 'your manager') + '</strong> directly.</p>' +

          // Divider
          '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 20px;"><tr><td style="border-top:1px solid #334155;font-size:0;">&nbsp;</td></tr></table>' +

          '<p style="margin:0;font-size:14px;color:#64748b;font-family:Arial,sans-serif;">Thank you,<br><strong style="color:#94a3b8;">PS Operations</strong></p>' +
        '</td></tr>' +

        // Footer
        '<tr><td align="center" style="padding:12px 0;">' +
          '<p style="margin:0 0 4px;font-size:11px;color:#475569;font-family:Arial,sans-serif;">Note: The data used for this email could be up to 2 hours old.</p>' +
          '<p style="margin:0;font-size:11px;color:#475569;font-family:Arial,sans-serif;">This is an automated reminder. Please do not reply to this email.</p>' +
        '</td></tr>' +

      '</table>' +
      '</td></tr>' +
      '</table>';

    var mailOpts = { name: 'PS Operations', htmlBody: htmlBody };
    if (managerEmail) mailOpts.cc = managerEmail;

    try {
      MailApp.sendEmail(email, subject, body, mailOpts);
      Logger.log('✓ ' + name + ' <' + email + '> — ' + loggedDisplay + ' hrs logged' +
                 (managerEmail ? ' (CC: ' + managerEmail + ')' : ''));
      sent++;
    } catch(e) {
      Logger.log('✗ Error sending to ' + name + ': ' + e.message);
      errors.push(name);
    }
  }

  Logger.log('Done. Sent: ' + sent + ', Skipped (no email): ' + skipped +
             (errors.length ? ', Errors: ' + errors.join(', ') : '.'));
}

/**
 * Run this function ONCE from the GAS editor to install the weekly trigger.
 * It removes any existing trigger for sendMissingTimecardEmails first,
 * so it is safe to re-run without creating duplicates.
 */
function createMissingTimecardTrigger() {
  // Remove any pre-existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendMissingTimecardEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Install new Monday 9am trigger in the script's timezone
  ScriptApp.newTrigger('sendMissingTimecardEmails')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  Logger.log('Trigger created: sendMissingTimecardEmails fires every Monday at 9am (' +
             Session.getScriptTimeZone() + ').');
}
