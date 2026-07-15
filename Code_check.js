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

    // ── Color constants (white theme) ───────────────────────
    var BG     = '#FFFFFF';
    var CARD   = '#F5F5F5';
    var CARD2  = '#EBEBEB';
    var RED    = '#FF4438';
    var TEXT   = '#1A1A2E';
    var MUTED  = '#595959';
    var DIM    = '#AAAAAA';
    var GREEN  = '#107C41';
    var YELLOW = '#C27C0E';
    var BLUE   = '#1565C0';

    // ── Redis logo (base64 PNG, 300×93px wordmark) ──────────
    var REDIS_LOGO_B64 = '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';

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
      tb.setContentAlignment(
        opts.vAlign === 'top' ? SlidesApp.ContentAlignment.TOP : SlidesApp.ContentAlignment.MIDDLE
      );
      tb.getFill().setTransparent();
      tb.getBorder().setTransparent();
      var ts = tb.getText().getTextStyle();
      ts.setFontFamily('Arial');
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
        ht.getTextStyle().setFontFamily('Arial').setFontSize(7.5).setForegroundColor('#FFFFFF').setBold(true);
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
          dt.getTextStyle().setFontFamily('Arial').setFontSize(opts.fontSize || 7.5).setForegroundColor(TEXT);
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
    logo.getText().getTextStyle().setFontFamily('Arial').setFontSize(22).setForegroundColor('#FFFFFF').setBold(true);
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
        ph.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor(YELLOW);
        ph.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        if (ph.getText().getParagraphs().length > 0)
          ph.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        return;
      }
      var imgData = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(imgData, 'image/png', 'screenshot.png');
      slide.insertImage(blob, ix, y, iw, imgH);
    }

    // ── Helper: commentary area (separator bar + editable text box) ──
    // NOTE: Uses a thin filled rectangle instead of insertLine().
    // Line objects (returned by insertLine) have no text body in the GAS
    // Slides API, which can cause "The object has no text" runtime errors
    // when the API resolves shape references on the same slide.
    function commentaryArea(slide, yTop) {
      // Thin red separator bar — a 1-pt tall rectangle, visually identical to a line
      var sep = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, PAD, yTop, CW, 1);
      sep.getFill().setSolidFill(RED);
      sep.getBorder().setTransparent();

      // Editable placeholder — presenter fills in commentary
      var tb = slide.insertTextBox('1:\n2:\n3:', PAD, yTop + 4, CW, 46);
      tb.getFill().setTransparent();
      tb.getBorder().setTransparent();
      tb.getText().getTextStyle().setFontFamily('Arial').setFontSize(9).setForegroundColor(MUTED);
    }


    // ── Helper: footer bar (logo + separator + label) ───────
    // IMPORTANT: txt() is called FIRST so its insertTextBox + getText() batch flush
    // resolves before insertImage is added to any batch. This prevents GAS Slides'
    // element-ID mapping from resolving the tb reference onto the preceding Image
    // element (which has no text body) instead of the TextBox — the root cause of
    // the "The object has no text" error at a consistent element index.
    function footer(slide) {
      // "Professional Services" label (text box must be inserted before the image)
      txt(slide, 'Professional Services', W - PAD - 130, 388, 130, 14, { size: 7.5, color: MUTED, align: 'right' });
      // Redis wordmark logo — inserted after text box so it gets its own batch slot
      try {
        var logoBytes = Utilities.base64Decode(REDIS_LOGO_B64);
        var logoBlob  = Utilities.newBlob(logoBytes, 'image/png', 'redis_logo.png');
        slide.insertImage(logoBlob, 23.5, 386, 50, 15.5);
      } catch(imgErr) {
        Logger.log('footer: logo image insertion failed — ' + imgErr);
      }
    }

    // ── SLIDE 2: ACTIVE PROJECTS ─────────────────────────────
    _deckSection = 'slide2-active-projects';
    var s2 = pres.appendSlide();
    setBg(s2);
    titleBar(s2, 'Active Projects Summary', p.generatedDate);

    var ap = p.activeProjects || {};
    var sCounts   = ap.stageCounts   || {};
    var sProjects = ap.stageProjects || {};

    // 7 KPI boxes across the top
    var kpiDef = [
      { label: 'Total Open',         count: ap.totalCount || 0,                               color: BLUE   },
      { label: 'Awaiting Handoff',   count: sCounts['Awaiting Handoff']              || 0,    color: RED    },
      { label: 'Awaiting Kickoff',   count: sCounts['Awaiting Kickoff']              || 0,    color: YELLOW },
      { label: 'In Progress',        count: sCounts['In Progress']                   || 0,    color: GREEN  },
      { label: 'On Hold',            count: sCounts['On Hold']                       || 0,    color: MUTED  },
      { label: 'NL Awaiting GL',     count: sCounts['New Logo Awaiting Go Live']     || 0,    color: BLUE   },
      { label: 'NL Awaiting Close',  count: sCounts['New Logo Live / Awaiting Closure'] || 0, color: GREEN  },
    ];
    var kbw = Math.floor((CW - 6 * 5) / 7);
    kpiDef.forEach(function(b, i) {
      var bx = PAD + i * (kbw + 5);
      rect(s2, bx, 68, kbw, 40, CARD);
      txt(s2, String(b.count), bx, 70, kbw, 22, { size: 15, bold: true, color: b.color, align: 'center' });
      txt(s2, b.label, bx, 93, kbw, 13, { size: 6.5, color: MUTED, align: 'center' });
    });

    // 6 project-list cells in a 3×2 grid (one per stage, excluding total)
    var STAGE_DEF = [
      { key: 'Awaiting Handoff',              label: 'Awaiting Handoff',       color: RED    },
      { key: 'Awaiting Kickoff',              label: 'Awaiting Kickoff',        color: YELLOW },
      { key: 'In Progress',                   label: 'In Progress',             color: GREEN  },
      { key: 'On Hold',                       label: 'On Hold',                 color: MUTED  },
      { key: 'New Logo Awaiting Go Live',     label: 'NL Awaiting Go Live',     color: BLUE   },
      { key: 'New Logo Live / Awaiting Closure', label: 'NL Awaiting Closure',  color: GREEN  },
    ];
    var lcw = Math.floor((CW - 2 * 8) / 3);   // column width
    var cellH = Math.floor((H - 116 - 6) / 2); // height per row (2 rows)
    var listTop2 = 114;
    STAGE_DEF.forEach(function(st, idx) {
      var col = idx % 3, row = Math.floor(idx / 3);
      var cx  = PAD + col * (lcw + 8);
      var cy  = listTop2 + row * (cellH + 6);
      var cnt = sCounts[st.key] || 0;
      var projs = sProjects[st.key] || [];
      var hdrColor = (st.color === MUTED) ? CARD2 : st.color;

      // Header
      rect(s2, cx, cy, lcw, 17, hdrColor);
      var hdrTxt = (st.color === MUTED) ? TEXT : '#FFFFFF';
      txt(s2, st.label + '  (' + cnt + ')', cx + 4, cy + 1, lcw - 8, 15, { size: 7.5, bold: true, color: hdrTxt, align: 'center' });

      var rowH2 = Math.min(22, Math.floor((cellH - 19) / 5));
      var py = cy + 19;
      projs.forEach(function(proj, pi) {
        if (py + rowH2 > cy + cellH) return;
        rect(s2, cx, py, lcw, rowH2, pi % 2 === 0 ? CARD : CARD2);
        var acct = (proj.account || '').substring(0, 24);
        var hrs  = proj.hrsRem > 0 ? Math.round(proj.hrsRem).toLocaleString() + ' hrs' : '';
        var acctTb = txt(s2, acct, cx + 4, py + 1, lcw - (hrs ? 52 : 8), rowH2 - 2, { size: 7.5, bold: true, vAlign: 'middle' });
        // Only apply hyperlink if account text is non-empty (getRange(0,0) on empty text throws)
        if (proj.sfId && acct.length > 0) {
          try {
            var sfLink = 'https://redis.lightning.force.com/lightning/r/project_cloud__Project__c/' + proj.sfId + '/view';
            acctTb.getText().getRange(0, acct.length).getTextStyle().setLinkUrl(sfLink);
          } catch(linkErr) {
            Logger.log('sfId link error (proj=' + proj.sfId + '): ' + linkErr);
          }
        }
        if (hrs) txt(s2, hrs, cx + lcw - 54, py + 1, 50, rowH2 - 2, { size: 7, color: BLUE, align: 'right', vAlign: 'middle' });
        py += rowH2;
      });
      if (projs.length === 0) {
        txt(s2, 'No projects', cx + 4, py, lcw - 8, 14, { size: 7, color: DIM, italic: true, align: 'center' });
      }
    });

    footer(s2);

    // ── SLIDE 3: PROJECTS BY TIER & OWNER ───────────────────
    _deckSection = 'slide3-tier-owner';
    var s3 = pres.appendSlide();
    setBg(s3);
    var to = p.tierOwner || {};
    titleBar(s3, 'Projects by Owner & Tier', p.generatedDate + '  •  Excludes Sophi Operations');
    var toHdrs = to.headers || [];
    var toRows = to.rows   || [];
    var centerC3 = [];
    for (var ci3 = 1; ci3 < toHdrs.length; ci3++) centerC3.push(ci3);
    makeTable(s3, 68, toHdrs, toRows, { maxRows: 28, centerCols: centerC3, fontSize: 9 });

    footer(s3);

    // ── SLIDE 4: GO LIVES (screenshot) ───────────────────────
    _deckSection = 'slide4-go-lives';
    var s4 = pres.appendSlide();
    setBg(s4);
    titleBar(s4, 'Go Live Summary', p.generatedDate);
    insertShot(s4, (p.screenshots || {}).slide4, 110, COMMENT_H);
    commentaryArea(s4, H - COMMENT_H);

    footer(s4);

    // ── SLIDE 5: ANNUAL PLAN & RE SUMMARY ───────────────────
    _deckSection = 'slide5-annual-plan';
    var s5 = pres.appendSlide();
    setBg(s5);
    var aps = p.annualPlanSummary || {};
    titleBar(s5, 'Annual Plan & RE Summary', 'Project count by region and utilization ratio');
    // Bigger, nicer table — leave room for commentary at bottom
    var apRows5 = aps.rows || [];
    var nc5 = (aps.headers || []).length;
    var nr5 = apRows5.length + 1;
    var maxTblH5 = H - 80 - COMMENT_H;
    var rh5 = Math.min(38, Math.floor(maxTblH5 / nr5));
    var th5 = rh5 * nr5;
    if (nr5 >= 2 && nc5 >= 2) {
      var tbl5 = s5.insertTable(nr5, nc5, PAD, 68, CW, th5);
      var hdrs5 = aps.headers || [];
      for (var c5 = 0; c5 < nc5; c5++) {
        var hc5 = tbl5.getCell(0, c5);
        hc5.getFill().setSolidFill(RED);
        hc5.getText().setText(String(hdrs5[c5] || ''));
        hc5.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor('#FFFFFF').setBold(true);
        if (hc5.getText().getParagraphs().length > 0)
          hc5.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(c5 > 0 ? SlidesApp.ParagraphAlignment.CENTER : SlidesApp.ParagraphAlignment.START);
      }
      apRows5.forEach(function(row, r5) {
        var bg5 = r5 === apRows5.length - 1 ? CARD2 : (r5 % 2 === 0 ? CARD : CARD2);
        for (var c5b = 0; c5b < nc5; c5b++) {
          var dc5 = tbl5.getCell(r5 + 1, c5b);
          dc5.getFill().setSolidFill(bg5);
          var val5 = row[c5b]; var str5 = (val5 === null || val5 === undefined) ? '—' : String(val5);
          dc5.getText().setText(str5);
          dc5.getText().getTextStyle().setFontFamily('Arial').setFontSize(c5b === 0 ? 12 : 14).setForegroundColor(TEXT);
          if (c5b > 0) dc5.getText().getTextStyle().setBold(true);
          if (dc5.getText().getParagraphs().length > 0)
            dc5.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(c5b > 0 ? SlidesApp.ParagraphAlignment.CENTER : SlidesApp.ParagraphAlignment.START);
          dc5.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        }
      });
    }

    commentaryArea(s5, H - COMMENT_H);

    footer(s5);

    // ── SLIDE 6: ANNUAL PLAN < 75% (screenshot) ─────────────
    _deckSection = 'slide6-under75';
    var s6 = pres.appendSlide();
    setBg(s6);
    titleBar(s6, 'Annual Plan & RE  —  Under 75% Utilization', p.generatedDate);
    insertShot(s6, (p.screenshots || {}).slide6, 68, COMMENT_H);
    commentaryArea(s6, H - COMMENT_H);

    footer(s6);

    // ── SLIDE 7: UTILIZATION SUMMARY (screenshot) ────────────
    _deckSection = 'slide7-utilization';
    var s7 = pres.appendSlide();
    setBg(s7);
    var utQ = p.utilization ? p.utilization.quarter : '';
    titleBar(s7, 'Utilization by Resource  —  Under 60% Billable', utQ || p.currentQuarter);

    // ── Regional summary scorecards ──────────────────────────
    var regSum = (p.utilization && p.utilization.regionSummary) ? p.utilization.regionSummary : [];
    var cardTop = 68;
    var cardH   = 62;
    var nReg    = regSum.length || 1;
    var cardW   = Math.floor((CW - (nReg - 1) * 6) / nReg);
    regSum.forEach(function(reg, ri) {
      var cx = PAD + ri * (cardW + 6);
      // Card background
      var card = s7.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cardTop, cardW, cardH);
      card.getFill().setSolidFill(CARD2);
      card.getBorder().setTransparent();
      // Accent bar on left edge
      var accent = s7.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cardTop, 3, cardH);
      accent.getFill().setSolidFill(RED);
      accent.getBorder().setTransparent();

      // Region name
      txt(s7, String(reg.region).toUpperCase(), cx + 8, cardTop + 4, cardW - 12, 10, { size: 6.5, bold: true, color: MUTED });
      // Count
      txt(s7, (reg.count || 0) + ' resources', cx + 8, cardTop + 14, cardW - 12, 9, { size: 6, color: DIM });

      // Two metric blocks side by side
      var metW = Math.floor((cardW - 20) / 2);
      var metrics = [
        { label: 'BILLABLE UTIL', val: reg.billable || '—' },
        { label: 'PTO ADJ UTIL',  val: reg.ptoAdj  || '—' },
      ];
      metrics.forEach(function(m, mi) {
        var mx = cx + 8 + mi * (metW + 4);
        var numVal = parseFloat(String(m.val).replace('%',''));
        var numColor = !isNaN(numVal) ? (numVal >= 75 ? GREEN : numVal >= 60 ? YELLOW : RED) : TEXT;
        txt(s7, m.label, mx, cardTop + 25, metW, 9, { size: 5.5, color: DIM });
        txt(s7, m.val,   mx, cardTop + 33, metW, 22, { size: 17, bold: true, color: numColor });
      });
    });

    var shotTop7 = cardTop + cardH + 12;
    var COMMENT_H7 = 62;  // slide 7: cards push content down, commentary slightly lower
    insertShot(s7, (p.screenshots || {}).slide7, shotTop7, COMMENT_H7);
    commentaryArea(s7, H - COMMENT_H7);

    footer(s7);

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
    ph.getText().getTextStyle().setFontFamily('Arial').setFontSize(14).setForegroundColor(DIM).setItalic(true);
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

/**
 * Writes KPI values to specific cells in the external tracking spreadsheet.
 * kpiArray: [{cell: 'A4', value: 0.72}, {cell: 'G3', value: 1500000}, ...]
 * Values that are null/undefined are skipped.
 * Called from the dashboard frontend via google.script.run.writeKPIsToSheet(...)
 */
function writeKPIsToSheet(kpiArray) {
  try {
    var TARGET_SS_ID = '1HNr9QtJ_De_ZLQcHx6-rDBgHze3P7InQfFVQGax8Vjg';
    var TARGET_GID   = 363748673;

    var ss = SpreadsheetApp.openById(TARGET_SS_ID);
    var targetSheet = null;
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === TARGET_GID) {
        targetSheet = sheets[i];
        break;
      }
    }
    if (!targetSheet) throw new Error('Target sheet not found (gid ' + TARGET_GID + ')');

    kpiArray.forEach(function(item) {
      if (item.value !== null && item.value !== undefined) {
        targetSheet.getRange(item.cell).setValue(item.value);
      }
    });

    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}
