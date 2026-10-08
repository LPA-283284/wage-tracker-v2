// hr-core.js — generated from Code_HR_v2.14.gs (UNCHANGED logic), run in the browser.
// Sheets calls are served by hr-shim.js from Supabase. Do not edit by hand.
var HRCORE = (function () {
// LPA Management — Apps Script backend (v2.14)
// v2.14 change: Advances and Staff Hours (pay-period weekly boxes) read/write
// Supabase (shared with the TeamOS app); the sheets are still written as backup.
// Touched: sbCash_ (error text only), submitAdvance, getPayPeriodReport (data
// source only — all calculations unchanged), setPayPeriodWeekValue(s).
// v2.13 change: Cash List reads/writes Supabase (shared with TeamOS); sheet kept as backup.
// This file should be named "Code.gs" in the Google Apps Script editor.
//
// v2.6 change (on top of v2.5): ONE new function appended at the end —
// applyGridEdits(edits) — a single-call batch writer for the Weekly Grid.
// v2.7 change: applyGridEdits made fast enough for big batches. v2.6 wrote
// every H/OFF/S mark through setHolidayMarkCore_, which re-reads the whole
// Holidays_Log per mark — 100+ marks blew Apps Script's execution limit.
// Marks are now applied with ONE Holidays_Log read and grouped writes, and
// Hours_Log row deletions are grouped into consecutive runs (deleteRows).
// Behaviour is IDENTICAL (same in-place update / delete / append semantics);
// failed items now also carry staffId so the client can match them. Nothing
// outside applyGridEdits was touched.
// v2.8 change: THREE new functions appended at the end for correcting tip
// pool entries — getTipPoolEntries (rows with references), updateTipPoolEntry
// and deleteTipPoolEntry (content-keyed: date + amount, row hint first, same
// safety pattern as the hours editors). Nothing existing was touched.
// v2.9 change: the pay-period report no longer prints Holiday in the Notes
// column (user request) — removed from exportPayPeriodPdf here and from
// renderPeriodSummary in the Index. Sick still shows. Holiday day COUNTING
// (entitlement, Holidays tab) is untouched.
// v2.10 change: ONE new function appended at the end —
// setPayPeriodWeekValues(year, month, edits) — a single-call batch writer
// for the Staff Hours weekly boxes. The screen used to fire one locked call
// per cell; fast typing queued 15-20 calls and the tail blew the 30s lock
// wait ("Lock timeout"). Now the client collects the edits and ONE call
// applies them all: validate-first (one bad value = nothing written), one
// sheet read, grouped writes, same text-format guard on PeriodKey. The old
// setPayPeriodWeekValue stays untouched.
// v2.11 change: ONE new function appended at the end —
// submitMissedBankingDay(entry) — fills ALL THREE ledgers a Streamlit
// banking day normally writes (Revenue_Log, Tip_Pool_Log cash tips,
// Cash_List Cash In Hand) in one locked call. Gap-filling and idempotent:
// each piece is written ONLY if that day doesn't already have it, and
// existing figures are reported back instead of ever being overwritten or
// doubled. Mirrors the backfill's conventions (Cash In Hand description).
// v2.12 change: Take-in entries can be corrected. getTakeinViewData rows now
// carry a row reference (additive field, nothing else in that function
// changed), and TWO new functions are appended — updateRevenueEntry and
// deleteRevenueEntry, content-keyed (date + both amounts, row hint first,
// same safety pattern as Hours/Tips). The DATE of an entry cannot be
// edited: date is the ledger's one-entry-per-day key — to move a day,
// delete the row and re-enter it with the Manual Entry card.
//
// v2.2 changes (on top of v2):
//  - Quarter-hour validation on ALL hour inputs: only .00/.25/.5/.75
//    fractions are accepted; .15/.30/.45 are rejected with a clear hint
//  - View Entries edit/delete now finds the row BY CONTENT, so deleting
//    several rows in a row works without reloading the list
//  - Tip distribution uses the tip weight in force ON EACH DAY for both
//    the person's share and the daily total (was inconsistent before)
//  - Rate changes are effective from TODAY INCLUSIVE (midnight), so a
//    raise entered at 2pm still applies to today's hours
//  - PayPeriod PeriodKey fix merged in (text format + tolerant read)
//  - editStaffBatch reports skipped rows instead of silently dropping them
//
// v2.3 changes:
//  - Penny-exact tip distribution (largest-remainder): daily distributed
//    total EXACTLY equals the daily pool; leftover pennies deterministic
//  - Unknown staff are EXCLUDED from tip distribution with a warning
//    (never a silent weight-1 guess); undistributed pool is reported
//  - Dates before the first Rate_History entry are costed with the
//    earliest known rate but flagged and reported (preHistoryHours) —
//    validated historical totals do NOT silently shift
//  - Central validators: tip weight 0–10, holiday entitlement 0–366
//    (addStaff / editStaff / editStaffBatch / recordRateHistory_)
//  - Day-of-week averages computed over the SAME day set (revenue days);
//    revenue days with no hours logged reported (revenueDaysMissingHours)
//  - Pure test suite: runAllTests() — never touches any sheet
//
// v2.4 changes:
//  - Monthly summary now carries total hours per month + yearly total
//    (monthly[].hours, totals.hours) — ALL logged hours, incl. unmatched
//  - Cash List edit/delete: updateCashEntry / deleteCashEntry with
//    content-based row matching (same safe pattern as Hours)
//  - getCashViewData rows now include a row reference for the UI
//
// v2.5 changes:
//  - COST EXCLUSIONS: a person's cost can be excluded for a given month
//    (Cost_Exclusions_Log: Month | Year | StaffID; single source of truth).
//    Affects ONLY £ figures (Dashboard KPIs, Monthly Summary cost/%/dept
//    breakdown). HOURS are NEVER affected anywhere — validated hour totals
//    (e.g. June 3079.75) stay identical. Exclusions are always surfaced
//    with a visible banner: an adjusted figure is never silently presented
//    as a raw one. Daily costs on Take-in and tip distribution unaffected.
//  - getCostExclusions / addCostExclusion / removeCostExclusion +
//    pure buildExclusionSet_ (covered by runAllTests).

var SHEET_STAFF = 'Staff';
var SHEET_HOURS_LOG = 'Hours_Log';
var SHEET_REVENUE_LOG = 'Revenue_Log';
var SHEET_TIPS_LOG = 'Tip_Pool_Log';
var SHEET_CASH_LIST = 'Cash_List_Log';
var SHEET_SETTINGS = 'Settings';
var SHEET_HOLIDAYS = 'Holidays_Log';
var SHEET_ADVANCES = 'Advances_Log';
var SHEET_PAYPERIOD_MANUAL = 'PayPeriod_Manual_Log';
var SHEET_RATE_HISTORY = 'Rate_History';
var SHEET_REQUESTS = 'Requests_Log';
var SHEET_COST_EXCLUSIONS = 'Cost_Exclusions_Log';

var MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
var DAY_NAMES = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LPA Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------
// Locking — every function that WRITES to a sheet runs inside this.
// ------------------------------------------------------------
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // wait up to 30s
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// Validation helpers (server-side; never trust the client)
// ------------------------------------------------------------
function assertDateStr_(str) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(str || ''))) {
    throw new Error('Invalid date: "' + str + '". Expected yyyy-mm-dd.');
  }
  return str;
}

function assertHours_(h) {
  var n = Number(h);
  if (isNaN(n) || n < 0 || n > 24) throw new Error('Invalid hours value: ' + h + ' (must be 0–24).');
  return n;
}

function assertAmount_(a, label) {
  var n = Number(a);
  if (isNaN(n) || n < 0 || n > 1000000) throw new Error('Invalid amount for ' + (label || 'field') + ': ' + a);
  return n;
}

// Central tip-weight validator: number, 0–10. Empty/undefined -> default 1.
function assertTipWeight_(w) {
  if (w === undefined || w === null || w === '') return 1;
  var n = Number(w);
  if (isNaN(n) || n < 0 || n > 10) throw new Error('Invalid tip weight: ' + w + ' (must be a number 0–10).');
  return n;
}

// Central holiday-entitlement validator: number, 0–366. Empty -> default 28.
function assertHolidayEntitlement_(h) {
  if (h === undefined || h === null || h === '') return 28;
  var n = Number(h);
  if (isNaN(n) || n < 0 || n > 366) throw new Error('Invalid holiday entitlement: ' + h + ' (must be a number 0–366).');
  return n;
}

// Pure: labor % with a zero-revenue guard (0, never NaN/Infinity).
function safePct_(cost, revenue) {
  return revenue > 0 ? cost / revenue : 0;
}

// Pure: days in a month (month is 1-based). Handles leap years.
function daysInMonth_(year, month) {
  return new Date(Number(year), Number(month), 0).getDate();
}

// Hours must be decimal with quarter fractions ONLY (.00/.25/.5/.75).
// Minute-style typos (.15/.30/.45) are rejected with a helpful hint so
// they can never be silently counted as decimal hours.
function validateQuarterHours_(raw) {
  var s = String(raw === undefined || raw === null ? '' : raw).trim().replace(',', '.');
  if (s === '') return NaN; // empty = no entry; callers skip NaN
  if (/^\d+\.$/.test(s)) s = s.slice(0, -1); // "10." is just 10
  if (!/^\d+(\.\d{1,2})?$/.test(s)) {
    throw new Error('Invalid hours: "' + raw + '". Enter decimal hours in quarter steps, e.g. 12, 12.25, 12.5, 12.75.');
  }
  var n = Number(s);
  var frac = Math.round((n - Math.floor(n)) * 100);
  if (frac !== 0 && frac !== 25 && frac !== 50 && frac !== 75) {
    var hint = '';
    if (frac === 15) hint = ' Did you mean 15 minutes? Enter .25 instead.';
    else if (frac === 30) hint = ' Did you mean 30 minutes? Enter .5 instead.';
    else if (frac === 45) hint = ' Did you mean 45 minutes? Enter .75 instead.';
    throw new Error('Invalid hours: "' + raw + '". Only .00, .25, .5 and .75 fractions are accepted.' + hint);
  }
  return n;
}

// PeriodKey normalizer: Sheets may auto-convert "2026-7" text to a Date;
// this makes both forms comparable.
function normPeriodKey_(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + (v.getMonth() + 1);
  return String(v);
}

// Today at 00:00 local — used as the effective date of rate changes so a
// raise entered at 2pm still covers today's hours (Hours_Log dates are
// midnight timestamps).
function todayMidnight_() {
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Converts a "yyyy-mm-dd" string to a local Date without timezone drift.
function parseDateStr(str) {
  assertDateStr_(str);
  var parts = String(str).split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function dayIndexMonFirst(date) {
  var jsDay = date.getDay(); // 0=Sun
  return (jsDay + 6) % 7; // 0=Mon
}

function dateKey_(d) {
  return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
}

// Display-only: 'yyyy-mm-dd' string -> 'dd-mm-yyyy' (for messages/warnings).
function dmyStr_(str) {
  var p = String(str || '').split('-');
  return (p.length === 3 && p[0].length === 4) ? p[2] + '-' + p[1] + '-' + p[0] : String(str);
}

function fmtDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Creates a sheet with headers if it doesn't exist yet.
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}

function getSettings() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  if (!sh) return { targetLaborPct: 0.34, year: new Date().getFullYear() };
  var data = sh.getDataRange().getValues();
  var settings = { targetLaborPct: 0.34, year: new Date().getFullYear() };
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][0] || '');
    if (label.indexOf('Target') !== -1) settings.targetLaborPct = Number(data[i][1]) || 0.34;
    if (label.indexOf('Year') !== -1) settings.year = Number(data[i][1]) || new Date().getFullYear();
  }
  return settings;
}

// ------------------------------------------------------------
// Caches. Staff (2 min) + per-year dashboard aggregates (10 min).
// ------------------------------------------------------------
function getStaffRawData_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('staff_raw_v2');
  if (cached) return JSON.parse(cached);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
  var data = sh.getDataRange().getValues();
  try { cache.put('staff_raw_v2', JSON.stringify(data), 120); } catch (e) {}
  return data;
}

function invalidateStaffCache_() {
  CacheService.getScriptCache().remove('staff_raw_v2');
  invalidateDataCaches_();
}

function invalidateDataCaches_() {
  var cache = CacheService.getScriptCache();
  var y = new Date().getFullYear();
  var keys = [];
  for (var i = y - 6; i <= y + 1; i++) keys.push('dash_v2_' + i);
  cache.removeAll(keys);
}

// ------------------------------------------------------------
// Staff
// Columns: A=ID, B=Name, C=Department, D=Hourly Rate, E=Active,
//          F=Tip Weight, G=Order, H=Holiday Entitlement, I=Start, J=End
// ------------------------------------------------------------
function getStaffList(activeOnly) {
  if (activeOnly === undefined) activeOnly = true;
  var data = getStaffRawData_();
  var staff = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      var active = (data[i][4] === undefined || data[i][4] === '' || data[i][4] === true);
      if (activeOnly && !active) continue;
      var tipWeight = (data[i][5] === undefined || data[i][5] === '') ? 1 : Number(data[i][5]);
      var order = (data[i][6] === undefined || data[i][6] === '' || isNaN(Number(data[i][6]))) ? (i + 1000) : Number(data[i][6]);
      var holidayEntitlement = (data[i][7] === undefined || data[i][7] === '' || isNaN(Number(data[i][7]))) ? 28 : Number(data[i][7]);
      var startDate = data[i][8] ? fmtDate_(new Date(data[i][8])) : '';
      var endDate = data[i][9] ? fmtDate_(new Date(data[i][9])) : '';
      staff.push({
        id: String(data[i][0]),
        name: String(data[i][1]),
        dept: String(data[i][2]),
        rate: Number(data[i][3]) || 0,
        active: active,
        tipWeight: tipWeight,
        order: order,
        holidayEntitlement: holidayEntitlement,
        startDate: startDate,
        endDate: endDate
      });
    }
  }
  staff.sort(function (a, b) {
    if (a.active !== b.active) return a.active ? -1 : 1;
    return a.order - b.order;
  });
  return staff;
}

// FRESH read (never the cache) — used to find a row before writing to it.
function findStaffRowFresh_(sh, id) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ------------------------------------------------------------
// Rate_History
// Columns: A=StaffID, B=Effective Date, C=Hourly Rate, D=Tip Weight
// ------------------------------------------------------------
function rateHistorySheet_() {
  return ensureSheet_(SHEET_RATE_HISTORY, ['StaffID', 'Effective Date', 'Hourly Rate', 'Tip Weight']);
}

function recordRateHistory_(staffId, effectiveDate, rate, tipWeight) {
  var sh = rateHistorySheet_();
  sh.appendRow([String(staffId), effectiveDate, assertAmount_(rate || 0, 'hourly rate'), assertTipWeight_(tipWeight)]);
  invalidateDataCaches_();
}

// { staffId: [{time, rate, weight}, ...] } sorted by date ascending
function getRateHistoryMap_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RATE_HISTORY);
  var map = {};
  if (!sh) return map;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] || !data[i][1]) continue;
    var id = String(data[i][0]);
    if (!map[id]) map[id] = [];
    map[id].push({
      time: new Date(data[i][1]).getTime(),
      rate: Number(data[i][2]) || 0,
      weight: (data[i][3] === '' || data[i][3] === undefined || isNaN(Number(data[i][3]))) ? 1 : Number(data[i][3])
    });
  }
  Object.keys(map).forEach(function (id) {
    map[id].sort(function (a, b) { return a.time - b.time; });
  });
  return map;
}

// Rate + tip weight in force for staffId on `date`.
// preHistory:true = the date is BEFORE the first Rate_History entry (or no
// history exists at all). We still cost using the earliest known rate so
// validated historical totals don't silently shift, but callers surface
// these hours as an explicit warning instead of hiding the approximation.
function rateForDate_(histMap, staffById, staffId, date) {
  var t = date.getTime();
  var hist = histMap[staffId];
  if (hist && hist.length) {
    var found = null;
    for (var i = 0; i < hist.length; i++) {
      if (hist[i].time <= t) found = hist[i]; else break;
    }
    if (found) return { rate: found.rate, weight: found.weight, preHistory: false };
    return { rate: hist[0].rate, weight: hist[0].weight, preHistory: true };
  }
  var s = staffById[staffId];
  if (s) return { rate: s.rate, weight: s.tipWeight, preHistory: true };
  return null; // unknown staff entirely
}

// ONE-TIME: seeds Rate_History from the current Staff sheet.
function initializeRateHistory() {
  return withLock_(function () {
    var sh = rateHistorySheet_();
    var existing = {};
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) existing[String(data[i][0])] = true;
    }
    var staff = getStaffList(false);
    var added = 0;
    var rows = [];
    staff.forEach(function (s) {
      if (existing[s.id]) return;
      var eff = s.startDate ? parseDateStr(s.startDate) : new Date(2000, 0, 1);
      rows.push([s.id, eff, s.rate, s.tipWeight]);
      added++;
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    invalidateDataCaches_();
    var msg = added + ' staff members seeded into Rate_History (' + (staff.length - added) + ' already had history).';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return { added: added };
  });
}

// ---- Add staff ----
function addStaff(name, dept, rate, tipWeight, holidayEntitlement, startDateStr) {
  return withLock_(function () {
    name = String(name || '').trim();
    dept = String(dept || '').trim();
    rate = assertAmount_(rate || 0, 'hourly rate');
    tipWeight = assertTipWeight_(tipWeight);
    holidayEntitlement = assertHolidayEntitlement_(holidayEntitlement);
    if (!name || !dept) throw new Error('Name and department are required.');

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var data = sh.getDataRange().getValues();
    var maxId = 0, maxOrder = 0;
    for (var i = 1; i < data.length; i++) {
      var idNum = Number(data[i][0]);
      if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
      var o = Number(data[i][6]);
      if (!isNaN(o) && o > maxOrder) maxOrder = o;
    }
    var newId = maxId + 1;
    var startDate = startDateStr ? parseDateStr(startDateStr) : '';
    sh.appendRow([newId, name, dept, rate, true, tipWeight, maxOrder + 1, holidayEntitlement, startDate, '']);
    recordRateHistory_(newId, startDate || todayMidnight_(), rate, tipWeight);
    invalidateStaffCache_();
    return { ok: true, id: newId };
  });
}

function setStaffActive(id, active) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var row = findStaffRowFresh_(sh, id);
    if (row === -1) throw new Error('Staff member not found (ID: ' + id + ')');
    sh.getRange(row, 5).setValue(active);
    invalidateStaffCache_();

    var others = getStaffList(false).filter(function (s) { return s.id !== String(id); });
    var sameGroup = others.filter(function (s) { return s.active === active; });
    var maxOrder = 0;
    sameGroup.forEach(function (s) { if (s.order > maxOrder) maxOrder = s.order; });
    sh.getRange(row, 7).setValue(maxOrder + 1);
    invalidateStaffCache_();
    return { ok: true };
  });
}

// Permanent delete. Historical costing keeps working via Rate_History.
function deleteStaffPermanently(id) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var row = findStaffRowFresh_(sh, id);
    if (row === -1) throw new Error('Staff member not found (ID: ' + id + ')');
    sh.deleteRow(row);
    invalidateStaffCache_();
    return { ok: true };
  });
}

// ---- Edit one staff member ----
// v2.2: rate/weight changes are recorded effective from TODAY MIDNIGHT so
// today's hours (logged at 00:00) are covered by the new rate.
function editStaff(id, newName, newDept, rate, tipWeight, holidayEntitlement, startDateStr, endDateStr) {
  return withLock_(function () {
    newName = String(newName || '').trim();
    newDept = String(newDept || '').trim();
    rate = assertAmount_(rate || 0, 'hourly rate');
    tipWeight = assertTipWeight_(tipWeight);
    holidayEntitlement = assertHolidayEntitlement_(holidayEntitlement);
    if (!newName || !newDept) throw new Error('Name and department are required.');

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var row = findStaffRowFresh_(sh, id);
    if (row === -1) throw new Error('Staff member not found (ID: ' + id + ')');

    var current = sh.getRange(row, 1, 1, 10).getValues()[0];
    var oldRate = Number(current[3]) || 0;
    var oldWeight = (current[5] === '' || current[5] === undefined) ? 1 : Number(current[5]);

    sh.getRange(row, 2, 1, 3).setValues([[newName, newDept, rate]]);
    sh.getRange(row, 6).setValue(tipWeight);
    sh.getRange(row, 8).setValue(holidayEntitlement);
    if (startDateStr !== undefined && startDateStr !== null) {
      sh.getRange(row, 9).setValue(startDateStr ? parseDateStr(startDateStr) : '');
    }
    if (endDateStr !== undefined && endDateStr !== null) {
      sh.getRange(row, 10).setValue(endDateStr ? parseDateStr(endDateStr) : '');
    }

    if (oldRate !== rate || oldWeight !== tipWeight) {
      recordRateHistory_(id, todayMidnight_(), rate, tipWeight); // effective from today INCLUSIVE
    }
    invalidateStaffCache_();
    return { ok: true };
  });
}

// ---- Save All ----
// v2.2: invalid rows are reported back instead of silently skipped.
function editStaffBatch(updates) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var data = sh.getDataRange().getValues();
    var idToRow = {};
    for (var i = 1; i < data.length; i++) idToRow[String(data[i][0])] = i + 1;

    var savedCount = 0;
    var skipped = [];
    var historyRows = [];

    updates.forEach(function (u) {
      var row = idToRow[String(u.id)];
      if (!row) { skipped.push('ID ' + u.id + ' (not found)'); return; }
      var newName = String(u.name || '').trim();
      var newDept = String(u.dept || '').trim();
      if (!newName || !newDept) { skipped.push('ID ' + u.id + ' (missing name/dept)'); return; }

      var rate, tipWeight, holidayEntitlement;
      try {
        rate = assertAmount_(u.rate || 0, 'hourly rate');
        tipWeight = assertTipWeight_(u.tipWeight);
        holidayEntitlement = assertHolidayEntitlement_(u.holidayEntitlement);
      } catch (err) {
        skipped.push(newName + ' (' + err.message + ')');
        return; // this row is skipped and reported; the rest of the batch continues
      }

      var oldRate = Number(data[row - 1][3]) || 0;
      var oldWeight = (data[row - 1][5] === '' || data[row - 1][5] === undefined) ? 1 : Number(data[row - 1][5]);

      sh.getRange(row, 2, 1, 3).setValues([[newName, newDept, rate]]);
      sh.getRange(row, 6).setValue(tipWeight);
      sh.getRange(row, 8).setValue(holidayEntitlement);
      if (u.startDate !== undefined && u.startDate !== null) {
        sh.getRange(row, 9).setValue(u.startDate ? parseDateStr(u.startDate) : '');
      }
      if (u.endDate !== undefined && u.endDate !== null) {
        sh.getRange(row, 10).setValue(u.endDate ? parseDateStr(u.endDate) : '');
      }

      if (oldRate !== rate || oldWeight !== tipWeight) {
        historyRows.push([String(u.id), todayMidnight_(), rate, tipWeight]);
      }
      savedCount++;
    });

    if (historyRows.length) {
      var hs = rateHistorySheet_();
      hs.getRange(hs.getLastRow() + 1, 1, historyRows.length, 4).setValues(historyRows);
    }
    invalidateStaffCache_();
    return { ok: true, count: savedCount, skipped: skipped };
  });
}

function setStaffPosition(id, position) {
  return withLock_(function () {
    var list = getStaffList(false);
    var idx = -1;
    for (var i = 0; i < list.length; i++) if (list[i].id === String(id)) { idx = i; break; }
    if (idx === -1) throw new Error('Staff member not found (ID: ' + id + ')');

    var isActive = list[idx].active;
    var group = list.filter(function (s) { return s.active === isActive; });
    var groupIdx = -1;
    for (var g = 0; g < group.length; g++) if (group[g].id === String(id)) { groupIdx = g; break; }

    var target = Math.max(1, Math.min(Number(position) || 1, group.length));
    var item = group.splice(groupIdx, 1)[0];
    group.splice(target - 1, 0, item);

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    group.forEach(function (s, newIdx) {
      var row = findStaffRowFresh_(sh, s.id);
      if (row !== -1) sh.getRange(row, 7).setValue(newIdx + 1);
    });
    invalidateStaffCache_();
    return { ok: true };
  });
}

function moveStaff(id, direction) {
  return withLock_(function () {
    var list = getStaffList(false);
    var idx = -1;
    for (var i = 0; i < list.length; i++) if (list[i].id === String(id)) { idx = i; break; }
    if (idx === -1) throw new Error('Staff member not found (ID: ' + id + ')');

    var swapIdx = idx + direction;
    if (swapIdx < 0 || swapIdx >= list.length) return { ok: true };
    if (list[idx].active !== list[swapIdx].active) return { ok: true };

    var a = list[idx], b = list[swapIdx];
    var sh2 = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var rowA = findStaffRowFresh_(sh2, a.id);
    var rowB = findStaffRowFresh_(sh2, b.id);
    sh2.getRange(rowA, 7).setValue(b.order);
    sh2.getRange(rowB, 7).setValue(a.order);
    invalidateStaffCache_();
    return { ok: true };
  });
}

// ==============================================================
// ONE-TIME MAINTENANCE TOOLS
// ==============================================================
function repairStaffOrder() {
  return withLock_(function () {
    invalidateStaffCache_();
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var list = getStaffList(false);
    var activeCounter = 0, inactiveCounter = 0;
    list.forEach(function (s) {
      var newOrder = s.active ? (++activeCounter) : (++inactiveCounter);
      var row = findStaffRowFresh_(sh, s.id);
      if (row !== -1) sh.getRange(row, 7).setValue(newOrder);
    });
    invalidateStaffCache_();
    var msg = list.length + ' staff members re-numbered (Active: 1-' + activeCounter + ', Inactive: 1-' + inactiveCounter + ').';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  });
}

function repairStaffIds() {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STAFF);
    var data = sh.getDataRange().getValues();
    var usedIds = {};
    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      var idVal = String(data[i][0]).trim();
      var idNum = Number(idVal);
      if (idVal !== '' && !isNaN(idNum) && !usedIds[idVal]) {
        usedIds[idVal] = true;
        if (idNum > maxId) maxId = idNum;
      }
    }
    var fixedCount = 0;
    for (var r = 1; r < data.length; r++) {
      if (!data[r][1]) continue;
      var idVal2 = String(data[r][0]).trim();
      var idNum2 = Number(idVal2);
      var isValid = (idVal2 !== '' && !isNaN(idNum2));
      var isDuplicate = isValid && usedIds[idVal2] === 'seen';
      if (!isValid || isDuplicate) {
        maxId += 1;
        sh.getRange(r + 1, 1).setValue(maxId);
        fixedCount++;
      } else if (isValid) {
        usedIds[idVal2] = 'seen';
      }
    }
    invalidateStaffCache_();
    var msg = fixedCount + ' row IDs fixed (blank or duplicate ones).';
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  });
}

// ------------------------------------------------------------
// Hours entry
// Hours_Log columns: A=Date, B=Department, C=Staff Name, D=Hours, E=StaffID
// ------------------------------------------------------------
function findHoursConflicts_(pairs) {
  // pairs: [{id, name, dateKey}] ; returns [{name, date}]
  var wanted = {};
  pairs.forEach(function (p) { wanted[p.id + '|' + p.dateKey] = p; });
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var data = sh.getDataRange().getValues();
  var conflicts = [];
  var seen = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var key = String(data[i][4]) + '|' + dateKey_(new Date(data[i][0]));
    if (wanted[key] && !seen[key]) {
      seen[key] = true;
      conflicts.push({ name: wanted[key].name, date: fmtDate_(new Date(data[i][0])) });
    }
  }
  return conflicts;
}

// entries: [{id, name, dept, hours: [mon..sun]}, ...]
function submitWeeklyHours(weekStartStr, entries, force) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
    var weekStart = parseDateStr(weekStartStr);
    var rows = [];
    var pairs = [];

    // TWO-PASS for atomicity (see submitDailyHours): validate everything
    // first; one bad cell means NOTHING is written.
    var markOps = [];
    entries.forEach(function (e) {
      for (var d = 0; d < 7; d++) {
        var raw = String(e.hours[d] === undefined || e.hours[d] === null ? '' : e.hours[d]).trim().toUpperCase();
        var mDate = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + d);
        if (raw === 'O') raw = 'OFF';
        if (raw === 'H' || raw === 'S' || raw === 'OFF') {
          markOps.push({ date: fmtDate_(mDate), id: e.id, type: raw });
          continue;
        }
        var h = validateQuarterHours_(e.hours[d]); // throws -> nothing written
        if (!isNaN(h) && h > 0) {
          assertHours_(h);
          rows.push([mDate, e.dept, e.name, h, e.id]);
          pairs.push({ id: String(e.id), name: e.name, dateKey: dateKey_(mDate) });
        }
      }
    });

    if (force !== true && rows.length > 0) {
      var conflicts = findHoursConflicts_(pairs);
      if (conflicts.length > 0) return { added: 0, marks: 0, conflicts: conflicts };
    }

    markOps.forEach(function (m) { setHolidayMarkCore_(m.date, m.id, m.type); });
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }
    if (markOps.length || rows.length) invalidateDataCaches_();
    return { added: rows.length, marks: markOps.length, conflicts: [] };
  });
}

// entries: [{id, name, dept, hours}, ...]
function submitDailyHours(dateStr, entries, force) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
    var date = parseDateStr(dateStr);
    var rows = [];
    var pairs = [];

    // TWO-PASS for atomicity: first validate EVERYTHING and collect the
    // work; only when the whole submission is valid does anything get
    // written. One bad cell = nothing saved, never a half-saved day.
    var markOps = [];
    entries.forEach(function (e) {
      var raw = String(e.hours === undefined || e.hours === null ? '' : e.hours).trim().toUpperCase();
      if (raw === 'O') raw = 'OFF';
      if (raw === 'H' || raw === 'S' || raw === 'OFF') {
        // Absence marks go to Holidays_Log (single source of truth) —
        // they show in the Holidays tab, count toward entitlement and
        // flow into the Period Summary notes. NOT an Hours_Log row.
        markOps.push({ date: dateStr, id: e.id, type: raw });
        return;
      }
      var h = validateQuarterHours_(e.hours); // throws on bad input -> nothing written
      if (!isNaN(h) && h > 0) {
        assertHours_(h);
        rows.push([date, e.dept, e.name, h, e.id]);
        pairs.push({ id: String(e.id), name: e.name, dateKey: dateKey_(date) });
      }
    });

    if (force !== true && rows.length > 0) {
      var conflicts = findHoursConflicts_(pairs);
      if (conflicts.length > 0) return { added: 0, marks: 0, conflicts: conflicts };
    }

    markOps.forEach(function (m) { setHolidayMarkCore_(m.date, m.id, m.type); });
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }
    if (markOps.length || rows.length) invalidateDataCaches_();
    return { added: rows.length, marks: markOps.length, conflicts: [] };
  });
}

// ------------------------------------------------------------
// Hours — View / Edit / Delete logged entries
// v2.2: the target row is found BY CONTENT (date + staff + hours), trying
// the remembered row number first. Deleting several entries in a row now
// works even though row numbers shift after each delete.
// ------------------------------------------------------------
function getHoursEntries(fromStr, toStr) {
  var from = parseDateStr(fromStr);
  var to = parseDateStr(toStr);
  var toEnd = new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = new Date(data[i][0]);
    if (d >= from && d < toEnd) {
      out.push({
        row: i + 1,
        date: fmtDate_(d),
        dept: String(data[i][1] || ''),
        name: String(data[i][2] || ''),
        hours: Number(data[i][3]) || 0,
        staffId: String(data[i][4] || '')
      });
    }
  }
  out.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.name < b.name ? -1 : 1;
  });
  return out;
}

function findHoursRowByContent_(sh, expected) {
  var data = sh.getDataRange().getValues();
  function matches(i) { // i = 0-based data index
    var vals = data[i];
    if (!vals || !vals[0]) return false;
    return fmtDate_(new Date(vals[0])) === expected.date &&
      String(vals[4]) === String(expected.staffId) &&
      Math.abs((Number(vals[3]) || 0) - Number(expected.hours)) < 0.001;
  }
  var hinted = Number(expected.row) - 1;
  if (hinted >= 1 && hinted < data.length && matches(hinted)) return hinted + 1;
  for (var i = 1; i < data.length; i++) {
    if (matches(i)) return i + 1;
  }
  return -1;
}

function updateHoursEntry(expected, newHours) {
  return withLock_(function () {
    newHours = validateQuarterHours_(newHours);
    assertHours_(newHours);
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
    var row = findHoursRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Entry not found (' + dmyStr_(expected.date) + ', ' + expected.hours + 'h). It may have been edited or deleted already — please reload the list.');
    }
    sh.getRange(row, 4).setValue(Number(newHours));
    invalidateDataCaches_();
    return { ok: true };
  });
}

function deleteHoursEntry(expected) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
    var row = findHoursRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Entry not found (' + dmyStr_(expected.date) + ', ' + expected.hours + 'h). It may have been deleted already — please reload the list.');
    }
    sh.deleteRow(row);
    invalidateDataCaches_();
    return { ok: true };
  });
}

// Daily / weekly / monthly hours summary per staff member
function getSummary(dateStr, weekStartStr, month, year) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var data = sh.getDataRange().getValues();
  var staff = getStaffList(false);

  var targetDate = parseDateStr(dateStr);
  var weekStart = parseDateStr(weekStartStr);
  var weekEnd = new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + 7);

  var result = staff.map(function (s) {
    return { id: s.id, name: s.name, dept: s.dept, active: s.active, daily: 0, weekly: 0, monthly: 0 };
  });
  var index = {};
  result.forEach(function (r, i) { index[r.id] = i; });

  // Hours belonging to IDs NOT in the Staff list (deleted / import
  // leftovers) are NOT silently dropped: they are collected into one
  // visible row so this report always reconciles with Monthly Summary.
  var unknown = { daily: 0, weekly: 0, monthly: 0 };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var d = new Date(row[0]);
    var staffId = String(row[4]);
    var hours = parseFloat(row[3]) || 0;
    var match = (staffId in index) ? result[index[staffId]] : unknown;
    if (sameDay(d, targetDate)) match.daily += hours;
    if (d >= weekStart && d < weekEnd) match.weekly += hours;
    if ((d.getMonth() + 1) === Number(month) && d.getFullYear() === Number(year)) match.monthly += hours;
  }
  var out = result.filter(function (r) {
    return r.active || r.daily > 0 || r.weekly > 0 || r.monthly > 0;
  });
  if (unknown.daily > 0 || unknown.weekly > 0 || unknown.monthly > 0) {
    out.push({
      id: '', name: '(Not in staff list / deleted)', dept: '', active: true,
      daily: unknown.daily, weekly: unknown.weekly, monthly: unknown.monthly
    });
  }
  return out;
}

// ------------------------------------------------------------
// Revenue entry — one entry per date.
// ------------------------------------------------------------
function submitRevenue(entry) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
    var date = parseDateStr(entry.date);
    var net = assertAmount_(entry.netSales || 0, 'net sales');
    var sc = assertAmount_(entry.serviceCharge || 0, 'service charge');

    var data = sh.getDataRange().getValues();
    var key = dateKey_(date);
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && dateKey_(new Date(data[i][0])) === key) {
        throw new Error('A revenue entry already exists for ' + dmyStr_(entry.date) + ' (' +
          '£' + (Number(data[i][1]) || 0).toFixed(2) + ' + SC £' + (Number(data[i][2]) || 0).toFixed(2) +
          '). To correct it, edit or delete that row in the Revenue_Log sheet.');
      }
    }
    sh.appendRow([date, net, sc, entry.notes || '']);
    invalidateDataCaches_();
    return { ok: true };
  });
}

// Take-in view: recent entries WITH that day's staff cost (uses Rate_History)
function getTakeinViewData(n) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      row: i + 1, // v2.12: sheet row reference for the edit/delete UI
      date: fmtDate_(new Date(data[i][0])),
      dateObj: new Date(data[i][0]),
      netSales: Number(data[i][1]) || 0,
      serviceCharge: Number(data[i][2]) || 0,
      notes: data[i][3] || ''
    });
  }
  rows.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  rows = rows.slice(0, n || 30);

  var wantedDates = {};
  rows.forEach(function (r) { wantedDates[dateKey_(r.dateObj)] = 0; });

  var histMap = getRateHistoryMap_();
  var staffById = {};
  getStaffList(false).forEach(function (s) { staffById[s.id] = s; });

  var shH = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var hData = shH.getDataRange().getValues();
  for (var j = 1; j < hData.length; j++) {
    if (!hData[j][0]) continue;
    var d = new Date(hData[j][0]);
    var k = dateKey_(d);
    if (!(k in wantedDates)) continue;
    var rw = rateForDate_(histMap, staffById, String(hData[j][4]), d);
    if (rw) wantedDates[k] += (Number(hData[j][3]) || 0) * rw.rate;
  }

  return rows.map(function (r) {
    var cost = wantedDates[dateKey_(r.dateObj)] || 0;
    var total = r.netSales + r.serviceCharge;
    return {
      row: r.row, // v2.12: carried through for the edit/delete UI
      date: r.date, netSales: r.netSales, serviceCharge: r.serviceCharge,
      total: total, staffCost: cost,
      laborPct: safePct_(cost, total),
      notes: r.notes
    };
  });
}

// ------------------------------------------------------------
// Tip pool entry
// ------------------------------------------------------------
function submitTipPool(entry) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
    var date = parseDateStr(entry.date);
    var amount = assertAmount_(entry.amount || 0, 'tips');
    sh.appendRow([date, amount, entry.notes || '']);
    invalidateDataCaches_();
    return { ok: true };
  });
}

function getRecentTipPool(n) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      date: fmtDate_(new Date(data[i][0])),
      amount: Number(data[i][1]) || 0,
      notes: data[i][2] || ''
    });
  }
  rows.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  return rows.slice(0, n || 30);
}

// ------------------------------------------------------------
// Cash List
// ------------------------------------------------------------
// v2.13: Cash List now reads/writes Supabase (LPA HR) — the single source
// shared with TeamOS. The sheet is still written as a backup copy.
// Script Properties needed: SUPABASE_URL, SUPABASE_KEY, CASH_SYNC_SECRET.
function sbCash_(fnName, payload) {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('SUPABASE_URL');
  var key = props.getProperty('SUPABASE_KEY');
  var secret = props.getProperty('CASH_SYNC_SECRET');
  if (!url || !key || !secret) throw new Error('Supabase cash sync is not configured (Script Properties).');
  payload.p_secret = secret;
  var res = UrlFetchApp.fetch(url + '/rest/v1/rpc/' + fnName, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { apikey: key, Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload)
  });
  var body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() !== 200 || !body || body.ok !== true) {
    throw new Error('Supabase error: ' + ((body && (body.error || body.message)) || res.getResponseCode()));
  }
  return body;
}

function submitCashEntry(entry) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH_LIST);
    var date = parseDateStr(entry.date);
    var amount = assertAmount_(entry.amount || 0, 'amount');
    var income = (entry.type === 'income') ? amount : 0;
    var expense = (entry.type === 'expense') ? amount : 0;
    if (!String(entry.description || '').trim()) throw new Error('Description is required.');
    // Supabase first (source of truth); the sheet is only written if that succeeded
    sbCash_('hr_cash_write', { p_op: 'insert', p_old: null, p_new: {
      date: entry.date, description: String(entry.description).trim(),
      income: income, expense: expense, source: 'MANUAL', notes: entry.notes || '' } });
    sh.appendRow([date, entry.description, income, expense, 'Manual', entry.notes || '']);
    return { ok: true };
  });
}

function getCashViewData(month, year) {
  var monthStart = new Date(Number(year), Number(month) - 1, 1);
  var monthEnd = new Date(Number(year), Number(month), 1);

  var all = sbCash_('hr_cash_list', {}).rows.map(function (r) {
    return {
      row: r.id, // Supabase id, used as a HINT by edit/delete (content re-verified)
      date: parseDateStr(r.date),
      description: r.description || '',
      income: Number(r.income) || 0,
      expense: Number(r.expense) || 0,
      source: r.source || '',
      notes: r.notes || ''
    };
  });
  all.sort(function (a, b) { return a.date - b.date || a.row - b.row; });

  var carryForward = 0;
  var totalBalance = 0;
  var monthEntries = [];
  all.forEach(function (e) {
    var delta = e.income - e.expense;
    totalBalance += delta;
    if (e.date < monthStart) carryForward += delta;
    else if (e.date < monthEnd) monthEntries.push(e);
  });

  var balance = carryForward;
  var rows = monthEntries.map(function (e) {
    balance += (e.income - e.expense);
    return {
      row: e.row,
      date: fmtDate_(e.date),
      description: e.description,
      income: e.income,
      expense: e.expense,
      balance: balance,
      source: e.source,
      notes: e.notes
    };
  });

  return { carryForward: carryForward, rows: rows, endingBalance: balance, currentBalance: totalBalance };
}

// Locates the Cash_List SHEET row matching expected {date, description,
// income, expense} by content (the row hint is now a Supabase id, so the
// sheet is always scanned by content). Used only for the backup copy.
function findCashRowByContent_(sh, expected) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var vals = data[i];
    if (!vals || !vals[0]) continue;
    if (fmtDate_(new Date(vals[0])) === expected.date &&
      String(vals[1] || '') === String(expected.description || '') &&
      Math.abs((Number(vals[2]) || 0) - (Number(expected.income) || 0)) < 0.001 &&
      Math.abs((Number(vals[3]) || 0) - (Number(expected.expense) || 0)) < 0.001) return i + 1;
  }
  return -1;
}

// newEntry: {date, description, type: 'income'|'expense', amount, notes}
function updateCashEntry(expected, newEntry) {
  return withLock_(function () {
    var date = parseDateStr(newEntry.date);
    var amount = assertAmount_(newEntry.amount || 0, 'amount');
    var income = (newEntry.type === 'income') ? amount : 0;
    var expense = (newEntry.type === 'expense') ? amount : 0;
    if (!String(newEntry.description || '').trim()) throw new Error('Description is required.');
    try {
      sbCash_('hr_cash_write', { p_op: 'update',
        p_old: { id: expected.row, date: expected.date, description: expected.description, income: expected.income, expense: expected.expense },
        p_new: { date: newEntry.date, description: String(newEntry.description).trim(), income: income, expense: expense, notes: newEntry.notes || '' } });
    } catch (e) {
      throw new Error('Cash entry not found (' + dmyStr_(expected.date) + ', "' + expected.description + '"). It may have been edited or deleted already — please reload the list.');
    }
    // backup copy in the sheet (rows added from TeamOS are not in the sheet — that's fine)
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH_LIST);
    var row = findCashRowByContent_(sh, expected);
    if (row !== -1) {
      sh.getRange(row, 1, 1, 4).setValues([[date, String(newEntry.description).trim(), income, expense]]);
      sh.getRange(row, 6).setValue(newEntry.notes || '');
    }
    return { ok: true };
  });
}

function deleteCashEntry(expected) {
  return withLock_(function () {
    try {
      sbCash_('hr_cash_write', { p_op: 'delete', p_new: null,
        p_old: { id: expected.row, date: expected.date, description: expected.description, income: expected.income, expense: expected.expense } });
    } catch (e) {
      throw new Error('Cash entry not found (' + dmyStr_(expected.date) + ', "' + expected.description + '"). It may have been deleted already — please reload the list.');
    }
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH_LIST);
    var row = findCashRowByContent_(sh, expected);
    if (row !== -1) sh.deleteRow(row);
    return { ok: true };
  });
}

// ------------------------------------------------------------
// Holidays
// ------------------------------------------------------------
// Core WITHOUT the lock, so it can be called from inside other locked
// writes (e.g. hours submission). LockService locks are NOT re-entrant;
// nesting withLock_ would deadlock.
function setHolidayMarkCore_(dateStr, staffId, type) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOLIDAYS);
  var data = sh.getDataRange().getValues();
  var date = parseDateStr(dateStr);
  var targetKey = dateKey_(date);

  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (dateKey_(new Date(data[i][0])) === targetKey && String(data[i][1]) === String(staffId)) {
      foundRow = i + 1;
      break;
    }
  }

  if (!type) {
    if (foundRow !== -1) sh.deleteRow(foundRow);
    return { ok: true };
  }
  if (['H', 'OFF', 'S'].indexOf(String(type)) === -1) throw new Error('Invalid mark: ' + type);

  if (foundRow !== -1) sh.getRange(foundRow, 3).setValue(type);
  else sh.appendRow([date, staffId, type, '']);
  return { ok: true };
}

function setHolidayMark(dateStr, staffId, type) {
  return withLock_(function () {
    return setHolidayMarkCore_(dateStr, staffId, type);
  });
}

// One combined call for the whole Holidays tab.
function getHolidaysViewData(month, year) {
  month = Number(month); year = Number(year);
  var staffAll = getStaffList(false);
  var staffActive = staffAll.filter(function (s) { return s.active; });
  var staffById = {};
  staffAll.forEach(function (s) { staffById[s.id] = s; });

  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOLIDAYS);
  var data = sh.getDataRange().getValues();

  var daysInMonth = daysInMonth_(year, month); // leap-year safe
  var marks = {};
  var yearCounts = {};
  staffAll.forEach(function (s) {
    yearCounts[s.id] = { holidayDays: 0, offDays: 0, sickDays: 0 };
  });
  var today = new Date();
  var todayRows = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = new Date(data[i][0]);
    var staffId = String(data[i][1]);
    var type = String(data[i][2]);

    if (d.getFullYear() === year && (d.getMonth() + 1) === month && staffById[staffId] && staffById[staffId].active) {
      marks[staffId + '|' + d.getDate()] = type;
    }
    if (d.getFullYear() === year && yearCounts[staffId]) {
      if (type === 'H') yearCounts[staffId].holidayDays++;
      else if (type === 'OFF') yearCounts[staffId].offDays++;
      else if (type === 'S') yearCounts[staffId].sickDays++;
    }
    if (sameDay(d, today) && staffById[staffId] && staffById[staffId].active) {
      todayRows.push({ name: staffById[staffId].name, dept: staffById[staffId].dept, type: type });
    }
  }

  var summary = staffAll.filter(function (s) {
    var c = yearCounts[s.id];
    return s.active || c.holidayDays > 0 || c.offDays > 0 || c.sickDays > 0;
  }).map(function (s) {
    var c = yearCounts[s.id];
    return {
      id: s.id, name: s.name, dept: s.dept, active: s.active,
      entitlement: s.holidayEntitlement,
      holidayDays: c.holidayDays, offDays: c.offDays, sickDays: c.sickDays,
      remaining: s.holidayEntitlement - c.holidayDays
    };
  });

  return {
    staff: staffActive.map(function (s) { return { id: s.id, name: s.name, dept: s.dept }; }),
    daysInMonth: daysInMonth,
    marks: marks,
    today: todayRows,
    summary: summary
  };
}

// ------------------------------------------------------------
// Advances
// ------------------------------------------------------------
function submitAdvance(entry) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ADVANCES);
    var date = parseDateStr(entry.date);
    var amount = assertAmount_(entry.amount || 0, 'advance');
    if (!entry.staffId) throw new Error('Staff member is required.');
    // v2.14: Supabase first (the shared ledger); the sheet is the backup copy
    sbCash_('hr_advance_add', { p_date: fmtDate_(date), p_staff_id: Number(entry.staffId), p_amount: amount, p_notes: entry.notes || '' });
    sh.appendRow([date, entry.staffId, amount, entry.notes || '']);
    return { ok: true };
  });
}

// ------------------------------------------------------------
// Requests
// Requests_Log columns: A=Date, B=StaffID ('' = general), C=Request,
//                       D=Due Date, E=Status (OPEN/DONE), F=Note
// ------------------------------------------------------------
function requestsSheet_() {
  var sh = ensureSheet_(SHEET_REQUESTS, ['Date', 'StaffID', 'Request', 'Due Date', 'Status', 'Note', 'Done Date']);
  // Older sheets predate the Done Date column: add the header once.
  if (String(sh.getRange(1, 7).getValue()) !== 'Done Date') {
    sh.getRange(1, 7).setValue('Done Date').setFontWeight('bold');
  }
  return sh;
}

function addRequest(entry) {
  return withLock_(function () {
    var sh = requestsSheet_();
    var date = parseDateStr(entry.date);
    var due = entry.dueDate ? parseDateStr(entry.dueDate) : '';
    if (!String(entry.request || '').trim()) throw new Error('Request text is required.');
    sh.appendRow([date, entry.staffId || '', String(entry.request).trim(), due, 'OPEN', entry.notes || '']);
    return { ok: true };
  });
}

function getRequests(includeDone) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REQUESTS);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  var staffById = {};
  getStaffList(false).forEach(function (s) { staffById[s.id] = s; });
  var todayStr = fmtDate_(new Date());

  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][2]) continue;
    var status = String(data[i][4] || 'OPEN');
    if (!includeDone && status === 'DONE') continue;
    var dueStr = data[i][3] ? fmtDate_(new Date(data[i][3])) : '';
    var sid = String(data[i][1] || '');
    out.push({
      row: i + 1,
      date: data[i][0] ? fmtDate_(new Date(data[i][0])) : '',
      staffId: sid,
      staffName: sid && staffById[sid] ? staffById[sid].name : (sid ? '(ID ' + sid + ')' : 'General'),
      request: String(data[i][2] || ''),
      dueDate: dueStr,
      status: status,
      overdue: status === 'OPEN' && dueStr !== '' && dueStr < todayStr,
      notes: String(data[i][5] || ''),
      doneDate: data[i][6] ? fmtDate_(new Date(data[i][6])) : ''
    });
  }
  out.sort(function (a, b) {
    if (a.status !== b.status) return a.status === 'OPEN' ? -1 : 1;
    if (a.status === 'DONE') {
      // History: most recently completed first
      var add = a.doneDate || '0000', bdd = b.doneDate || '0000';
      return add > bdd ? -1 : 1;
    }
    var ad = a.dueDate || '9999', bd = b.dueDate || '9999';
    return ad < bd ? -1 : 1;
  });
  return out;
}

function setRequestStatus(row, expectedRequestText, status) {
  return withLock_(function () {
    var sh = requestsSheet_();
    var current = sh.getRange(row, 3).getValue();
    if (String(current) !== String(expectedRequestText)) {
      throw new Error('This request changed since you loaded the list. Please reload and try again.');
    }
    sh.getRange(row, 5).setValue(status === 'DONE' ? 'DONE' : 'OPEN');
    // History needs a WHEN: stamp completion, clear it on reopen.
    sh.getRange(row, 7).setValue(status === 'DONE' ? new Date() : '');
    return { ok: true };
  });
}

function deleteRequest(row, expectedRequestText) {
  return withLock_(function () {
    var sh = requestsSheet_();
    var current = sh.getRange(row, 3).getValue();
    if (String(current) !== String(expectedRequestText)) {
      throw new Error('This request changed since you loaded the list. Please reload and try again.');
    }
    sh.deleteRow(row);
    return { ok: true };
  });
}

// ==============================================================
// COST EXCLUSIONS (v2.5)
// Cost_Exclusions_Log columns: A=Month (1-12), B=Year, C=StaffID
// Single source of truth. Excluding NEVER deletes or hides hours —
// it only removes that person's £ cost from monthly/yearly figures,
// and the UI always shows a banner for affected months so an
// adjusted figure is never mistaken for a raw one.
// ==============================================================
function costExclusionsSheet_() {
  return ensureSheet_(SHEET_COST_EXCLUSIONS, ['Month', 'Year', 'StaffID']);
}

// PURE: builds a lookup set {'staffId|month': true} from exclusion rows.
function buildExclusionSet_(rows) {
  var set = {};
  (rows || []).forEach(function (r) {
    var m = Number(r.month), sid = String(r.staffId);
    if (!isNaN(m) && m >= 1 && m <= 12 && sid !== '') set[sid + '|' + m] = true;
  });
  return set;
}

// All exclusions for a year, with resolved names for the UI.
function getCostExclusions(year) {
  year = Number(year);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_COST_EXCLUSIONS);
  if (!sh) return [];
  var staffById = {};
  getStaffList(false).forEach(function (s) { staffById[s.id] = s; });
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === undefined || data[i][0] === null) continue;
    if (Number(data[i][1]) !== year) continue;
    var sid = String(data[i][2]);
    out.push({
      month: Number(data[i][0]),
      year: year,
      staffId: sid,
      name: staffById[sid] ? staffById[sid].name : ('(ID ' + sid + ')')
    });
  }
  out.sort(function (a, b) {
    if (a.month !== b.month) return a.month - b.month;
    return a.name < b.name ? -1 : 1;
  });
  return out;
}

function addCostExclusion(month, year, staffId) {
  return withLock_(function () {
    month = Number(month); year = Number(year); staffId = String(staffId || '').trim();
    if (isNaN(month) || month < 1 || month > 12) throw new Error('Invalid month: ' + month);
    if (isNaN(year) || year < 2000 || year > 2100) throw new Error('Invalid year: ' + year);
    if (!staffId) throw new Error('Staff member is required.');
    var sh = costExclusionsSheet_();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === month && Number(data[i][1]) === year && String(data[i][2]) === staffId) {
        return { ok: true, already: true }; // idempotent: no duplicate rows
      }
    }
    sh.appendRow([month, year, staffId]);
    invalidateDataCaches_();
    return { ok: true, already: false };
  });
}

function removeCostExclusion(month, year, staffId) {
  return withLock_(function () {
    month = Number(month); year = Number(year); staffId = String(staffId || '').trim();
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_COST_EXCLUSIONS);
    if (!sh) throw new Error('No exclusions recorded yet.');
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === month && Number(data[i][1]) === year && String(data[i][2]) === staffId) {
        sh.deleteRow(i + 1);
        invalidateDataCaches_();
        return { ok: true };
      }
    }
    throw new Error('Exclusion not found — it may have been removed already. Please reload.');
  });
}

// ------------------------------------------------------------
// Pay Period Report
// ------------------------------------------------------------
// PURE: builds the 26th–25th pay period and its Monday–Sunday weekly
// buckets (first/last week may be short). Extracted so it can be tested
// without touching Sheets.
function buildPayPeriodWeeks_(year, month) {
  var periodEnd = new Date(year, month - 1, 25);
  var periodStart = new Date(year, month - 2, 26);
  var weeks = [];
  var cursor = new Date(periodStart);
  while (cursor <= periodEnd) {
    var jsDay = cursor.getDay();
    var daysUntilSunday = (jsDay === 0) ? 0 : (7 - jsDay);
    var weekEnd = new Date(cursor);
    weekEnd.setDate(weekEnd.getDate() + daysUntilSunday);
    if (weekEnd > periodEnd) weekEnd = new Date(periodEnd);
    weeks.push({ start: new Date(cursor), end: new Date(weekEnd) });
    cursor = new Date(weekEnd);
    cursor.setDate(cursor.getDate() + 1);
  }
  return { periodStart: periodStart, periodEnd: periodEnd, weeks: weeks };
}

function getPayPeriodReport(month, year) {
  month = Number(month);
  year = Number(year);
  var periodKey = year + '-' + month;

  var pp = buildPayPeriodWeeks_(year, month);
  var periodStart = pp.periodStart;
  var periodEnd = pp.periodEnd;
  var weeks = pp.weeks;

  var previewStart = new Date(year, month - 1, 26);
  var previewEnd = new Date(year, month, 0);

  var staff = getStaffList(false);
  var shHours = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var hoursData = shHours.getDataRange().getValues();
  // v2.14: advances + weekly boxes come from Supabase, rebuilt in the SAME
  // row shape the sheets had (row 0 = header), so the loops below are unchanged.
  var advRes = sbCash_('hr_advances_list', { p_from: fmtDate_(periodStart), p_to: fmtDate_(periodEnd) });
  var advData = [['Date', 'StaffID', 'Amount', 'Note']].concat((advRes.rows || []).map(function (a) {
    return [parseDateStr(a.date), String(a.staff_id), Number(a.amount), a.notes || ''];
  }));
  var manRes = sbCash_('hr_ppml_list', { p_period_start: fmtDate_(periodStart) });
  var manualData = [['PeriodKey', 'StaffID', 'WeekIndex', 'Value']].concat((manRes.rows || []).map(function (m) {
    var mv = String(m.value);
    return [periodKey, String(m.staff_id), Number(m.week_index), (mv === 'H' || mv === 'S') ? mv : Number(mv)];
  }));

  var rows = staff.map(function (s) {
    return {
      id: s.id, name: s.name, dept: s.dept, active: s.active,
      startDate: s.startDate, endDate: s.endDate,
      weeklyHours: weeks.map(function () { return ''; }),
      previewHours: 0, totalHours: 0, advances: 0,
      holidayDays: 0, sickDays: 0
    };
  });
  var byId = {};
  rows.forEach(function (r) { byId[r.id] = r; });

  for (var m = 1; m < manualData.length; m++) {
    var mrow = manualData[m];
    if (!mrow[0] || normPeriodKey_(mrow[0]) !== periodKey) continue;
    var mStaffId = String(mrow[1]);
    var weekIdx = Number(mrow[2]);
    if (!(mStaffId in byId) || isNaN(weekIdx) || weekIdx < 0 || weekIdx >= weeks.length) continue;
    var rawVal = String(mrow[3]).trim().toUpperCase();
    if (rawVal === 'H' || rawVal === 'S') {
      byId[mStaffId].weeklyHours[weekIdx] = rawVal; // week mark, counts as 0h
    } else {
      var value = Number(mrow[3]);
      byId[mStaffId].weeklyHours[weekIdx] = isNaN(value) ? '' : value;
    }
  }
  rows.forEach(function (r) {
    r.totalHours = r.weeklyHours.reduce(function (sum, v) { return sum + (Number(v) || 0); }, 0);
  });

  for (var i = 1; i < hoursData.length; i++) {
    var hrow = hoursData[i];
    if (!hrow[0]) continue;
    var d = new Date(hrow[0]);
    var staffId = String(hrow[4]);
    if (!(staffId in byId)) continue;
    if (d >= previewStart && d <= previewEnd) {
      byId[staffId].previewHours += Number(hrow[3]) || 0;
    }
  }

  // Holiday (H) and Sick (S) day marks inside the period, from the
  // Holidays tab — the single source of truth for absence marking.
  var shHol = SpreadsheetApp.getActive().getSheetByName(SHEET_HOLIDAYS);
  var holData = shHol ? shHol.getDataRange().getValues() : [];
  for (var hh = 1; hh < holData.length; hh++) {
    if (!holData[hh][0]) continue;
    var hDate = new Date(holData[hh][0]);
    if (hDate < periodStart || hDate > periodEnd) continue;
    var hId = String(holData[hh][1]);
    if (!(hId in byId)) continue;
    var hType = String(holData[hh][2]);
    if (hType === 'H') byId[hId].holidayDays++;
    else if (hType === 'S') byId[hId].sickDays++;
  }

  for (var j = 1; j < advData.length; j++) {
    if (!advData[j][0]) continue;
    var ad = new Date(advData[j][0]);
    if (ad < periodStart || ad > periodEnd) continue;
    var advStaffId = String(advData[j][1]);
    if (!(advStaffId in byId)) continue;
    byId[advStaffId].advances += Number(advData[j][2]) || 0;
  }

  var visibleRows = rows.filter(function (r) {
    return r.active || r.totalHours > 0 || r.previewHours > 0 || r.advances > 0;
  });

  return {
    periodKey: periodKey,
    periodStart: fmtDate_(periodStart),
    periodEnd: fmtDate_(periodEnd),
    previewStart: fmtDate_(previewStart),
    previewEnd: fmtDate_(previewEnd),
    weekLabels: weeks.map(function (w) {
      return Utilities.formatDate(w.start, Session.getScriptTimeZone(), 'dd MMM') + '-' + Utilities.formatDate(w.end, Session.getScriptTimeZone(), 'dd MMM');
    }),
    rows: visibleRows,
    staffSelect: staff.filter(function (s) { return s.active; }).map(function (s) {
      return { id: s.id, name: s.name, dept: s.dept };
    })
  };
}

// Writes the cell as TEXT first so Sheets never converts "2026-7" to a
// date; reads tolerate both text and old date-converted keys.
function setPayPeriodWeekValue(year, month, staffId, weekIndex, value) {
  return withLock_(function () {
    var periodKey = Number(year) + '-' + Number(month);
    // v2.14: Supabase first (validated the same way as the batch writer)
    var sbRaw = String(value === undefined || value === null ? '' : value).trim().toUpperCase();
    var sbEdit = { staff_id: Number(staffId), week_index: Number(weekIndex) };
    if (sbRaw === '') sbEdit.kind = 'clear';
    else if (sbRaw === 'H' || sbRaw === 'S') { sbEdit.kind = 'mark'; sbEdit.val = sbRaw; }
    else {
      var sbV = validateQuarterHours_(value);
      if (isNaN(sbV)) sbV = 0;
      if (sbV < 0 || sbV > 200) throw new Error('Invalid weekly hours: ' + value);
      sbEdit.kind = 'num'; sbEdit.val = sbV;
    }
    sbCash_('hr_ppml_write', { p_period_start: fmtDate_(buildPayPeriodWeeks_(Number(year), Number(month)).periodStart), p_edits: [sbEdit] });
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PAYPERIOD_MANUAL);
    var data = sh.getDataRange().getValues();

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (normPeriodKey_(data[i][0]) === periodKey && String(data[i][1]) === String(staffId) && Number(data[i][2]) === Number(weekIndex)) {
        foundRow = i + 1;
        break;
      }
    }

    if (value === '' || value === null || value === undefined) {
      if (foundRow !== -1) sh.deleteRow(foundRow);
      return { ok: true };
    }
    // Week marks: H = holiday, S = sick. Stored as text, count as 0 hours.
    var mark = String(value).trim().toUpperCase();
    if (mark === 'H' || mark === 'S') {
      if (foundRow !== -1) {
        sh.getRange(foundRow, 4).setNumberFormat('@').setValue(mark);
      } else {
        var mRow = sh.getLastRow() + 1;
        sh.getRange(mRow, 1).setNumberFormat('@').setValue(periodKey);
        sh.getRange(mRow, 4).setNumberFormat('@');
        sh.getRange(mRow, 2, 1, 3).setValues([[staffId, weekIndex, mark]]);
      }
      return { ok: true };
    }
    var v = validateQuarterHours_(value);
    if (isNaN(v)) v = 0;
    if (v < 0 || v > 200) throw new Error('Invalid weekly hours: ' + value);

    if (foundRow !== -1) {
      sh.getRange(foundRow, 4).setValue(v);
    } else {
      var row = sh.getLastRow() + 1;
      sh.getRange(row, 1).setNumberFormat('@').setValue(periodKey);
      sh.getRange(row, 2, 1, 3).setValues([[staffId, weekIndex, v]]);
    }
    return { ok: true };
  });
}

// ------------------------------------------------------------
// Dashboard / Monthly Summary
// ------------------------------------------------------------
function getDashboardData(year) {
  year = Number(year) || new Date().getFullYear();

  var cache = CacheService.getScriptCache();
  var cached = cache.get('dash_v2_' + year);
  if (cached) {
    var parsed = JSON.parse(cached);
    parsed.openRequests = countOpenRequests_(); // always live
    return parsed;
  }

  var histMap = getRateHistoryMap_();
  var staffById = {};
  getStaffList(false).forEach(function (s) { staffById[s.id] = s; });

  // v2.5: month+person cost exclusions. £ ONLY — hours are never touched.
  var exclRows = getCostExclusions(year);
  var exclSet = buildExclusionSet_(exclRows);

  var shRevenue = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
  var revenueData = shRevenue.getDataRange().getValues();
  var shHours = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var hoursData = shHours.getDataRange().getValues();

  var revenueByDate = {};
  for (var i = 1; i < revenueData.length; i++) {
    var row = revenueData[i];
    if (!row[0]) continue;
    var d = new Date(row[0]);
    if (d.getFullYear() !== year) continue;
    var key = dateKey_(d);
    revenueByDate[key] = revenueByDate[key] || { revenue: 0, date: d };
    revenueByDate[key].revenue += (Number(row[1]) || 0) + (Number(row[2]) || 0);
  }

  var shTips = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
  var tipsData = shTips ? shTips.getDataRange().getValues() : [];
  var tipsByDate = {};
  for (var t = 1; t < tipsData.length; t++) {
    var trow = tipsData[t];
    if (!trow[0]) continue;
    var td = new Date(trow[0]);
    if (td.getFullYear() !== year) continue;
    tipsByDate[dateKey_(td)] = (tipsByDate[dateKey_(td)] || 0) + (Number(trow[1]) || 0);
  }

  var costByDate = {};
  var unmatchedHours = 0;
  var unmatchedIds = {};
  var preHistoryHours = 0;
  var preHistoryIds = {};
  var excludedCostByMonth = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // reporting only
  var hoursByMonth = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // ALL logged hours, incl. unmatched
  for (var j = 1; j < hoursData.length; j++) {
    var hrow = hoursData[j];
    if (!hrow[0]) continue;
    var hd = new Date(hrow[0]);
    if (hd.getFullYear() !== year) continue;
    var dept = String(hrow[1]);
    var staffId = String(hrow[4]);
    var hours = Number(hrow[3]) || 0;
    hoursByMonth[hd.getMonth()] += hours; // hours are hours, whatever the rate situation
    var rw = rateForDate_(histMap, staffById, staffId, hd);
    if (!rw) {
      unmatchedHours += hours;
      unmatchedIds[staffId + ' (' + String(hrow[2] || '?') + ')'] = true;
      continue;
    }
    if (rw.preHistory) {
      // Date predates the first Rate_History entry. Costed with the
      // earliest known rate (approximation) and reported explicitly —
      // NOT silently trusted, NOT silently zeroed.
      preHistoryHours += hours;
      preHistoryIds[staffId + ' (' + String(hrow[2] || '?') + ')'] = true;
    }
    var cost = hours * rw.rate;
    // v2.5: excluded person+month -> the cost is measured but NOT added
    // to any £ figure; it is reported so the banner can show the amount.
    if (exclSet[staffId + '|' + (hd.getMonth() + 1)]) {
      excludedCostByMonth[hd.getMonth()] += cost;
      continue;
    }
    var hkey = dateKey_(hd);
    costByDate[hkey] = costByDate[hkey] || { total: 0, FOH: 0, BOH: 0, CLEANING: 0, date: hd };
    costByDate[hkey].total += cost;
    if (costByDate[hkey][dept] !== undefined) costByDate[hkey][dept] += cost;
  }

  var monthly = [];
  for (var m = 0; m < 12; m++) {
    monthly.push({ month: MONTH_NAMES[m], revenue: 0, cost: 0, foh: 0, boh: 0, cleaning: 0, tips: 0, hours: 0 });
  }
  Object.keys(revenueByDate).forEach(function (k) {
    monthly[revenueByDate[k].date.getMonth()].revenue += revenueByDate[k].revenue;
  });
  Object.keys(tipsByDate).forEach(function (k) {
    monthly[Number(k.split('-')[1]) - 1].tips += tipsByDate[k];
  });
  Object.keys(costByDate).forEach(function (k) {
    var v = costByDate[k];
    monthly[v.date.getMonth()].cost += v.total;
    monthly[v.date.getMonth()].foh += v.FOH;
    monthly[v.date.getMonth()].boh += v.BOH;
    monthly[v.date.getMonth()].cleaning += v.CLEANING;
  });
  monthly.forEach(function (mo, mi) {
    mo.hours = hoursByMonth[mi];
    mo.excludedCost = excludedCostByMonth[mi];
    mo.laborPct = safePct_(mo.cost, mo.revenue);
    mo.profit = mo.revenue - mo.cost;
  });

  // Day-of-week averages over the SAME day set (days that have revenue),
  // so avgRevenue and avgCost are comparable per day. Revenue days with no
  // hours logged are counted as £0 cost AND reported separately — a day
  // that traded with zero staff cost is almost always a missing entry.
  var dow = [];
  for (var w = 0; w < 7; w++) {
    dow.push({ day: DAY_NAMES[w], revenueSum: 0, costSum: 0, count: 0 });
  }
  var revenueDaysMissingHours = [];
  Object.keys(revenueByDate).forEach(function (k) {
    var v = revenueByDate[k];
    var idx = dayIndexMonFirst(v.date);
    var dayCost = costByDate[k] ? costByDate[k].total : 0;
    dow[idx].revenueSum += v.revenue;
    dow[idx].costSum += dayCost;
    dow[idx].count += 1;
    if (dayCost === 0 && v.revenue > 0) revenueDaysMissingHours.push(fmtDate_(v.date));
  });
  revenueDaysMissingHours.sort();
  var dowResult = dow.map(function (r) {
    return {
      day: r.day,
      avgRevenue: r.count > 0 ? r.revenueSum / r.count : 0,
      avgCost: r.count > 0 ? r.costSum / r.count : 0
    };
  });

  var totals = { revenue: 0, cost: 0, foh: 0, boh: 0, cleaning: 0, tips: 0, hours: 0, excludedCost: 0 };
  monthly.forEach(function (mo) {
    totals.revenue += mo.revenue;
    totals.cost += mo.cost;
    totals.foh += mo.foh;
    totals.boh += mo.boh;
    totals.cleaning += mo.cleaning;
    totals.tips += mo.tips;
    totals.hours += mo.hours;
    totals.excludedCost += mo.excludedCost;
  });
  totals.laborPct = safePct_(totals.cost, totals.revenue);
  totals.profit = totals.revenue - totals.cost;

  var settings = getSettings();

  var result = {
    year: year,
    monthly: monthly,
    dayOfWeek: dowResult,
    totals: totals,
    targetLaborPct: settings.targetLaborPct,
    unmatchedHours: unmatchedHours,
    unmatchedStaff: Object.keys(unmatchedIds),
    preHistoryHours: preHistoryHours,
    preHistoryStaff: Object.keys(preHistoryIds),
    revenueDaysMissingHours: revenueDaysMissingHours,
    costExclusions: exclRows
  };

  try { cache.put('dash_v2_' + year, JSON.stringify(result), 600); } catch (e) {}
  result.openRequests = countOpenRequests_();
  return result;
}

function countOpenRequests_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REQUESTS);
  if (!sh) return { open: 0, overdue: 0 };
  var data = sh.getDataRange().getValues();
  var todayStr = fmtDate_(new Date());
  var open = 0, overdue = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][2]) continue;
    if (String(data[i][4] || 'OPEN') !== 'OPEN') continue;
    open++;
    if (data[i][3] && fmtDate_(new Date(data[i][3])) < todayStr) overdue++;
  }
  return { open: open, overdue: overdue };
}

// ------------------------------------------------------------
// Weekly Tips
// v2.2: the tip weight in force ON EACH DAY is used both for the person's
// share AND the daily weighted total, so distributed tips always sum to
// the pool even if someone's weight changes mid-week.
// ------------------------------------------------------------

// PURE: distributes a daily pool (in PENCE) across entries by weighted
// hours using the largest-remainder method. Deterministic: ties go to the
// earlier entry in the list (list is already in staff order). Guarantees
// sum(shares) === poolPence whenever there is at least one participant
// with weightedHours > 0; otherwise everything is reported undistributed.
function distributeDailyTipsPence_(poolPence, entries) {
  var shares = entries.map(function () { return 0; });
  var totalWH = 0;
  entries.forEach(function (e) { totalWH += e.weightedHours; });
  if (poolPence <= 0 || totalWH <= 0 || entries.length === 0) {
    return { shares: shares, undistributed: poolPence > 0 ? poolPence : 0 };
  }
  var remainders = [];
  var sumFloors = 0;
  entries.forEach(function (e, i) {
    var raw = poolPence * e.weightedHours / totalWH;
    var fl = Math.floor(raw);
    shares[i] = fl;
    sumFloors += fl;
    remainders.push({ i: i, frac: raw - fl });
  });
  var leftover = poolPence - sumFloors;
  remainders.sort(function (a, b) {
    if (b.frac !== a.frac) return b.frac - a.frac; // largest remainder first
    return a.i - b.i; // deterministic tie-break: earlier staff order wins
  });
  for (var j = 0; j < leftover; j++) shares[remainders[j].i] += 1;
  return { shares: shares, undistributed: 0 };
}

function getTipsViewData(weekStartStr) {
  var weekStart = parseDateStr(weekStartStr);
  var days = [];
  for (var d = 0; d < 7; d++) {
    days.push(new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate() + d));
  }
  var dayKeys = days.map(dateKey_);
  var dayKeyIndex = {};
  dayKeys.forEach(function (k, i) { dayKeyIndex[k] = i; });

  var shTips = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
  var tipsData = shTips.getDataRange().getValues();
  var poolByDate = {};
  for (var i = 1; i < tipsData.length; i++) {
    var row = tipsData[i];
    if (!row[0]) continue;
    var key = dateKey_(new Date(row[0]));
    poolByDate[key] = (poolByDate[key] || 0) + (Number(row[1]) || 0);
  }

  var histMap = getRateHistoryMap_();
  var staffById = {};
  var orderById = {};
  getStaffList(false).forEach(function (s, i) { staffById[s.id] = s; orderById[s.id] = i; });

  var shHours = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var hoursData = shHours.getDataRange().getValues();

  // participants discovered from the log itself
  var participants = {}; // id -> {id, name, hours[7], weights[7]}
  var warnings = [];
  var warnedUnknown = {};
  var warnedPreHistory = {};

  for (var j = 1; j < hoursData.length; j++) {
    var hrow = hoursData[j];
    if (!hrow[0]) continue;
    if (String(hrow[1]) !== 'FOH') continue; // the LOG's department decides
    var k2 = dateKey_(new Date(hrow[0]));
    if (!(k2 in dayKeyIndex)) continue;
    var di = dayKeyIndex[k2];
    var staffId = String(hrow[4]);
    var hours = Number(hrow[3]) || 0;
    if (hours <= 0) continue;

    var rw = rateForDate_(histMap, staffById, staffId, days[di]);
    if (!rw) {
      // Unknown staff: NEVER silently assume weight 1. Exclude from the
      // distribution and surface it — their would-be share stays in the
      // pool and shows up as "undistributed" so no money silently moves.
      var uName = String(hrow[2] || ('ID ' + staffId));
      if (!warnedUnknown[staffId]) {
        warnedUnknown[staffId] = true;
        warnings.push('Unknown staff "' + uName + '" (ID ' + staffId + ') has FOH hours this week — no rate/weight record. Their hours are EXCLUDED from tip distribution; fix the staff record and reload.');
      }
      continue;
    }
    if (rw.preHistory && !warnedPreHistory[staffId]) {
      warnedPreHistory[staffId] = true;
      var pName = staffById[staffId] ? staffById[staffId].name : String(hrow[2] || ('ID ' + staffId));
      warnings.push('Tip weight for "' + pName + '" on some days predates their Rate_History — earliest known weight was used (approximation).');
    }
    var weight = rw.weight;

    if (!participants[staffId]) {
      participants[staffId] = {
        id: staffId,
        name: staffById[staffId] ? staffById[staffId].name : String(hrow[2] || ('ID ' + staffId)),
        hours: [0, 0, 0, 0, 0, 0, 0],
        weights: [1, 1, 1, 1, 1, 1, 1]
      };
    }
    participants[staffId].hours[di] += hours;
    participants[staffId].weights[di] = weight; // weight in force THAT day
  }

  var dailyPool = dayKeys.map(function (k) { return poolByDate[k] || 0; });

  var list = Object.keys(participants).map(function (id) { return participants[id]; });
  list.sort(function (a, b) {
    var oa = (a.id in orderById) ? orderById[a.id] : 9999;
    var ob = (b.id in orderById) ? orderById[b.id] : 9999;
    return oa - ob;
  });

  // PENNY DISTRIBUTION per day: convert the pool to pence, floor the raw
  // shares, hand leftover pennies to the largest remainders. The daily
  // distributed total is EXACTLY the daily pool — no float dust, no
  // penny appearing or vanishing.
  var tipsPence = list.map(function () { return [0, 0, 0, 0, 0, 0, 0]; });
  var dailyUndistributedPence = [0, 0, 0, 0, 0, 0, 0];
  for (var di = 0; di < 7; di++) {
    var poolPence = Math.round(dailyPool[di] * 100);
    var entries = list.map(function (p) {
      return { weightedHours: p.hours[di] * p.weights[di] };
    });
    var dist = distributeDailyTipsPence_(poolPence, entries);
    dist.shares.forEach(function (pence, li) { tipsPence[li][di] = pence; });
    dailyUndistributedPence[di] = dist.undistributed;
    if (dist.undistributed > 0) {
      warnings.push(days[di] ? (dmyStr_(fmtDate_(days[di])) + ': £' + (dist.undistributed / 100).toFixed(2) + ' pool has NO eligible FOH hours — nothing was distributed for that day.') : '');
    }
  }

  var result = list.map(function (p, li) {
    var tips = tipsPence[li].map(function (pence) { return pence / 100; });
    // Display weight: the weight on the latest day this person worked.
    var displayWeight = 1;
    for (var dj = 6; dj >= 0; dj--) {
      if (p.hours[dj] > 0) { displayWeight = p.weights[dj]; break; }
    }
    return {
      id: p.id, name: p.name, tipWeight: displayWeight,
      hours: p.hours, tips: tips,
      weekHours: p.hours.reduce(function (a, b) { return a + b; }, 0),
      weekTips: tips.reduce(function (a, b) { return a + b; }, 0)
    };
  });

  var dailyTotalHours = [0, 0, 0, 0, 0, 0, 0];
  list.forEach(function (p) {
    p.hours.forEach(function (h, dj) { dailyTotalHours[dj] += h; });
  });

  var weekUndistributed = dailyUndistributedPence.reduce(function (a, b) { return a + b; }, 0) / 100;

  return {
    days: days.map(fmtDate_),
    dailyPool: dailyPool,
    dailyTotalHours: dailyTotalHours,
    staff: result,
    weekTotalPool: dailyPool.reduce(function (a, b) { return a + b; }, 0),
    weekTotalTipsDistributed: result.reduce(function (a, b) { return a + b.weekTips; }, 0),
    weekUndistributed: weekUndistributed,
    warnings: warnings,
    recentPool: getRecentTipPool(30)
  };
}

// ==============================================================
// LPA BANKING BACKFILL (v2 — smart, date-based, safe to re-run)
// WARNING: do NOT run backfillFromBankingSheet() — it would push 3 years
// of Cash In Hand into Cash_List and break the cash balance. Kept only
// for reference; dry-run is safe.
// ==============================================================
var BANKING_SHEET_ID = '1FX_qVFBtuX6eWgHxbMpGQcHYhj5s-NFnVV0I3XbjwhQ'; // "Lpa Banking" (Primary)
var BANKING_WORKSHEET_NAME = 'BANKING';

function parseBankingDate_(val) {
  if (val instanceof Date) return val;
  var s = String(val || '').trim();
  var parts = s.split('/');
  if (parts.length === 3) {
    var d = Number(parts[0]), m = Number(parts[1]), y = Number(parts[2]);
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y, m - 1, d);
  }
  // Also accept ISO yyyy-mm-dd
  var iso = s.split('-');
  if (iso.length === 3 && iso[0].length === 4) {
    var y2 = Number(iso[0]), m2 = Number(iso[1]), d2 = Number(iso[2]);
    if (!isNaN(d2) && !isNaN(m2) && !isNaN(y2)) return new Date(y2, m2 - 1, d2);
  }
  return null;
}

function existingDateSet_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var set = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) set[dateKey_(new Date(data[i][0]))] = true;
  }
  return set;
}

// Run from the Apps Script editor. First run backfillFromBankingDryRun() to
// see the counts, then backfillFromBankingSheet() to actually write.
function backfillFromBankingDryRun() { return backfillFromBankingSheet_(true); }
function backfillFromBankingSheet() { return backfillFromBankingSheet_(false); }

function backfillFromBankingSheet_(dryRun) {
  return withLock_(function () {
    var bankingSS = SpreadsheetApp.openById(BANKING_SHEET_ID);
    var bankingSh = bankingSS.getSheetByName(BANKING_WORKSHEET_NAME);
    var data = bankingSh.getDataRange().getValues();

    var haveRevenue = existingDateSet_(SHEET_REVENUE_LOG);
    var haveTips = existingDateSet_(SHEET_TIPS_LOG);
    var haveCash = existingDateSet_(SHEET_CASH_LIST);

    var revenueRows = [], tipRows = [], cashRows = [];

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var date = parseBankingDate_(row[0]);
      if (!date) continue;
      var key = dateKey_(date);

      var takenIn = Number(row[1]) || 0;
      var serviceCharge = Number(row[2]) || 0;
      var cashTips = Number(row[4]) || 0;
      var cashInHand = Number(row[5]) || 0;

      if (!haveRevenue[key] && (takenIn > 0 || serviceCharge > 0)) {
        revenueRows.push([date, takenIn, serviceCharge, 'LPA Banking (backfill)']);
        haveRevenue[key] = true;
      }
      if (!haveTips[key] && cashTips > 0) {
        tipRows.push([date, cashTips, 'LPA Banking (backfill)']);
        haveTips[key] = true;
      }
      if (!haveCash[key] && cashInHand > 0) {
        cashRows.push([date, 'LPA Banking - Cash In Hand', cashInHand, 0, 'Automatic', 'LPA Banking (backfill)']);
        haveCash[key] = true;
      }
    }

    if (!dryRun) {
      if (revenueRows.length) {
        var shR = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
        shR.getRange(shR.getLastRow() + 1, 1, revenueRows.length, 4).setValues(revenueRows);
      }
      if (tipRows.length) {
        var shT = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
        shT.getRange(shT.getLastRow() + 1, 1, tipRows.length, 3).setValues(tipRows);
      }
      if (cashRows.length) {
        var shC = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH_LIST);
        shC.getRange(shC.getLastRow() + 1, 1, cashRows.length, 6).setValues(cashRows);
      }
      invalidateDataCaches_();
    }

    var result = {
      dryRun: !!dryRun,
      revenue: revenueRows.length,
      tips: tipRows.length,
      cash: cashRows.length
    };
    var msg = (dryRun ? '[DRY RUN — nothing written]\n' : 'Backfill complete.\n') +
      '\nMissing days that ' + (dryRun ? 'WOULD be' : 'were') + ' added:' +
      '\nRevenue: ' + result.revenue +
      '\nTip Pool: ' + result.tips +
      '\nCash List: ' + result.cash;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return result;
  });
}

// ==============================================================
// PURE TESTS (v2.3+) — run runAllTests() from the Apps Script editor.
// These NEVER read or write any sheet: they feed hand-built data into
// the pure calculation functions and assert the results. Safe to run
// on the live project at any time.
// ==============================================================
function runAllTests() {
  var results = [];
  function check(name, fn) {
    try {
      fn();
      results.push('PASS  ' + name);
    } catch (e) {
      results.push('FAIL  ' + name + ' — ' + e.message);
    }
  }
  function assertEq(actual, expected, label) {
    var a = JSON.stringify(actual), b = JSON.stringify(expected);
    if (a !== b) throw new Error((label || 'value') + ': expected ' + b + ', got ' + a);
  }
  function assertThrows(fn, label) {
    try { fn(); } catch (e) { return; }
    throw new Error((label || 'call') + ' should have thrown but did not');
  }
  var D = function (y, m, d) { return new Date(y, m - 1, d); };

  // 1) Mid-week tip weight change: weight 1 until Wed, 2 from Thu.
  check('Mid-week tip weight change', function () {
    var hist = { '7': [
      { time: D(2026, 7, 1).getTime(), rate: 13, weight: 1 },
      { time: D(2026, 7, 16).getTime(), rate: 13, weight: 2 } // Thursday
    ] };
    assertEq(rateForDate_(hist, {}, '7', D(2026, 7, 15)).weight, 1, 'Wednesday weight');
    assertEq(rateForDate_(hist, {}, '7', D(2026, 7, 16)).weight, 2, 'Thursday weight');
    assertEq(rateForDate_(hist, {}, '7', D(2026, 7, 17)).weight, 2, 'Friday weight');
  });

  // 2) Same-day rate change: two entries with the SAME effective midnight —
  //    the later-recorded one (later in sorted-stable order) must win.
  check('Same-day rate change (last write wins)', function () {
    var t = D(2026, 7, 17).getTime();
    var hist = { '3': [
      { time: t, rate: 14.0, weight: 1 },
      { time: t, rate: 14.5, weight: 1 }
    ] };
    assertEq(rateForDate_(hist, {}, '3', D(2026, 7, 17)).rate, 14.5, 'same-day rate');
    assertEq(rateForDate_(hist, {}, '3', D(2026, 7, 18)).rate, 14.5, 'next-day rate');
  });

  // 3) Date BEFORE first Rate_History entry: earliest rate used but
  //    explicitly flagged preHistory (never silent, never zeroed).
  check('Date before first Rate_History is flagged', function () {
    var hist = { '5': [{ time: D(2024, 6, 1).getTime(), rate: 12, weight: 1 }] };
    var r = rateForDate_(hist, {}, '5', D(2024, 1, 15));
    assertEq(r.rate, 12, 'earliest rate used');
    assertEq(r.preHistory, true, 'preHistory flag');
    var r2 = rateForDate_(hist, {}, '5', D(2024, 6, 1));
    assertEq(r2.preHistory, false, 'on/after first entry not flagged');
  });

  // 4) Missing staff entirely (no history, not in Staff): must be null,
  //    never a silent weight-1 guess.
  check('Missing staff returns null', function () {
    assertEq(rateForDate_({}, {}, '99', D(2026, 7, 1)), null, 'unknown staff');
  });

  // 5) Penny rounding: £100.00 across 3 equal shares = 3334+3333+3333,
  //    total EXACTLY 10000 pence, extra penny to the earliest in order.
  check('Penny rounding (largest remainder, deterministic)', function () {
    var d = distributeDailyTipsPence_(10000, [
      { weightedHours: 5 }, { weightedHours: 5 }, { weightedHours: 5 }
    ]);
    assertEq(d.shares, [3334, 3333, 3333], 'shares');
    assertEq(d.shares[0] + d.shares[1] + d.shares[2], 10000, 'sum equals pool');
    assertEq(d.undistributed, 0, 'nothing undistributed');
    // Weighted case: weight 2 vs 1 on equal hours -> 2:1 split of £50.00
    var d2 = distributeDailyTipsPence_(5000, [
      { weightedHours: 8 * 2 }, { weightedHours: 8 * 1 }
    ]);
    assertEq(d2.shares, [3333, 1667], '2:1 split');
    assertEq(d2.shares[0] + d2.shares[1], 5000, 'weighted sum equals pool');
    // Pool with no eligible hours -> fully undistributed, no invented shares
    var d3 = distributeDailyTipsPence_(4200, []);
    assertEq(d3.undistributed, 4200, 'pool with no FOH hours stays undistributed');
  });

  // 6) Negative tip weight rejected; valid range accepted.
  check('Negative tip weight rejection', function () {
    assertThrows(function () { assertTipWeight_(-1); }, 'weight -1');
    assertThrows(function () { assertTipWeight_(11); }, 'weight 11');
    assertThrows(function () { assertTipWeight_('abc'); }, 'weight abc');
    assertEq(assertTipWeight_(''), 1, 'empty defaults to 1');
    assertEq(assertTipWeight_(0), 0, 'zero is allowed (no tips)');
    assertEq(assertTipWeight_(2.5), 2.5, '2.5 accepted');
  });

  // 7) Holiday entitlement limits.
  check('Holiday entitlement limits', function () {
    assertEq(assertHolidayEntitlement_(0), 0, 'zero ok');
    assertEq(assertHolidayEntitlement_(366), 366, '366 ok');
    assertThrows(function () { assertHolidayEntitlement_(367); }, '367');
    assertThrows(function () { assertHolidayEntitlement_(-1); }, 'negative');
    assertEq(assertHolidayEntitlement_(''), 28, 'empty defaults to 28');
  });

  // 8) Zero revenue: labor % must be 0, never NaN/Infinity.
  check('Zero revenue guard', function () {
    assertEq(safePct_(500, 0), 0, 'cost with zero revenue');
    assertEq(safePct_(0, 0), 0, 'both zero');
    assertEq(safePct_(340, 1000), 0.34, 'normal case');
  });

  // 9) Leap year.
  check('Leap year day counts', function () {
    assertEq(daysInMonth_(2024, 2), 29, 'Feb 2024');
    assertEq(daysInMonth_(2026, 2), 28, 'Feb 2026');
    assertEq(daysInMonth_(2000, 2), 29, 'Feb 2000 (divisible by 400)');
    assertEq(daysInMonth_(1900, 2), 28, 'Feb 1900 (divisible by 100, not 400)');
    assertEq(daysInMonth_(2026, 12), 31, 'Dec');
  });

  // 10) 26–25 pay period structure.
  check('26–25 pay period weeks', function () {
    var pp = buildPayPeriodWeeks_(2026, 7); // 26 Jun – 25 Jul 2026
    assertEq(fmtTestDate_(pp.periodStart), '2026-06-26', 'period start');
    assertEq(fmtTestDate_(pp.periodEnd), '2026-07-25', 'period end');
    assertEq(fmtTestDate_(pp.weeks[0].start), '2026-06-26', 'first week starts on the 26th');
    assertEq(pp.weeks[0].end.getDay(), 0, 'first week ends on a Sunday');
    var last = pp.weeks[pp.weeks.length - 1];
    assertEq(fmtTestDate_(last.end), '2026-07-25', 'last week ends on the 25th');
    // Weeks must be contiguous with no gap and no overlap
    for (var i = 1; i < pp.weeks.length; i++) {
      var prevEnd = pp.weeks[i - 1].end;
      var expectedStart = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), prevEnd.getDate() + 1);
      assertEq(fmtTestDate_(pp.weeks[i].start), fmtTestDate_(expectedStart), 'week ' + i + ' contiguous');
    }
    // Middle weeks are full Mon–Sun
    if (pp.weeks.length > 2) {
      assertEq(pp.weeks[1].start.getDay(), 1, 'middle week starts Monday');
      assertEq(pp.weeks[1].end.getDay(), 0, 'middle week ends Sunday');
    }
    // February period across a leap year: 26 Jan – 25 Feb 2024
    var ppFeb = buildPayPeriodWeeks_(2024, 2);
    assertEq(fmtTestDate_(ppFeb.periodStart), '2024-01-26', 'leap Feb period start');
    assertEq(fmtTestDate_(ppFeb.periodEnd), '2024-02-25', 'leap Feb period end');
  });

  // 11) Quarter-hour input validation (regression guard for v2.2).
  check('Quarter-hour input validation', function () {
    assertEq(validateQuarterHours_('12.25'), 12.25, '12.25 ok');
    assertEq(validateQuarterHours_('12.5'), 12.5, '12.5 ok');
    assertEq(validateQuarterHours_('12,75'), 12.75, 'comma accepted as decimal point');
    assertEq(validateQuarterHours_('10.'), 10, 'trailing dot accepted as whole number');
    assertEq(isNaN(validateQuarterHours_('')), true, 'empty is NaN');
    assertThrows(function () { validateQuarterHours_('12.15'); }, '.15 rejected');
    assertThrows(function () { validateQuarterHours_('12.30'); }, '.30 rejected');
    assertThrows(function () { validateQuarterHours_('12.45'); }, '.45 rejected');
    assertThrows(function () { validateQuarterHours_('8.7'); }, '.7 rejected');
    assertThrows(function () { validateQuarterHours_('-1'); }, 'negative rejected');
  });

  // 12) v2.5: exclusion set building — month+person keys, junk ignored.
  check('Cost exclusion set (v2.5)', function () {
    var set = buildExclusionSet_([
      { month: 7, year: 2026, staffId: '3' },
      { month: 7, year: 2026, staffId: '12' },
      { month: 2, year: 2026, staffId: '3' },
      { month: 13, year: 2026, staffId: '5' },  // invalid month: ignored
      { month: 4, year: 2026, staffId: '' }     // missing staff: ignored
    ]);
    assertEq(set['3|7'], true, 'July person 3 excluded');
    assertEq(set['12|7'], true, 'July person 12 excluded');
    assertEq(set['3|2'], true, 'Feb person 3 excluded');
    assertEq(set['3|3'], undefined, 'March person 3 NOT excluded');
    assertEq(set['5|13'], undefined, 'invalid month ignored');
    assertEq(set['|4'], undefined, 'missing staff ignored');
  });

  var failed = results.filter(function (r) { return r.indexOf('FAIL') === 0; }).length;
  var summary = results.join('\n') + '\n\n' + (results.length - failed) + '/' + results.length + ' tests passed' + (failed ? ' — ' + failed + ' FAILED' : '');
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) {}
  return { passed: results.length - failed, failed: failed, details: results };
}

// Test-only date formatter (no Session dependency, so tests stay pure).
function fmtTestDate_(d) {
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var dd = ('0' + d.getDate()).slice(-2);
  return d.getFullYear() + '-' + mm + '-' + dd;
}

// ==============================================================
// END-OF-MONTH DETAIL (26th -> month end) + PDF EXPORT (v2.5)
// The pay period closes on the 25th; the 26th->end tail belongs to
// the NEXT period. This report lists per-person hour totals for that
// tail, grouped FLOOR (FOH) / KITCHEN (BOH) / CLEANING like the old
// Excel sheet, and can be exported as a one-tap PDF for WhatsApp.
// ==============================================================

// Quarter-accurate hours as text (mirror of the UI's hrs()).
function hrsStr_(n) {
  n = Number(n) || 0;
  return (n === Math.floor(n)) ? String(n) : n.toFixed(2);
}

function getMonthEndDetail(month, year) {
  month = Number(month); year = Number(year);
  var lastDay = daysInMonth_(year, month);
  var rangeStart = new Date(year, month - 1, 26);
  var rangeEnd = new Date(year, month - 1, lastDay);
  var rangeEndExcl = new Date(year, month - 1, lastDay + 1);

  var staff = getStaffList(false);
  var totalsById = {};
  staff.forEach(function (s) { totalsById[s.id] = 0; });
  var unknownTotal = 0;

  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var d = new Date(data[i][0]);
    if (d < rangeStart || d >= rangeEndExcl) continue;
    var id = String(data[i][4]);
    var h = Number(data[i][3]) || 0;
    if (id in totalsById) totalsById[id] += h;
    else unknownTotal += h;
  }

  // ALL active staff (zeros included); inactive only if they have hours
  // in the range. Groups keep staff-list order.
  var groups = { FLOOR: [], KITCHEN: [], CLEANING: [], OTHER: [] };
  var groupTotals = { FLOOR: 0, KITCHEN: 0, CLEANING: 0, OTHER: 0 };
  function groupOf(dept) {
    if (dept === 'FOH') return 'FLOOR';
    if (dept === 'BOH') return 'KITCHEN';
    if (dept === 'CLEANING') return 'CLEANING';
    return 'OTHER';
  }
  staff.forEach(function (s) {
    var h = totalsById[s.id];
    if (!s.active && h <= 0) return;
    var g = groupOf(s.dept);
    groups[g].push({ name: s.name, hours: h, active: s.active });
    groupTotals[g] += h;
  });
  if (unknownTotal > 0) {
    groups.OTHER.push({ name: '(Not in staff list / deleted)', hours: unknownTotal, active: true });
    groupTotals.OTHER += unknownTotal;
  }

  return {
    month: month, year: year,
    rangeStart: fmtDate_(rangeStart),
    rangeEnd: fmtDate_(rangeEnd),
    label: '26\u2013' + lastDay + ' ' + MONTH_NAMES[month - 1] + ' ' + year,
    groups: groups,
    groupTotals: groupTotals,
    grandTotal: groupTotals.FLOOR + groupTotals.KITCHEN + groupTotals.CLEANING + groupTotals.OTHER
  };
}

// Builds the PDF via a temporary Google Doc, saves ONLY the PDF in
// Drive (same name is overwritten on re-export, so no file pile-up),
// and returns a download link. The temp Doc is trashed.
function exportMonthEndPdf(month, year) {
  return withLock_(function () {
    var d = getMonthEndDetail(month, year);
    var title = 'Staff Hours ' + d.label;

    var doc = DocumentApp.create(title);
    var body = doc.getBody();
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);

    ['FLOOR', 'KITCHEN', 'CLEANING', 'OTHER'].forEach(function (g) {
      var rows = d.groups[g];
      if (!rows.length) return;
      body.appendParagraph(g === 'OTHER' ? 'OTHER' : g).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      var cells = rows.map(function (r) { return [r.name, hrsStr_(r.hours)]; });
      cells.push(['TOTAL', hrsStr_(d.groupTotals[g])]);
      var table = body.appendTable(cells);
      var totalRow = table.getRow(table.getNumRows() - 1);
      for (var c = 0; c < 2; c++) totalRow.getCell(c).editAsText().setBold(true);
    });
    body.appendParagraph('GRAND TOTAL: ' + hrsStr_(d.grandTotal) + ' h').editAsText().setBold(true);
    doc.saveAndClose();

    var pdfBlob = doc.getAs('application/pdf').setName(title + '.pdf');

    // Overwrite an existing PDF with the same name instead of piling up
    var existing = DriveApp.getFilesByName(title + '.pdf');
    while (existing.hasNext()) existing.next().setTrashed(true);

    var pdfFile = DriveApp.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { /* workspace policy may forbid; owner can still share from Drive */ }

    // Trash the temporary Doc; only the PDF remains
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return {
      ok: true,
      fileName: title + '.pdf',
      viewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view',
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
    };
  });
}

// ==============================================================
// ENTRY CHECK (v2.5) — "which days have I not entered yet?"
// For the last N days: does each day have Hours / Revenue / Tips?
// Answers "where did I leave off" at a glance.
// ==============================================================
function getEntryStatus(nDays, offsetDays, staffId) {
  nDays = Math.max(7, Math.min(Number(nDays) || 14, 31));
  offsetDays = Math.max(0, Math.min(Number(offsetDays) || 0, 3650)); // up to ~10 years back
  staffId = staffId ? String(staffId) : ''; // '' = all staff
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate() - offsetDays);

  // idCol/idVal: optional filter (used to narrow Hours to one person)
  function dateSetWithSum_(sheetName, sumCol, idCol, idVal) {
    var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    var map = {};
    if (!sh) return map;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (idVal && String(data[i][idCol]) !== idVal) continue;
      var k = dateKey_(new Date(data[i][0]));
      map[k] = (map[k] || 0) + (sumCol >= 0 ? (Number(data[i][sumCol]) || 0) : 1);
    }
    return map;
  }

  var hoursByDay = dateSetWithSum_(SHEET_HOURS_LOG, 3, 4, staffId); // sum of hours (per person if staffId)
  var revByDay = dateSetWithSum_(SHEET_REVENUE_LOG, -1);            // presence (always whole restaurant)
  var tipsByDay = dateSetWithSum_(SHEET_TIPS_LOG, -1);              // presence (always whole restaurant)

  var rows = [];
  for (var d = 0; d < nDays; d++) {
    var day = new Date(today.getFullYear(), today.getMonth(), today.getDate() - d);
    var k = dateKey_(day);
    rows.push({
      date: fmtDate_(day),
      dayName: DAY_NAMES[(day.getDay() + 6) % 7],
      hours: hoursByDay[k] || 0,
      hasHours: (hoursByDay[k] || 0) > 0,
      hasRevenue: (revByDay[k] || 0) > 0,
      hasTips: (tipsByDay[k] || 0) > 0
    });
  }

  // Last day that has ANY hours logged (the "where I left off" marker)
  var lastHoursDate = '';
  var allHourKeys = Object.keys(hoursByDay);
  for (var j = 0; j < allHourKeys.length; j++) {
    var p = allHourKeys[j].split('-');
    var dt = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    var f = fmtDate_(dt);
    if (dt <= today && f > lastHoursDate) lastHoursDate = f;
  }

  return { rows: rows, lastHoursDate: lastHoursDate };
}

// ==============================================================
// PAY PERIOD SUMMARY PDF (v2.5) — the old Excel's right-hand block:
// per-person PERIOD total hours (manual weekly entries) + advances,
// grouped FLOOR / KITCHEN / CLEANING, one-tap PDF for WhatsApp.
// ==============================================================
function exportPayPeriodPdf(month, year) {
  return withLock_(function () {
    var rep = getPayPeriodReport(month, year);
    var title = 'Staff Hours ' + dmyStr_(rep.periodStart) + ' - ' + dmyStr_(rep.periodEnd);

    var groups = { FLOOR: [], KITCHEN: [], CLEANING: [], OTHER: [] };
    var totals = { FLOOR: { h: 0, a: 0 }, KITCHEN: { h: 0, a: 0 }, CLEANING: { h: 0, a: 0 }, OTHER: { h: 0, a: 0 } };
    function groupOf(dept) {
      if (dept === 'FOH') return 'FLOOR';
      if (dept === 'BOH') return 'KITCHEN';
      if (dept === 'CLEANING') return 'CLEANING';
      return 'OTHER';
    }
    rep.rows.forEach(function (r) {
      var g = groupOf(r.dept);
      groups[g].push(r);
      totals[g].h += Number(r.totalHours) || 0;
      totals[g].a += Number(r.advances) || 0;
    });
    var grandH = totals.FLOOR.h + totals.KITCHEN.h + totals.CLEANING.h + totals.OTHER.h;
    var grandA = totals.FLOOR.a + totals.KITCHEN.a + totals.CLEANING.a + totals.OTHER.a;

    var doc = DocumentApp.create(title);
    var body = doc.getBody();
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);

    ['FLOOR', 'KITCHEN', 'CLEANING', 'OTHER'].forEach(function (g) {
      var rows = groups[g];
      if (!rows.length) return;
      body.appendParagraph(g).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      var cells = [['Name', 'In', 'Out', 'Hours', 'Advance', 'Notes']];
      rows.forEach(function (r) {
        var adv = (Number(r.advances) || 0);
        var notes = [];
        var hWeeks = 0, sWeeks = 0;
        (r.weeklyHours || []).forEach(function (w) {
          if (w === 'H') hWeeks++; else if (w === 'S') sWeeks++;
        });
        // v2.9: Holiday is NOT printed in the report notes (user request);
        // week mark still wins over the day count for Sick.
        if (sWeeks > 0) notes.push('Sick ' + sWeeks + 'wk');
        else if (Number(r.sickDays) > 0) notes.push('Sick ' + r.sickDays + 'd');
        cells.push([
          r.name,
          r.startDate ? dmyStr_(r.startDate) : '',
          r.endDate ? dmyStr_(r.endDate) : '',
          hrsStr_(r.totalHours),
          adv > 0 ? ('\u00a3' + adv.toFixed(2)) : '',
          notes.join(', ')
        ]);
      });
      cells.push(['TOTAL', '', '', hrsStr_(totals[g].h), totals[g].a > 0 ? ('\u00a3' + totals[g].a.toFixed(2)) : '', '']);
      var table = body.appendTable(cells);
      for (var c = 0; c < 6; c++) {
        table.getRow(0).getCell(c).editAsText().setBold(true);
        table.getRow(table.getNumRows() - 1).getCell(c).editAsText().setBold(true);
      }
    });
    var gp = body.appendParagraph('GRAND TOTAL: ' + hrsStr_(grandH) + ' h' + (grandA > 0 ? ('   |   Advances: \u00a3' + grandA.toFixed(2)) : ''));
    gp.editAsText().setBold(true);
    doc.saveAndClose();

    var pdfBlob = doc.getAs('application/pdf').setName(title + '.pdf');
    var existing = DriveApp.getFilesByName(title + '.pdf');
    while (existing.hasNext()) existing.next().setTrashed(true);
    var pdfFile = DriveApp.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { /* owner can still share from Drive */ }
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return {
      ok: true,
      fileName: title + '.pdf',
      viewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view',
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
    };
  });
}

// ============================================================
// v2.6 — Weekly Grid batch save. ONE server call applies every pending
// cell change: hour updates, row deletes, new rows and H/OFF/S marks.
// Same safety rules as the single-call paths:
//  - every new hours value is validated FIRST (one bad value = nothing written)
//  - hour rows are matched BY CONTENT (date + staff + old hours), row hint first
//  - a row that can no longer be found is reported back, never guessed at
//  - new rows are refused if that person already has hours that day
//  - marks go through setHolidayMarkCore_ (Holidays_Log single source of truth)
// Edit shape: { date:'yyyy-mm-dd', staffId, name, dept, row,
//               oldHours:number|null, newHours:number|null,
//               oldMark:''|'H'|'OFF'|'S', newMark:''|'H'|'OFF'|'S' }
// ============================================================
function applyGridEdits(edits) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOURS_LOG);

    // ---- pass 1: validate EVERYTHING before touching anything ----
    edits.forEach(function (e) {
      if (e.newHours !== null && e.newHours !== undefined && e.newHours !== '') {
        var h = validateQuarterHours_(e.newHours); // throws on bad input
        assertHours_(h);
        e._newH = h;
      } else {
        e._newH = null;
      }
      var nm = String(e.newMark || '');
      if (nm && ['H', 'OFF', 'S'].indexOf(nm) === -1) throw new Error('Invalid mark: ' + nm);
    });

    // ---- pass 2: resolve rows on ONE snapshot ----
    var data = sh.getDataRange().getValues();
    var used = {}; // rows already claimed by this batch (duplicate pairs stay distinct)

    function findRow(e) { // 0-based data index, -1 if not found
      function matches(i) {
        if (used[i]) return false;
        var v = data[i];
        if (!v || !v[0]) return false;
        return fmtDate_(new Date(v[0])) === e.date &&
          String(v[4]) === String(e.staffId) &&
          Math.abs((Number(v[3]) || 0) - Number(e.oldHours)) < 0.001;
      }
      var hinted = Number(e.row) - 1;
      if (hinted >= 1 && hinted < data.length && matches(hinted)) return hinted;
      for (var i = 1; i < data.length; i++) if (matches(i)) return i;
      return -1;
    }
    function hasAnyRow(e) { // does this person already have hours that day?
      for (var i = 1; i < data.length; i++) {
        if (used[i]) continue;
        var v = data[i];
        if (!v || !v[0]) continue;
        if (fmtDate_(new Date(v[0])) === e.date && String(v[4]) === String(e.staffId)) return true;
      }
      return false;
    }

    var updates = [], deletes = [], adds = [], markOps = [], failed = [];
    edits.forEach(function (e) {
      var hasOld = (e.oldHours !== null && e.oldHours !== undefined && e.oldHours !== '');
      if (hasOld) {
        var idx = findRow(e);
        if (idx === -1) {
          failed.push({ date: e.date, staffId: e.staffId, name: e.name, reason: 'logged row not found — it was changed or deleted elsewhere' });
          return; // do NOT apply the mark either: the cell's state is stale
        }
        used[idx] = true;
        if (e._newH !== null) updates.push({ row: idx + 1, h: e._newH });
        else deletes.push(idx + 1);
      } else if (e._newH !== null) {
        if (hasAnyRow(e)) {
          failed.push({ date: e.date, staffId: e.staffId, name: e.name, reason: 'already has hours that day (added elsewhere since the grid was loaded)' });
          return;
        }
        adds.push([parseDateStr(e.date), e.dept, e.name, e._newH, e.staffId]);
      }
      if (String(e.newMark || '') !== String(e.oldMark || '')) {
        markOps.push({ date: e.date, id: e.staffId, type: String(e.newMark || '') });
      }
    });

    // ---- pass 3: write. Updates first (original row numbers still valid),
    // then deletes bottom-up in consecutive runs (numbers stay valid while
    // deleting), then adds, then marks. ----
    updates.forEach(function (u) { sh.getRange(u.row, 4).setValue(u.h); });
    deletes.sort(function (a, b) { return b - a; });
    for (var di = 0; di < deletes.length; ) {
      var last = deletes[di], first = last;
      while (di + 1 < deletes.length && deletes[di + 1] === first - 1) { di++; first = deletes[di]; }
      sh.deleteRows(first, last - first + 1);
      di++;
    }
    if (adds.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, adds.length, 5).setValues(adds);
    }

    // marks: ONE Holidays_Log read, grouped writes — same semantics as
    // setHolidayMarkCore_ (update in place / delete when cleared / append new)
    if (markOps.length > 0) {
      var hsh = SpreadsheetApp.getActive().getSheetByName(SHEET_HOLIDAYS);
      var hdata = hsh.getDataRange().getValues();
      var idxByKey = {};
      for (var hi = 1; hi < hdata.length; hi++) {
        if (!hdata[hi][0]) continue;
        idxByKey[dateKey_(new Date(hdata[hi][0])) + '|' + String(hdata[hi][1])] = hi;
      }
      var mUpdates = [], mDeletes = [], mAdds = [];
      markOps.forEach(function (m) {
        var at = idxByKey[dateKey_(parseDateStr(m.date)) + '|' + String(m.id)];
        if (!m.type) {
          if (at !== undefined) mDeletes.push(at + 1);
        } else if (at !== undefined) {
          mUpdates.push({ row: at + 1, t: m.type });
        } else {
          mAdds.push([parseDateStr(m.date), m.id, m.type, '']);
        }
      });
      mUpdates.forEach(function (u) { hsh.getRange(u.row, 3).setValue(u.t); });
      mDeletes.sort(function (a, b) { return b - a; });
      for (var mi = 0; mi < mDeletes.length; ) {
        var mlast = mDeletes[mi], mfirst = mlast;
        while (mi + 1 < mDeletes.length && mDeletes[mi + 1] === mfirst - 1) { mi++; mfirst = mDeletes[mi]; }
        hsh.deleteRows(mfirst, mlast - mfirst + 1);
        mi++;
      }
      if (mAdds.length > 0) {
        hsh.getRange(hsh.getLastRow() + 1, 1, mAdds.length, 4).setValues(mAdds);
      }
    }

    var applied = updates.length + deletes.length + adds.length + markOps.length;
    if (applied > 0) invalidateDataCaches_();
    return { applied: applied, updates: updates.length, deletes: deletes.length, adds: adds.length, marks: markOps.length, failed: failed };
  });
}


// ============================================================
// v2.8 — Tip pool corrections. Entries are matched BY CONTENT
// (date + amount, ±0.001), trying the client's row hint first, exactly like
// updateHoursEntry / deleteHoursEntry. Notes are NOT part of the key (they
// are editable). Two identical rows are interchangeable — the hint picks one.
// ============================================================
function getTipPoolEntries(n) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      row: i + 1, // 1-based sheet row, used as a hint by the editors
      date: fmtDate_(new Date(data[i][0])),
      amount: Number(data[i][1]) || 0,
      notes: data[i][2] || ''
    });
  }
  rows.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : b.row - a.row); });
  return rows.slice(0, n || 30);
}

function findTipPoolRowByContent_(sh, expected) {
  var data = sh.getDataRange().getValues();
  function matches(i) {
    var v = data[i];
    if (!v || !v[0]) return false;
    return fmtDate_(new Date(v[0])) === expected.date &&
      Math.abs((Number(v[1]) || 0) - Number(expected.amount)) < 0.001;
  }
  var hinted = Number(expected.row) - 1;
  if (hinted >= 1 && hinted < data.length && matches(hinted)) return hinted + 1;
  for (var i = 1; i < data.length; i++) if (matches(i)) return i + 1;
  return -1;
}

function updateTipPoolEntry(expected, newAmount, newNotes) {
  return withLock_(function () {
    var amount = assertAmount_(newAmount, 'tips'); // validate BEFORE touching the sheet
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
    var row = findTipPoolRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Tip entry not found (' + expected.date + ', £' + expected.amount + '). It may have been edited or deleted already — the list will reload.');
    }
    sh.getRange(row, 2, 1, 2).setValues([[amount, String(newNotes || '')]]);
    invalidateDataCaches_();
    return { ok: true };
  });
}

function deleteTipPoolEntry(expected) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
    var row = findTipPoolRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Tip entry not found (' + expected.date + ', £' + expected.amount + '). It may have been edited or deleted already — the list will reload.');
    }
    sh.deleteRow(row);
    invalidateDataCaches_();
    return { ok: true };
  });
}


// ============================================================
// v2.10 — Staff Hours (pay period) batch save. ONE call applies every
// pending weekly box: numbers, H/S week marks and cleared boxes.
// Mirrors setPayPeriodWeekValue exactly (same validation, same PeriodKey
// text-format guard, same H/S-as-text storage); validate-first so a single
// bad value means NOTHING is written.
// edits: [{ staffId, weekIndex, value: ''|'H'|'S'|number-string }]
// ============================================================
function setPayPeriodWeekValues(year, month, edits) {
  return withLock_(function () {
    var periodKey = Number(year) + '-' + Number(month);

    // ---- pass 1: validate + normalise EVERYTHING before touching the sheet
    edits.forEach(function (e) {
      var raw = String(e.value === undefined || e.value === null ? '' : e.value).trim().toUpperCase();
      if (raw === '') { e._kind = 'clear'; return; }
      if (raw === 'H' || raw === 'S') { e._kind = 'mark'; e._val = raw; return; }
      var v = validateQuarterHours_(e.value); // throws -> nothing written
      if (isNaN(v)) { e._kind = 'clear'; return; }
      if (v < 0 || v > 200) throw new Error('Invalid weekly hours: ' + e.value);
      e._kind = 'num'; e._val = v;
    });

    // ---- v2.14: write Supabase first, all-or-nothing (the shared copy)
    sbCash_('hr_ppml_write', {
      p_period_start: fmtDate_(buildPayPeriodWeeks_(Number(year), Number(month)).periodStart),
      p_edits: edits.map(function (e) {
        var o = { staff_id: Number(e.staffId), week_index: Number(e.weekIndex), kind: e._kind };
        if (e._kind !== 'clear') o.val = e._val;
        return o;
      })
    });

    // ---- pass 2: resolve rows on ONE snapshot
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PAYPERIOD_MANUAL);
    var data = sh.getDataRange().getValues();
    var rowByKey = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || normPeriodKey_(data[i][0]) !== periodKey) continue;
      rowByKey[String(data[i][1]) + '|' + Number(data[i][2])] = i + 1;
    }

    var updates = [], deletes = [], addNums = [], addMarks = [];
    edits.forEach(function (e) {
      var found = rowByKey[String(e.staffId) + '|' + Number(e.weekIndex)];
      if (e._kind === 'clear') {
        if (found) deletes.push(found);
        return;
      }
      if (found) updates.push({ row: found, val: e._val, mark: (e._kind === 'mark') });
      else (e._kind === 'mark' ? addMarks : addNums).push([periodKey, e.staffId, Number(e.weekIndex), e._val]);
    });

    // ---- pass 3: write. Updates first (snapshot row numbers still valid),
    // then deletes bottom-up in consecutive runs, then batched appends with
    // the same text-format guards as the single-cell path.
    updates.forEach(function (u) {
      var rg = sh.getRange(u.row, 4);
      if (u.mark) rg.setNumberFormat('@');
      rg.setValue(u.val);
    });
    deletes.sort(function (a, b) { return b - a; });
    for (var di = 0; di < deletes.length; ) {
      var last = deletes[di], first = last;
      while (di + 1 < deletes.length && deletes[di + 1] === first - 1) { di++; first = deletes[di]; }
      sh.deleteRows(first, last - first + 1);
      di++;
    }
    if (addNums.length > 0) {
      var r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 1, addNums.length, 1).setNumberFormat('@'); // PeriodKey stays text
      sh.getRange(r0, 1, addNums.length, 4).setValues(addNums);
    }
    if (addMarks.length > 0) {
      var r1 = sh.getLastRow() + 1;
      sh.getRange(r1, 1, addMarks.length, 1).setNumberFormat('@'); // PeriodKey stays text
      sh.getRange(r1, 4, addMarks.length, 1).setNumberFormat('@'); // H/S stored as text
      sh.getRange(r1, 1, addMarks.length, 4).setValues(addMarks);
    }

    return {
      applied: updates.length + deletes.length + addNums.length + addMarks.length,
      updates: updates.length, deletes: deletes.length, adds: addNums.length + addMarks.length
    };
  });
}


// ============================================================
// v2.11 — Missed banking day: one call fills the gaps in all three
// ledgers that LPA Banking (Streamlit) normally writes for a day.
// GAP-FILLING + IDEMPOTENT: a piece that already exists for the date is
// NEVER touched or doubled — it is skipped and its existing figures are
// reported, so pressing SAVE twice (or Streamlit retrying later) is safe.
// entry: { date, netSales, serviceCharge, cashTips, cashInHand, notes }
// returns per piece: null (nothing to write) | {added:true}
//                  | {added:false, existing...} (skipped, already there)
// ============================================================
function submitMissedBankingDay(entry) {
  return withLock_(function () {
    // validate EVERYTHING before touching any sheet
    var date = parseDateStr(entry.date);
    var key = dateKey_(date);
    var net = assertAmount_(entry.netSales || 0, 'net sales');
    var sc = assertAmount_(entry.serviceCharge || 0, 'service charge');
    var tips = assertAmount_(entry.cashTips || 0, 'cash tips');
    var cash = assertAmount_(entry.cashInHand || 0, 'cash in hand');
    var note = String(entry.notes || '').trim() || 'Manual entry';

    var result = { revenue: null, tips: null, cash: null };
    var wrote = false;

    // ---- Revenue_Log: one entry per date ----
    if (net > 0 || sc > 0) {
      var shR = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
      var dataR = shR.getDataRange().getValues();
      var foundR = null;
      for (var i = 1; i < dataR.length; i++) {
        if (dataR[i][0] && dateKey_(new Date(dataR[i][0])) === key) {
          foundR = { netSales: Number(dataR[i][1]) || 0, serviceCharge: Number(dataR[i][2]) || 0 };
          break;
        }
      }
      if (foundR) {
        result.revenue = { added: false, existingNet: foundR.netSales, existingSC: foundR.serviceCharge };
      } else {
        shR.appendRow([date, net, sc, note]);
        result.revenue = { added: true };
        wrote = true;
      }
    }

    // ---- Tip_Pool_Log: skip if the day already has ANY pool entry ----
    if (tips > 0) {
      var shT = SpreadsheetApp.getActive().getSheetByName(SHEET_TIPS_LOG);
      var dataT = shT.getDataRange().getValues();
      var existingTips = 0, hasT = false;
      for (var t = 1; t < dataT.length; t++) {
        if (dataT[t][0] && dateKey_(new Date(dataT[t][0])) === key) {
          hasT = true;
          existingTips += Number(dataT[t][1]) || 0;
        }
      }
      if (hasT) {
        result.tips = { added: false, existingAmount: existingTips };
      } else {
        shT.appendRow([date, tips, note]);
        result.tips = { added: true };
        wrote = true;
      }
    }

    // ---- Cash_List: skip if the day already has a Cash In Hand row ----
    // (other manual rows on that date — expenses etc. — do NOT block this)
    if (cash > 0) {
      var shC = SpreadsheetApp.getActive().getSheetByName(SHEET_CASH_LIST);
      var dataC = shC.getDataRange().getValues();
      var foundC = null;
      for (var c = 1; c < dataC.length; c++) {
        if (!dataC[c][0]) continue;
        if (dateKey_(new Date(dataC[c][0])) !== key) continue;
        if (String(dataC[c][1] || '').indexOf('Cash In Hand') === -1) continue;
        foundC = { amount: Number(dataC[c][2]) || 0 };
        break;
      }
      if (foundC) {
        result.cash = { added: false, existingAmount: foundC.amount };
      } else {
        shC.appendRow([date, 'LPA Banking - Cash In Hand', cash, 0, 'Manual', note]);
        // v2.13: mirror to Supabase (skipped there if that day already has Cash In Hand)
        sbCash_('hr_cash_write', { p_op: 'insert', p_old: null, p_new: { date: fmtDate_(date), description: 'LPA Banking - Cash In Hand', income: cash, expense: 0, source: 'MANUAL', notes: note, unique_cash_in_hand: true } });
        result.cash = { added: true };
        wrote = true;
      }
    }

    if (wrote) invalidateDataCaches_();
    return result;
  });
}


// ============================================================
// v2.12 — Take-in corrections. Entries are matched BY CONTENT
// (date + net sales + service charge, ±0.001), row hint first — the same
// safety pattern as Hours/Tips editors. The DATE is the ledger's
// one-entry-per-day key and is NOT editable here: to move a day, delete
// the row and re-enter it on the right date with the Manual Entry card.
// ============================================================
function findRevenueRowByContent_(sh, expected) {
  var data = sh.getDataRange().getValues();
  function matches(i) {
    var v = data[i];
    if (!v || !v[0]) return false;
    return fmtDate_(new Date(v[0])) === expected.date &&
      Math.abs((Number(v[1]) || 0) - Number(expected.netSales)) < 0.001 &&
      Math.abs((Number(v[2]) || 0) - Number(expected.serviceCharge)) < 0.001;
  }
  var hinted = Number(expected.row) - 1;
  if (hinted >= 1 && hinted < data.length && matches(hinted)) return hinted + 1;
  for (var i = 1; i < data.length; i++) if (matches(i)) return i + 1;
  return -1;
}

function updateRevenueEntry(expected, newEntry) {
  return withLock_(function () {
    // validate BEFORE touching the sheet
    var net = assertAmount_(newEntry.netSales || 0, 'net sales');
    var sc = assertAmount_(newEntry.serviceCharge || 0, 'service charge');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
    var row = findRevenueRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Take-in entry not found (' + dmyStr_(expected.date) + ', £' + expected.netSales + ' + SC £' + expected.serviceCharge + '). It may have been edited or deleted already — the list will reload.');
    }
    sh.getRange(row, 2, 1, 3).setValues([[net, sc, String(newEntry.notes || '')]]);
    invalidateDataCaches_();
    return { ok: true };
  });
}

function deleteRevenueEntry(expected) {
  return withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REVENUE_LOG);
    var row = findRevenueRowByContent_(sh, expected);
    if (row === -1) {
      throw new Error('Take-in entry not found (' + dmyStr_(expected.date) + ', £' + expected.netSales + ' + SC £' + expected.serviceCharge + '). It may have been deleted already — the list will reload.');
    }
    sh.deleteRow(row);
    invalidateDataCaches_();
    return { ok: true };
  });
}
// v2.13: one-click check — Run this from the Apps Script editor after setting
// the Script Properties. Reads only; writes nothing.
function testCashSync() {
  var rows = sbCash_('hr_cash_list', {}).rows;
  var bal = 0;
  rows.forEach(function (r) { bal += (Number(r.income) || 0) - (Number(r.expense) || 0); });
  Logger.log('Supabase cash list OK: ' + rows.length + ' rows, balance £' + bal.toFixed(2));
}


// --- browser hooks (hr-shim.js) ---
sbCash_ = window.HRSHIM_sbCash;
return {
  doGet: doGet,
  withLock_: withLock_,
  assertDateStr_: assertDateStr_,
  assertHours_: assertHours_,
  assertAmount_: assertAmount_,
  assertTipWeight_: assertTipWeight_,
  assertHolidayEntitlement_: assertHolidayEntitlement_,
  safePct_: safePct_,
  daysInMonth_: daysInMonth_,
  validateQuarterHours_: validateQuarterHours_,
  normPeriodKey_: normPeriodKey_,
  todayMidnight_: todayMidnight_,
  parseDateStr: parseDateStr,
  sameDay: sameDay,
  dayIndexMonFirst: dayIndexMonFirst,
  dateKey_: dateKey_,
  dmyStr_: dmyStr_,
  fmtDate_: fmtDate_,
  ensureSheet_: ensureSheet_,
  getSettings: getSettings,
  getStaffRawData_: getStaffRawData_,
  invalidateStaffCache_: invalidateStaffCache_,
  invalidateDataCaches_: invalidateDataCaches_,
  getStaffList: getStaffList,
  findStaffRowFresh_: findStaffRowFresh_,
  rateHistorySheet_: rateHistorySheet_,
  recordRateHistory_: recordRateHistory_,
  getRateHistoryMap_: getRateHistoryMap_,
  rateForDate_: rateForDate_,
  initializeRateHistory: initializeRateHistory,
  addStaff: addStaff,
  setStaffActive: setStaffActive,
  deleteStaffPermanently: deleteStaffPermanently,
  editStaff: editStaff,
  editStaffBatch: editStaffBatch,
  setStaffPosition: setStaffPosition,
  moveStaff: moveStaff,
  repairStaffOrder: repairStaffOrder,
  repairStaffIds: repairStaffIds,
  findHoursConflicts_: findHoursConflicts_,
  submitWeeklyHours: submitWeeklyHours,
  submitDailyHours: submitDailyHours,
  getHoursEntries: getHoursEntries,
  findHoursRowByContent_: findHoursRowByContent_,
  updateHoursEntry: updateHoursEntry,
  deleteHoursEntry: deleteHoursEntry,
  getSummary: getSummary,
  submitRevenue: submitRevenue,
  getTakeinViewData: getTakeinViewData,
  submitTipPool: submitTipPool,
  getRecentTipPool: getRecentTipPool,
  sbCash_: sbCash_,
  submitCashEntry: submitCashEntry,
  getCashViewData: getCashViewData,
  findCashRowByContent_: findCashRowByContent_,
  updateCashEntry: updateCashEntry,
  deleteCashEntry: deleteCashEntry,
  setHolidayMarkCore_: setHolidayMarkCore_,
  setHolidayMark: setHolidayMark,
  getHolidaysViewData: getHolidaysViewData,
  submitAdvance: submitAdvance,
  requestsSheet_: requestsSheet_,
  addRequest: addRequest,
  getRequests: getRequests,
  setRequestStatus: setRequestStatus,
  deleteRequest: deleteRequest,
  costExclusionsSheet_: costExclusionsSheet_,
  buildExclusionSet_: buildExclusionSet_,
  getCostExclusions: getCostExclusions,
  addCostExclusion: addCostExclusion,
  removeCostExclusion: removeCostExclusion,
  buildPayPeriodWeeks_: buildPayPeriodWeeks_,
  getPayPeriodReport: getPayPeriodReport,
  setPayPeriodWeekValue: setPayPeriodWeekValue,
  getDashboardData: getDashboardData,
  countOpenRequests_: countOpenRequests_,
  distributeDailyTipsPence_: distributeDailyTipsPence_,
  getTipsViewData: getTipsViewData,
  parseBankingDate_: parseBankingDate_,
  existingDateSet_: existingDateSet_,
  backfillFromBankingDryRun: backfillFromBankingDryRun,
  backfillFromBankingSheet: backfillFromBankingSheet,
  backfillFromBankingSheet_: backfillFromBankingSheet_,
  runAllTests: runAllTests,
  fmtTestDate_: fmtTestDate_,
  hrsStr_: hrsStr_,
  getMonthEndDetail: getMonthEndDetail,
  exportMonthEndPdf: exportMonthEndPdf,
  getEntryStatus: getEntryStatus,
  exportPayPeriodPdf: exportPayPeriodPdf,
  applyGridEdits: applyGridEdits,
  getTipPoolEntries: getTipPoolEntries,
  findTipPoolRowByContent_: findTipPoolRowByContent_,
  updateTipPoolEntry: updateTipPoolEntry,
  deleteTipPoolEntry: deleteTipPoolEntry,
  setPayPeriodWeekValues: setPayPeriodWeekValues,
  submitMissedBankingDay: submitMissedBankingDay,
  findRevenueRowByContent_: findRevenueRowByContent_,
  updateRevenueEntry: updateRevenueEntry,
  deleteRevenueEntry: deleteRevenueEntry,
  testCashSync: testCashSync
};
})();
