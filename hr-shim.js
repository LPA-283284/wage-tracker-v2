// hr-shim.js — runs the HR backend (hr-core.js, = Code.gs v2.14) in the browser
// on top of Supabase. Every "sheet" the code reads is built from a Supabase
// table; every write the code makes is collected and sent to Supabase when the
// call finishes. The HR screen (hr.html = Index.html v2.5.12) is unchanged:
// its google.script.run calls are answered here.
(function () {
  'use strict';
  var SB_URL = 'https://gkdptxcyyrcfmndtazic.supabase.co';
  var SB_KEY = 'sb_publishable_sW5w5URiunt6MGcmqfuQrA_tBB6JVu4';
  var sb = window.supabase.createClient(SB_URL, SB_KEY);
  window.HR_SB = sb;
  var ME = null;

  // ---------------- small helpers ----------------
  function pad2(n) { return ('0' + n).slice(-2); }
  function d2s(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function s2d(s) {
    if (!s) return '';
    var p = String(s).slice(0, 10).split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }
  function toDateObj(v) {
    if (v === '' || v === null || v === undefined) return null;
    if (v instanceof Date) return isNaN(v) ? null : v;
    var s = String(v).trim();
    if (/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) return s2d(s);
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    var d = new Date(s);
    return isNaN(d) ? null : d;
  }
  function dateOrNull(v) { var d = toDateObj(v); return d ? d2s(d) : null; }
  function num(v, dflt) { var n = Number(v); return (v === '' || v === null || v === undefined || isNaN(n)) ? dflt : n; }
  function pence(v) { return Math.round(num(v, 0) * 100); }
  function pounds(p) { return Math.round(Number(p || 0)) / 100; }
  function sid(v) { return (v === '' || v === null || v === undefined) ? null : Number(v); }

  // ---------------- sheet <-> table mapping ----------------
  var staffNameById = {};
  var SPECS = {
    Staff: {
      t: 'staff', header: ['ID', 'Name', 'Department', 'Hourly Rate', 'Active', 'Tip Weight', 'Order', 'Holiday Entitlement', 'Start Date', 'End Date'],
      toRow: function (r) { return [r.id, r.name, r.department, num(r.hourly_rate, 0), !!r.is_active, num(r.tip_weight, 1), r.display_order, r.holiday_entitlement, s2d(r.start_date), s2d(r.end_date)]; },
      toDb: function (v) {
        return { name: String(v[1]).trim(), department: String(v[2]).trim(), hourly_rate: num(v[3], null),
          is_active: (v[4] === true || String(v[4]).toUpperCase() === 'TRUE'), tip_weight: num(v[5], 1),
          display_order: Math.round(num(v[6], 999)), holiday_entitlement: Math.round(num(v[7], 28)),
          start_date: dateOrNull(v[8]), end_date: dateOrNull(v[9]) };
      }
    },
    Rate_History: {
      t: 'rate_history', header: ['StaffID', 'Effective Date', 'Hourly Rate', 'Tip Weight'],
      toRow: function (r) { return [r.staff_id, s2d(r.effective_from), pounds(r.hourly_rate_pence), num(r.tip_weight, 1)]; },
      toDb: function (v) {
        var d = toDateObj(v[1]);
        // a timestamp that is not midnight takes effect the NEXT day (same rule as the sheet)
        if (d && (d.getHours() || d.getMinutes() || d.getSeconds())) d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
        return { staff_id: sid(v[0]), effective_from: d ? d2s(d) : null, hourly_rate_pence: pence(v[2]), tip_weight: num(v[3], 1) };
      }
    },
    Hours_Log: {
      t: 'hours_log', header: ['Date', 'Department', 'Staff', 'Hours', 'StaffID'],
      toRow: function (r) { return [s2d(r[1]), r[2], staffNameById[r[4]] || '', Number(r[3]), r[4]]; },
      toDb: function (v) { return { work_date: dateOrNull(v[0]), department: String(v[1]), hours: num(v[3], 0), staff_id: sid(v[4]) }; }
    },
    Revenue_Log: {
      t: 'revenue_log', header: ['Date', 'Net Sales', 'Service Charge', 'Note'],
      toRow: function (r) { return [s2d(r.revenue_date), pounds(r.net_sales_pence), pounds(r.service_charge_pence), r.notes || '']; },
      toDb: function (v) { return { revenue_date: dateOrNull(v[0]), net_sales_pence: pence(v[1]), service_charge_pence: pence(v[2]), notes: String(v[3] || '') || null }; },
      extraIns: { source: 'MANUAL' }
    },
    Tip_Pool_Log: {
      t: 'tip_pool_log', header: ['Date', 'Total Tips', 'Note'],
      toRow: function (r) { return [s2d(r.pool_date), pounds(r.pool_pence), r.notes || '']; },
      toDb: function (v) { return { pool_date: dateOrNull(v[0]), pool_pence: pence(v[1]), notes: String(v[2] || '') || null }; }
    },
    Cash_List_Log: {
      t: 'cash_list_log', header: ['Date', 'Description', 'Income', 'Expense', 'Source', 'Note'],
      toRow: function (r) { return [s2d(r.entry_date), r.description, pounds(r.income_pence), pounds(r.expense_pence), r.source === 'AUTOMATIC' ? 'Automatic' : 'Manual', r.notes || '']; },
      toDb: function (v, db) {
        var src = String(v[4]) === 'Automatic' ? 'AUTOMATIC' : ((db && db.source && db.source !== 'AUTOMATIC') ? db.source : 'MANUAL');
        return { entry_date: dateOrNull(v[0]), description: String(v[1] || '').trim(), income_pence: pence(v[2]), expense_pence: pence(v[3]), source: src, notes: String(v[5] || '') || null };
      }
    },
    Settings: {
      t: 'settings', readOnly: true, header: null,
      toRow: function (r) {
        var s = String(r.value || '').trim();
        var val = /%$/.test(s) ? Number(s.replace('%', '')) / 100 : (isNaN(Number(s)) ? s : Number(s));
        return [r.key, val];
      }
    },
    Holidays_Log: {
      t: 'holidays_log', header: ['Date', 'StaffID', 'Type', 'Note'],
      toRow: function (r) { return [s2d(r.holiday_date), r.staff_id, r.mark, r.notes || '']; },
      toDb: function (v) { return { holiday_date: dateOrNull(v[0]), staff_id: sid(v[1]), mark: String(v[2]).toUpperCase(), notes: String(v[3] || '') || null }; }
    },
    Advances_Log: {
      t: 'advances_log', header: ['Date', 'StaffID', 'Amount', 'Note'],
      toRow: function (r) { return [s2d(r.advance_date), r.staff_id, pounds(r.amount_pence), r.notes || '']; },
      toDb: function (v) { return { advance_date: dateOrNull(v[0]), staff_id: sid(v[1]), amount_pence: pence(v[2]), notes: String(v[3] || '') || null }; },
      extraIns: function () { return { method: 'CASH', source: 'HR', paid_by: ME ? ME.name : 'HR' }; }
    },
    PayPeriod_Manual_Log: {
      t: 'pay_period_manual_log', header: ['PeriodKey', 'StaffID', 'WeekIndex', 'Value'],
      toRow: function (r) {
        var ps = s2d(r.period_start); // period starting 26th of month M ends 25th of M+1 -> key "Y-(M+1)"
        var end = new Date(ps.getFullYear(), ps.getMonth() + 1, 25);
        return [end.getFullYear() + '-' + (end.getMonth() + 1), r.staff_id, r.week_index, r.marker ? r.marker : Number(r.hours)];
      },
      toDb: function (v) {
        var key = v[0] instanceof Date ? (v[0].getFullYear() + '-' + (v[0].getMonth() + 1)) : String(v[0]);
        var p = key.split('-');
        var ps = new Date(Number(p[0]), Number(p[1]) - 2, 26);
        var raw = String(v[3]).trim().toUpperCase();
        var mark = (raw === 'H' || raw === 'S') ? raw : null;
        return { period_start: d2s(ps), staff_id: sid(v[1]), week_index: Number(v[2]), hours: mark ? null : num(v[3], 0), marker: mark };
      }
    },
    Requests_Log: {
      t: 'requests_log', header: ['Date', 'StaffID', 'Request', 'Due Date', 'Status', 'Note', 'Done Date'],
      toRow: function (r) { return [s2d(r.request_date), r.staff_id == null ? '' : r.staff_id, r.body, s2d(r.due_date), r.is_done ? 'DONE' : 'OPEN', r.notes || '', r.done_at ? new Date(r.done_at) : '']; },
      toDb: function (v) {
        var done = String(v[4]).toUpperCase() === 'DONE';
        var da = v[6] instanceof Date ? v[6] : toDateObj(v[6]);
        return { request_date: dateOrNull(v[0]), staff_id: sid(v[1]), body: String(v[2] || ''), due_date: dateOrNull(v[3]),
          is_done: done, done_at: (done && da) ? da.toISOString() : null, notes: String(v[5] || '') || null };
      }
    },
    Cost_Exclusions_Log: {
      t: 'cost_exclusions', header: ['Month', 'Year', 'StaffID'],
      toRow: function (r) { return [r.month, r.year, r.staff_id]; },
      toDb: function (v) { return { month: Number(v[0]), year: Number(v[1]), staff_id: sid(v[2]) }; }
    }
  };
  var SHEET_ORDER = ['Staff', 'Rate_History', 'Hours_Log', 'Revenue_Log', 'Tip_Pool_Log', 'Cash_List_Log', 'Settings',
    'Holidays_Log', 'Advances_Log', 'PayPeriod_Manual_Log', 'Requests_Log', 'Cost_Exclusions_Log'];

  // ---------------- fake sheets ----------------
  function FakeSheet(name) { this.name = name; this.spec = SPECS[name]; this.rows = []; this.deleted = []; }
  FakeSheet.prototype.load = function (dbRows) {
    var spec = this.spec;
    this.rows = dbRows.map(function (r) {
      return { v: spec.toRow(r), id: Array.isArray(r) ? r[0] : (r.id === undefined ? null : r.id), db: r, dirty: false, isNew: false };
    });
    this.deleted = [];
  };
  FakeSheet.prototype.header = function () {
    if (this.spec.header) return this.spec.header.slice();
    return null;
  };
  FakeSheet.prototype.width = function () { return this.spec.header ? this.spec.header.length : 2; };
  FakeSheet.prototype.allValues = function () {
    var out = this.rows.map(function (r) { return r.v.slice(); });
    var h = this.header();
    return h ? [h].concat(out) : out;
  };
  FakeSheet.prototype.hasHeader = function () { return !!this.spec.header; };
  FakeSheet.prototype.getName = function () { return this.name; };
  FakeSheet.prototype.getDataRange = function () {
    var self = this;
    return { getValues: function () { return self.allValues(); }, getNumRows: function () { return self.getLastRow(); } };
  };
  FakeSheet.prototype.getLastRow = function () { return this.rows.length + (this.hasHeader() ? 1 : 0); };
  FakeSheet.prototype.getLastColumn = function () { return this.width(); };
  FakeSheet.prototype.getMaxRows = function () { return this.getLastRow(); };
  FakeSheet.prototype.idx = function (row) { return row - (this.hasHeader() ? 2 : 1); }; // sheet row -> rows[] index (-1 = header)
  FakeSheet.prototype.ensureRow = function (i) {
    if (this.spec.readOnly) throw new Error(this.name + ' is read-only here.');
    while (this.rows.length <= i) {
      var blank = []; for (var c = 0; c < this.width(); c++) blank.push('');
      this.rows.push({ v: blank, id: null, db: null, dirty: false, isNew: true });
    }
    return this.rows[i];
  };
  FakeSheet.prototype.appendRow = function (arr) {
    var r = this.ensureRow(this.rows.length);
    for (var c = 0; c < arr.length && c < this.width(); c++) r.v[c] = arr[c];
    return this;
  };
  FakeSheet.prototype.deleteRows = function (row, n) {
    var i = this.idx(row);
    if (i < 0) throw new Error('Cannot delete the header row.');
    var removed = this.rows.splice(i, n);
    var self = this;
    removed.forEach(function (r) { if (!r.isNew && r.id !== null) self.deleted.push(r); });
  };
  FakeSheet.prototype.deleteRow = function (row) { this.deleteRows(row, 1); };
  FakeSheet.prototype.getRange = function (row, col, nr, nc) { return new FakeRange(this, row, col, nr || 1, nc || 1); };
  FakeSheet.prototype.setFrozenRows = function () { return this; };
  FakeSheet.prototype.getSheetId = function () { return this.name; };

  function FakeRange(sh, row, col, nr, nc) { this.sh = sh; this.row = row; this.col = col; this.nr = nr; this.nc = nc; }
  FakeRange.prototype.getValues = function () {
    var all = this.sh.allValues(), out = [];
    for (var r = 0; r < this.nr; r++) {
      var src = all[this.row - 1 + r] || [], line = [];
      for (var c = 0; c < this.nc; c++) { var v = src[this.col - 1 + c]; line.push(v === undefined ? '' : v); }
      out.push(line);
    }
    return out;
  };
  FakeRange.prototype.getValue = function () { return this.getValues()[0][0]; };
  FakeRange.prototype.getDisplayValues = function () { return this.getValues().map(function (l) { return l.map(String); }); };
  FakeRange.prototype.setValues = function (vals) {
    for (var r = 0; r < vals.length; r++) {
      var i = this.sh.idx(this.row + r);
      if (i < 0) continue; // header writes are ignored (headers are fixed here)
      var rowObj = this.sh.ensureRow(i);
      for (var c = 0; c < vals[r].length; c++) rowObj.v[this.col - 1 + c] = vals[r][c];
      if (!rowObj.isNew) rowObj.dirty = true;
    }
    return this;
  };
  FakeRange.prototype.setValue = function (v) {
    var vals = [];
    for (var r = 0; r < this.nr; r++) { var l = []; for (var c = 0; c < this.nc; c++) l.push(v); vals.push(l); }
    return this.setValues(vals);
  };
  FakeRange.prototype.clearContent = function () { return this.setValue(''); };
  ['setNumberFormat', 'setFontWeight', 'setBackground', 'setFontColor', 'setHorizontalAlignment', 'setNumberFormats'].forEach(function (m) {
    FakeRange.prototype[m] = function () { return this; };
  });

  var SHEETS = {};
  SHEET_ORDER.forEach(function (n) { SHEETS[n] = new FakeSheet(n); });

  // ---------------- Apps Script globals ----------------
  var SS = {
    getSheetByName: function (n) { return SHEETS[n] || null; },
    insertSheet: function (n) { if (!SHEETS[n]) throw new Error('Unknown sheet ' + n); return SHEETS[n]; },
    getId: function () { return 'supabase'; }
  };
  window.SpreadsheetApp = { getActive: function () { return SS; }, getActiveSpreadsheet: function () { return SS; }, openById: function () { throw new Error('Not available in the web version.'); }, flush: function () {} };
  window.Session = { getScriptTimeZone: function () { return 'Europe/London'; } };
  window.LockService = { getScriptLock: function () { return { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {} }; } };
  window.CacheService = { getScriptCache: function () { return { get: function () { return null; }, put: function () {}, remove: function () {}, removeAll: function () {} }; } };
  window.PropertiesService = { getScriptProperties: function () { return { getProperty: function () { return null; } }; } };
  window.Logger = { log: function () { console.log.apply(console, arguments); } };
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  window.Utilities = {
    formatDate: function (d, tz, fmt) {
      d = (d instanceof Date) ? d : new Date(d);
      return fmt.replace(/yyyy|MMMM|MMM|MM|dd|HH|mm|ss|EEE|d/g, function (t) {
        switch (t) {
          case 'yyyy': return String(d.getFullYear());
          case 'MMMM': return ['January','February','March','April','May','June','July','August','September','October','November','December'][d.getMonth()];
          case 'MMM': return MON[d.getMonth()];
          case 'MM': return pad2(d.getMonth() + 1);
          case 'dd': return pad2(d.getDate());
          case 'd': return String(d.getDate());
          case 'HH': return pad2(d.getHours());
          case 'mm': return pad2(d.getMinutes());
          case 'ss': return pad2(d.getSeconds());
          case 'EEE': return DOW[d.getDay()];
        }
        return t;
      });
    },
    sleep: function () {}
  };

  // sbCash_ inside the core: the Supabase "bridge" calls become sheet operations
  window.HRSHIM_sbCash = function (fn, p) {
    if (fn === 'hr_cash_write' || fn === 'hr_advance_add' || fn === 'hr_ppml_write') return { ok: true }; // the sheet write that follows does it
    if (fn === 'hr_cash_list') {
      return { ok: true, rows: SHEETS.Cash_List_Log.rows.map(function (r, i) {
        return { id: r.id !== null ? r.id : ('new' + i), date: d2s(toDateObj(r.v[0])), description: r.v[1], income: Number(r.v[2]) || 0,
          expense: Number(r.v[3]) || 0, source: r.v[4], notes: r.v[5] || '' };
      }) };
    }
    if (fn === 'hr_advances_list') {
      var from = s2d(p.p_from), to = s2d(p.p_to);
      return { ok: true, rows: SHEETS.Advances_Log.rows.filter(function (r) {
        var d = toDateObj(r.v[0]); return d && d >= from && d <= to;
      }).map(function (r) { return { date: d2s(toDateObj(r.v[0])), staff_id: r.v[1], amount: Number(r.v[2]) || 0, notes: r.v[3] || '' }; }) };
    }
    if (fn === 'hr_ppml_list') {
      var ps = s2d(p.p_period_start), end = new Date(ps.getFullYear(), ps.getMonth() + 1, 25);
      var key = end.getFullYear() + '-' + (end.getMonth() + 1);
      return { ok: true, rows: SHEETS.PayPeriod_Manual_Log.rows.filter(function (r) {
        var k = r.v[0] instanceof Date ? (r.v[0].getFullYear() + '-' + (r.v[0].getMonth() + 1)) : String(r.v[0]);
        return k === key;
      }).map(function (r) { return { staff_id: r.v[1], week_index: r.v[2], value: String(r.v[3]) }; }) };
    }
    throw new Error('Unknown Supabase call ' + fn);
  };

  // ---------------- load / flush ----------------
  var lastLoad = 0;
  async function loadAll() {
    var res = await sb.rpc('hr_load');
    if (res.error) throw new Error(res.error.message);
    if (!res.data || !res.data.ok) throw new Error((res.data && res.data.error) || 'Could not load HR data');
    var d = res.data;
    staffNameById = {};
    d.staff.forEach(function (s) { staffNameById[s.id] = s.name; });
    SHEETS.Staff.load(d.staff);
    SHEETS.Rate_History.load(d.rate_history);
    SHEETS.Hours_Log.load(d.hours_log);
    SHEETS.Revenue_Log.load(d.revenue_log);
    SHEETS.Tip_Pool_Log.load(d.tip_pool_log);
    SHEETS.Cash_List_Log.load(d.cash_list_log);
    SHEETS.Settings.load(d.settings);
    SHEETS.Holidays_Log.load(d.holidays_log);
    SHEETS.Advances_Log.load(d.advances_log);
    SHEETS.PayPeriod_Manual_Log.load(d.pay_period_manual_log);
    SHEETS.Requests_Log.load(d.requests_log);
    SHEETS.Cost_Exclusions_Log.load(d.cost_exclusions);
    lastLoad = Date.now();
  }

  function pendingCount() {
    var n = 0;
    SHEET_ORDER.forEach(function (name) {
      var sh = SHEETS[name];
      n += sh.deleted.length;
      sh.rows.forEach(function (r) { if (r.isNew || r.dirty) n++; });
    });
    return n;
  }

  async function check(res) { if (res.error) throw new Error(res.error.message + (res.error.details ? ' (' + res.error.details + ')' : '')); return res.data; }

  async function flush() {
    var staffIdRemap = {};
    for (var k = 0; k < SHEET_ORDER.length; k++) {
      var name = SHEET_ORDER[k], sh = SHEETS[name], spec = sh.spec;
      if (spec.readOnly) continue;
      // 1) deletes
      if (sh.deleted.length) {
        var ids = sh.deleted.map(function (r) { return r.id; });
        if (spec.t === 'staff') await check(await sb.from('staff').update({ is_deleted: true, is_active: false }).in('id', ids));
        else if (spec.t === 'advances_log') await check(await sb.from('advances_log').update({ is_deleted: true }).in('id', ids));
        else await check(await sb.from(spec.t).delete().in('id', ids));
        sh.deleted = [];
      }
      // 2) updates (only rows the code touched)
      var dirty = sh.rows.filter(function (r) { return r.dirty && !r.isNew; });
      for (var i = 0; i < dirty.length; i += 8) {
        await Promise.all(dirty.slice(i, i + 8).map(async function (r) {
          var vals = spec.toDb(r.v, r.db);
          var upd = await sb.from(spec.t).update(vals).eq('id', r.id).select('id');
          await check(upd);
          if (!upd.data || !upd.data.length) throw new Error('A row was changed elsewhere — the screen will reload.');
          r.dirty = false;
        }));
      }
      // 3) inserts (in order, ids written back)
      var fresh = sh.rows.filter(function (r) { return r.isNew; });
      if (fresh.length) {
        var payload = fresh.map(function (r) {
          var o = spec.toDb(r.v, null);
          var extra = typeof spec.extraIns === 'function' ? spec.extraIns() : (spec.extraIns || {});
          Object.keys(extra).forEach(function (x) { if (o[x] === undefined) o[x] = extra[x]; });
          if (o.staff_id !== undefined && staffIdRemap[o.staff_id] !== undefined) o.staff_id = staffIdRemap[o.staff_id];
          return o;
        });
        var ins = await sb.from(spec.t).insert(payload).select('id');
        await check(ins);
        fresh.forEach(function (r, j) {
          var newId = ins.data && ins.data[j] ? ins.data[j].id : null;
          if (spec.t === 'staff' && newId !== null && Number(r.v[0]) !== newId) { staffIdRemap[Number(r.v[0])] = newId; r.v[0] = newId; }
          if (spec.t === 'staff') staffNameById[newId] = r.v[1];
          r.id = newId; r.isNew = false; r.db = null;
        });
      }
    }
  }

  // ---------------- google.script.run ----------------
  var chain = Promise.resolve();
  var OVERRIDES = {};
  window.HRSHIM_override = function (name, fn) { OVERRIDES[name] = fn; };

  function plain(x) { return x === undefined ? undefined : JSON.parse(JSON.stringify(x)); }

  function run(fnName, args, ok, fail) {
    chain = chain.then(async function () {
      var result;
      try {
        if (!(await BOOT)) return;
        if (!lastLoad || Date.now() - lastLoad > 60000) await loadAll();
        var fn = OVERRIDES[fnName] || HRCORE[fnName];
        if (typeof fn !== 'function') throw new Error('Unknown function: ' + fnName);
        result = await fn.apply(null, plain(args));
        if (pendingCount() > 0) await flush();
      } catch (e) {
        console.error(fnName, e);
        try { await loadAll(); } catch (e2) { console.error(e2); }
        if (fail) fail(e instanceof Error ? e : new Error(String(e)));
        return;
      }
      if (ok) {
        try { ok(plain(result)); } catch (e) { console.error(e); }
      }
    });
    return chain;
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, prop) {
        if (prop === 'withSuccessHandler') return function (h) { return runner(h, fail); };
        if (prop === 'withFailureHandler') return function (h) { return runner(ok, h); };
        if (prop === 'withUserObject') return function () { return runner(ok, fail); };
        return function () { return run(prop, Array.prototype.slice.call(arguments), ok, fail); };
      }
    });
  }
  window.google = { script: { run: runner(null, null) } };

  // ---------------- printable PDFs (replace the Google Docs export) ----------------
  function escH(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function hrsTxt(n) { n = Number(n) || 0; return n === Math.floor(n) ? String(n) : n.toFixed(2); }
  function gbpTxt(n) { return '£' + (Number(n) || 0).toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function dmyTxt(s) { var p = String(s || '').split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : String(s || ''); }
  function printDoc(title, body) {
    var html = '<!doctype html><html><head><meta charset="utf-8"><title>' + escH(title) + '</title><style>' +
      'body{font-family:Arial,sans-serif;margin:24px;color:#111}h1{font-size:18px;color:#6D1B2F;margin:0 0 12px}' +
      '.grid{display:flex;flex-wrap:wrap;gap:16px;align-items:flex-start}table{border-collapse:collapse;font-size:12px;min-width:260px}' +
      'th,td{border:1px solid #ccc;padding:4px 6px;text-align:left}th{background:#6D1B2F;color:#fff}td.n{text-align:right}' +
      'tr.t td{font-weight:bold;background:#E2EFDA}.g{font-weight:bold;margin-top:12px}@media print{body{margin:8mm}}</style></head><body>' +
      '<h1>' + escH(title) + '</h1>' + body + '<script>setTimeout(function(){print()},400)<\/script></body></html>';
    var url = URL.createObjectURL(new Blob([html], { type: 'text/html' }));
    return { fileName: title + '.pdf', downloadUrl: url, viewUrl: url };
  }
  OVERRIDES.exportPayPeriodPdf = function (month, year) {
    var d = HRCORE.getPayPeriodReport(month, year);
    var groups = { FLOOR: [], KITCHEN: [], CLEANING: [], OTHER: [] };
    d.rows.forEach(function (r) { groups[r.dept === 'FOH' ? 'FLOOR' : r.dept === 'BOH' ? 'KITCHEN' : r.dept === 'CLEANING' ? 'CLEANING' : 'OTHER'].push(r); });
    var gh = 0, ga = 0, body = '<div class="grid">';
    Object.keys(groups).forEach(function (g) {
      var rows = groups[g]; if (!rows.length) return;
      var th = 0, ta = 0;
      body += '<table><tr><th colspan="6">' + g + '</th></tr><tr><th>Name</th><th>In</th><th>Out</th><th>Hours</th><th>Advance</th><th>Notes</th></tr>';
      rows.forEach(function (r) {
        var adv = Number(r.advances) || 0; th += Number(r.totalHours) || 0; ta += adv;
        var sWeeks = (r.weeklyHours || []).filter(function (w) { return w === 'S'; }).length;
        var note = sWeeks > 0 ? 'Sick ' + sWeeks + 'wk' : (Number(r.sickDays) > 0 ? 'Sick ' + r.sickDays + 'd' : '');
        body += '<tr><td>' + escH(r.name) + '</td><td>' + (r.startDate ? dmyTxt(r.startDate) : '') + '</td><td>' + (r.endDate ? dmyTxt(r.endDate) : '') +
          '</td><td class="n">' + hrsTxt(r.totalHours) + '</td><td class="n">' + (adv > 0 ? gbpTxt(adv) : '') + '</td><td>' + note + '</td></tr>';
      });
      body += '<tr class="t"><td>TOTAL</td><td></td><td></td><td class="n">' + hrsTxt(th) + '</td><td class="n">' + (ta > 0 ? gbpTxt(ta) : '') + '</td><td></td></tr></table>';
      gh += th; ga += ta;
    });
    body += '</div><p class="g">GRAND TOTAL: ' + hrsTxt(gh) + ' h' + (ga > 0 ? ' | Advances: ' + gbpTxt(ga) : '') + '</p>';
    return printDoc('Pay Period ' + dmyTxt(d.periodStart) + ' – ' + dmyTxt(d.periodEnd), body);
  };
  OVERRIDES.exportMonthEndPdf = function (month, year) {
    var d = HRCORE.getMonthEndDetail(month, year);
    var body = '<div class="grid">';
    ['FLOOR', 'KITCHEN', 'CLEANING', 'OTHER'].forEach(function (g) {
      var rows = d.groups[g]; if (!rows || !rows.length) return;
      body += '<table><tr><th colspan="2">' + g + '</th></tr>';
      rows.forEach(function (r) { body += '<tr><td>' + escH(r.name) + '</td><td class="n">' + hrsTxt(r.hours) + '</td></tr>'; });
      body += '<tr class="t"><td>TOTAL</td><td class="n">' + hrsTxt(d.groupTotals[g]) + '</td></tr></table>';
    });
    body += '</div><p class="g">GRAND TOTAL: ' + hrsTxt(d.grandTotal) + ' h</p>';
    return printDoc('End of Month ' + d.label, body);
  };

  // ---------------- sign-in gate ----------------
  var BOOT = (async function () {
    var s = await sb.auth.getSession();
    if (!s.data.session) { location.href = 'index.html'; return false; }
    var p = await sb.rpc('my_profile');
    if (p.data && p.data.ok) ME = p.data.me;
    try { await loadAll(); } catch (e) {
      document.body.insertAdjacentHTML('afterbegin', '<div style="background:#9C0006;color:#fff;padding:12px 20px;font-weight:600">' +
        escH(e.message || e) + ' — <a href="index.html" style="color:#fff">back to TeamOS</a></div>');
      return false;
    }
    var who = document.getElementById('hrWho'); if (who && ME) who.textContent = ME.name; var av = document.getElementById('hrAv'); if (av && ME) av.textContent = (ME.name || '?').trim().charAt(0).toUpperCase();
    return true;
  })();
  window.HRSHIM_ready = BOOT;
  window.HRSHIM_signOut = async function () { await sb.auth.signOut(); location.href = 'index.html'; };
})();
