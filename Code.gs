/**
 * AVA Passenger Transport — Cash Book (v3)
 * Container-bound Google Apps Script web app.
 * Data: this Spreadsheet + "AVA Cashbook" folder on your Drive.
 * No third-party server. UAE VAT 5% ready. Audit log + health check included.
 */

var ROOT_FOLDER_NAME = 'AVA Cashbook';
var TX_SHEET  = 'Transactions';
var SET_SHEET = 'Settings';
var SUM_SHEET = 'Monthly Summary';
var VAT_SHEET = 'VAT Quarters';
var LOG_SHEET = 'Audit Log';
var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

var TX_HEADERS = [
  'ID','Date','Month','Type','Account','Category','Sub-category',
  'Vehicle','Driver','Route','Trip No','Odometer (km)',
  'Description','Reference','Amount (AED)','VAT Rate %','VAT Amount (AED)','Net Amount (AED)',
  'Transfer To','Receipt Name','Receipt Link','Receipt File ID','Created','Updated','Updated By'
];

var DEFAULT_CATS = {
  'Fuel':            ['Diesel','Petrol','Engine oil'],
  'Maintenance':     ['Service','Tyres','Spare parts','Repairs','Battery','AC'],
  'Vehicle expenses':['Salik','Parking','Fines','Insurance','Registration','Mulkiya','Permit'],
  'Salary':          ['Driver','Helper','Office staff','Bonus','Overtime'],
  'Office':          ['Stationery','Utilities','Internet','Phone','Software'],
  'Sales':           ['Passenger fares','Contract hire','School route','Corporate'],
  'Rent':            ['Office','Parking yard'],
  'Bank charges':    [],
  'Other':           []
};

var ALLOWED_MIME = {'image/jpeg':'.jpg','image/png':'.png','image/webp':'.webp','application/pdf':'.pdf'};
var MAX_RECEIPT_BYTES = 5 * 1024 * 1024;

/* ============================== WEB APP ============================== */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('AVA Cash Book')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('AVA Cashbook')
    .addItem('Rebuild monthly summary', 'rebuildSummary')
    .addItem('Rebuild VAT quarters',    'rebuildVatQuarters')
    .addItem('Run health check',        'healthCheckUi')
    .addItem('Backup now to Drive',     'backupNow')
    .addItem('Turn ON monthly backup',  'installBackupTrigger')
    .addItem('Turn OFF monthly backup', 'removeBackupTrigger')
    .addItem('Open receipts folder',    'openReceiptsFolder')
    .addToUi();
}

/* ============================== SETUP ============================== */
function ensureSetup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet(ss, TX_SHEET, TX_HEADERS);
  ensureSheet(ss, SET_SHEET, ['Key','Value']);
  ensureSheet(ss, SUM_SHEET, ['Month','Money In (AED)','Money Out (AED)','Net (AED)',
    'VAT Collected (AED)','VAT Paid (AED)','VAT Payable (AED)',
    'Opening Bank','Closing Bank','Opening Petty','Closing Petty','Entries']);
  ensureSheet(ss, VAT_SHEET, ['Quarter','Output VAT (AED)','Input VAT (AED)','Payable (AED)','Taxable Sales','Taxable Purchases']);
  ensureSheet(ss, LOG_SHEET, ['When','User','Action','Entry ID','Details']);
  var tx = ss.getSheetByName(TX_SHEET);
  tx.getRange('B1:C1000').setNumberFormat('@');     // dates/months as text (exact match)
  ['O','Q','R'].forEach(function(c){ tx.getRange(c + '2:' + c).setNumberFormat('#,##0.00'); });
  tx.setFrozenRows(1);
  if (!getSetting('cats')) {
    saveSettingsRaw('cats',    DEFAULT_CATS);
    saveSettingsRaw('ob',      {Bank:0, Petty:0});
    saveSettingsRaw('veh',     []);
    saveSettingsRaw('drivers', []);
    saveSettingsRaw('routes',  []);
    saveSettingsRaw('quick',   []);
    saveSettingsRaw('company', {name:'AVA Passenger Transport', trn:'', address:''});
  }
}

function ensureSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var need = headers.length;
  if (sh.getLastColumn() < need) sh.insertColumnsAfter(sh.getLastColumn() || 1, need - (sh.getLastColumn() || 1));
  var cur = sh.getRange(1, 1, 1, need).getValues()[0];
  var same = true;
  for (var i = 0; i < need; i++) if (String(cur[i]) !== headers[i]) { same = false; break; }
  if (!same) sh.getRange(1, 1, 1, need).setValues([headers]);
  return sh;
}

/* ============================== SETTINGS ============================== */
function getSetting(key) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SET_SHEET);
  if (!sh) return null;
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === key) {
      try { return JSON.parse(vals[i][1]); } catch (e) { return null; }
    }
  }
  return null;
}
function saveSettingsRaw(key, obj) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SET_SHEET);
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) {
        sh.getRange(i + 2, 2).setValue(JSON.stringify(obj));
        return;
      }
    }
  }
  sh.appendRow([key, JSON.stringify(obj)]);
}
function saveSettings(s) {
  return withLock(function () {
    ensureSetup();
    ['ob','cats','veh','drivers','routes','quick','company'].forEach(function (k) {
      if (s && s[k] !== undefined) saveSettingsRaw(k, s[k]);
    });
    logAudit('settings.save', '', {keys:Object.keys(s || {})});
    return true;
  });
}

/* ============================== HELPERS ============================== */
function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function userEmail() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function logAudit(action, id, details) {
  try {
    SpreadsheetApp.getActive().getSheetByName(LOG_SHEET)
      .appendRow([new Date(), userEmail(), action, id || '', JSON.stringify(details || {}).slice(0, 500)]);
  } catch (e) {}
}

/* ============================== BOOTSTRAP ============================== */
function getBootstrap() {
  ensureSetup();
  return {
    tx: readTx(),
    ob:      getSetting('ob')      || {Bank:0, Petty:0},
    cats:    getSetting('cats')    || {},
    veh:     getSetting('veh')     || [],
    drivers: getSetting('drivers') || [],
    routes:  getSetting('routes')  || [],
    quick:   getSetting('quick')   || [],
    company: getSetting('company') || {},
    backupOn: backupTriggerInstalled()
  };
}

/* ============================== ENTRY NORMALISATION ============================== */
function normalizeEntry(e) {
  e = e || {};
  var o = {};
  o.date = String(e.date || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date)) throw new Error('Date must be yyyy-mm-dd.');
  o.amt = r2(e.amt);
  if (!(o.amt > 0)) throw new Error('Amount must be greater than 0.');
  o.type = e.type === 'in' ? 'In' : (e.type === 'xfer' ? 'Transfer' : 'Out');
  o.acc  = e.acc === 'Petty' ? 'Petty' : 'Bank';
  o.desc = String(e.desc || '').slice(0, 300);
  o.ref  = String(e.ref  || '').slice(0, 100);
  if (o.type === 'Transfer') {
    o.toAcc = o.acc === 'Bank' ? 'Petty' : 'Bank';
    o.cat = 'Transfer'; o.sub = ''; o.veh = ''; o.driver = ''; o.route = ''; o.trip = ''; o.odo = 0;
    o.vatRate = 0; o.vat = 0; o.net = o.amt;
  } else {
    o.toAcc = '';
    o.cat = String(e.cat || ''); o.sub = String(e.sub || '');
    o.veh = String(e.veh || ''); o.driver = String(e.driver || '').slice(0, 60);
    o.route = String(e.route || '').slice(0, 60); o.trip = String(e.trip || '').slice(0, 30);
    o.odo = Number(e.odo) > 0 ? Math.round(Number(e.odo)) : 0;
    o.vatRate = Number(e.vat) === 5 ? 5 : 0;
    o.vat = o.vatRate ? r2(o.amt * 5 / 105) : 0;
    o.net = r2(o.amt - o.vat);
  }
  o.month = o.date.slice(0, 7);
  return o;
}

function newId(date) {
  var stamp = date.replace(/-/g, '').slice(2);
  var rand  = Utilities.getUuid().replace(/-/g, '').slice(0, 5).toUpperCase();
  return 'AVA-' + stamp + '-' + rand;
}

function txRow(id, e, rc, created, updated, updatedBy) {
  return [id, e.date, e.month, e.type, e.acc, e.cat || '', e.sub || '',
    e.veh || '', e.driver || '', e.route || '', e.trip || '', e.odo || 0,
    e.desc || '', e.ref || '', e.amt, e.vatRate || 0, e.vat || 0, e.net,
    e.toAcc || '', rc ? rc.name : '', rc ? rc.link : '', rc ? rc.fileId : '',
    created || new Date(), updated || new Date(),
    updatedBy || userEmail()];
}
function txFromRow(r) {
  return {
    id: r[0],
    date: r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1]).slice(0, 10),
    month: String(r[2] || '').slice(0, 7), type: r[3], acc: r[4], cat: r[5], sub: r[6],
    veh: r[7], driver: r[8], route: r[9], trip: r[10], odo: Number(r[11]) || 0,
    desc: r[12], ref: r[13], amt: Number(r[14]) || 0, vatRate: Number(r[15]) || 0,
    vat: Number(r[16]) || 0, net: Number(r[17]) || 0, toAcc: r[18],
    rcptName: r[19], rcptLink: r[20], rcptFileId: r[21],
    created: r[22] instanceof Date ? r[22].toISOString() : '',
    updated: r[23] instanceof Date ? r[23].toISOString() : '',
    updatedBy: r[24] || ''
  };
}
function readTx() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TX_HEADERS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(txFromRow);
}
function findRowById(id) {
  var sh = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

/* ============================== RECEIPTS ============================== */
function saveReceipt(id, e, receipt) {
  if (!receipt || !receipt.data) return null;
  var mime = receipt.mimeType || 'image/jpeg';
  if (!ALLOWED_MIME[mime]) throw new Error('Receipt must be JPG, PNG, WebP or PDF.');
  var bytes = Utilities.base64Decode(receipt.data);
  if (bytes.length > MAX_RECEIPT_BYTES) throw new Error('Receipt is larger than 5 MB.');
  var ext = ALLOWED_MIME[mime];
  var safe = String(e.desc || 'receipt').replace(/[^\w\s-]/g, '').replace(/\s+/g, '-').slice(0, 40);
  var name = e.date + '_' + id + '_' + (safe || 'receipt') + ext;
  var folder = receiptFolder(e.date, e.vat > 0);
  var file = folder.createFile(Utilities.newBlob(bytes, mime, name));
  file.setDescription('AVA Cashbook ' + id + ' | ' + e.date + ' | ' + (e.vat > 0 ? 'VAT' : 'Non-VAT'));
  return {name: name, link: file.getUrl(), fileId: file.getId()};
}
function trashReceipt(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
}

/* ============================== ADD / UPDATE / DELETE ============================== */
function addEntry(e, receipt) {
  return withLock(function () {
    ensureSetup();
    e = normalizeEntry(e);
    var id = newId(e.date);
    var rc = saveReceipt(id, e, receipt || null);
    var row = txRow(id, e, rc);
    SpreadsheetApp.getActive().getSheetByName(TX_SHEET).appendRow(row);
    rebuildSummary(); rebuildVatQuarters();
    logAudit('entry.add', id, {amt:e.amt, type:e.type, cat:e.cat});
    return txFromRow(row);
  });
}
function updateEntry(id, e, receipt) {
  return withLock(function () {
    ensureSetup();
    var row = findRowById(id);
    if (!row) throw new Error('Entry not found.');
    var sh = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);
    var old = txFromRow(sh.getRange(row, 1, 1, TX_HEADERS.length).getValues()[0]);
    e = normalizeEntry(e);
    var rc = old.rcptFileId ? {name: old.rcptName, link: old.rcptLink, fileId: old.rcptFileId} : null;
    if (receipt && receipt.data) {
      trashReceipt(old.rcptFileId);
      rc = saveReceipt(String(id), e, receipt);
    } else if (rc && (old.month !== e.month || (old.vat > 0) !== (e.vat > 0))) {
      // Transactional folder move: only update the stored link after the move succeeds
      try {
        var file = DriveApp.getFileById(rc.fileId);
        var newFolder = receiptFolder(e.date, e.vat > 0);
        var parents = file.getParents();
        if (parents.hasNext()) parents.next().removeFile(file);
        newFolder.addFile(file);
        rc.link = file.getUrl();
      } catch (err) { /* rollback: keep old link and old folder */ }
    }
    sh.getRange(row, 1, 1, TX_HEADERS.length).setValues([txRow(id, e, rc, old.created, new Date())]);
    rebuildSummary(); rebuildVatQuarters();
    logAudit('entry.update', id, {from:{amt:old.amt, type:old.type}, to:{amt:e.amt, type:e.type}});
    return txFromRow(sh.getRange(row, 1, 1, TX_HEADERS.length).getValues()[0]);
  });
}
function deleteEntry(id) {
  return withLock(function () {
    ensureSetup();
    var row = findRowById(id);
    if (!row) throw new Error('Entry not found.');
    var sh = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);
    var old = txFromRow(sh.getRange(row, 1, 1, TX_HEADERS.length).getValues()[0]);
    trashReceipt(old.rcptFileId);
    sh.deleteRow(row);
    rebuildSummary(); rebuildVatQuarters();
    logAudit('entry.delete', id, {amt:old.amt, type:old.type});
    return true;
  });
}

/* ============================== QUICK ENTRIES ============================== */
function postQuick(i) {
  return withLock(function () {
    ensureSetup();
    var q = (getSetting('quick') || [])[i];
    if (!q) throw new Error('Quick entry not found.');
    var ref = 'quick:' + (q.id || i);
    var cm  = todayStr().slice(0, 7);
    // Idempotence: never post the same monthly quick twice in the same month
    if (q.monthly) {
      var dup = readTx().some(function (t) { return t.ref === ref && t.date.indexOf(cm) === 0; });
      if (dup) throw new Error('Already posted this month.');
    }
    var e = normalizeEntry({
      date: todayStr(), type: q.type || 'out', acc: q.acc || 'Bank', amt: q.amt,
      cat: q.cat, sub: q.sub, veh: q.veh, driver: q.driver,
      desc: q.label, ref: ref, vat: q.vat ? 5 : 0
    });
    var id = newId(e.date);
    SpreadsheetApp.getActive().getSheetByName(TX_SHEET).appendRow(txRow(id, e, null));
    rebuildSummary(); rebuildVatQuarters();
    logAudit('entry.quick', id, {quick: q.label});
    return txFromRow(SpreadsheetApp.getActive().getSheetByName(TX_SHEET)
      .getRange(findRowById(id), 1, 1, TX_HEADERS.length).getValues()[0]);
  });
}

/* ============================== CSV EXPORT ============================== */
function csvCell(v) {
  v = String(v == null ? '' : v);
  return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}
function exportCsv(filter) {
  ensureSetup();
  filter = filter || {};
  var tx = readTx().filter(function (t) {
    if (filter.period && String(t.date).indexOf(filter.period) !== 0) return false;
    if (filter.quarter) {
      var m = Number(t.date.slice(5, 7)), y = t.date.slice(0, 4);
      var q = Math.ceil(m / 3);
      if (filter.quarter !== y + '-Q' + q) return false;
    }
    if (filter.vehicle && t.veh    !== filter.vehicle) return false;
    if (filter.driver  && t.driver !== filter.driver)  return false;
    return true;
  });
  tx.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  var lines = [TX_HEADERS.join(',')];
  tx.forEach(function (t) {
    lines.push([t.id, t.date, t.month, t.type, t.acc, t.cat, t.sub, t.veh, t.driver, t.route, t.trip, t.odo,
      t.desc, t.ref, t.amt, t.vatRate, t.vat, t.net, t.toAcc, t.rcptName, t.rcptLink, t.rcptFileId,
      t.created, t.updated, t.updatedBy].map(csvCell).join(','));
  });
  return lines.join('\r\n');
}

/* ============================== DRIVE FOLDERS ============================== */
function getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function receiptsRoot() {
  return getOrCreateFolder(getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME), 'Receipts');
}
function getReceiptsUrl() {
  return receiptsRoot().getUrl();
}
function receiptFolder(dateStr, hasVat) {
  var ym = String(dateStr).slice(0, 7), y = ym.slice(0, 4), m = Number(ym.slice(5, 7));
  var monthName = (m >= 1 && m <= 12) ? MONTHS[m - 1] : ym;
  var monthFolder = getOrCreateFolder(getOrCreateFolder(receiptsRoot(), y), ym + ' ' + monthName);
  return getOrCreateFolder(monthFolder, hasVat ? 'VAT' : 'Non-VAT');
}

/* ============================== SUMMARIES ============================== */
/* Movement applied to an account balance. Transfers move money between
   accounts without hitting money-in/out. */
function accountDelta(t, acc) {
  var d = 0;
  if (t.type === 'In'  && t.acc === acc) d += t.amt;
  if (t.type === 'Out' && t.acc === acc) d -= t.amt;
  if (t.type === 'Transfer') {
    if (t.acc   === acc) d -= t.amt;   // money leaves this account
    if (t.toAcc === acc) d += t.amt;   // money arrives in this account
  }
  return d;
}

function rebuildSummary() {
  ensureSetup();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SUM_SHEET);
  var tx = readTx().sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  var ob = getSetting('ob') || {Bank:0, Petty:0};

  // Clear old body
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 12).clearContent();
  if (!tx.length) return;

  // Collect months present
  var months = [], seen = {};
  tx.forEach(function (t) {
    if (t.month && !seen[t.month]) { seen[t.month] = true; months.push(t.month); }
  });
  months.sort();

  var rows = [];
  var runBank = Number(ob.Bank) || 0, runPetty = Number(ob.Petty) || 0;
  var mi = 0;
  months.forEach(function (m) {
    var openBank = runBank, openPetty = runPetty;
    var inn = 0, out = 0, vatIn = 0, vatOut = 0, n = 0;
    while (mi < tx.length && tx[mi].month <= m) {
      var t = tx[mi];
      if (t.month === m) {
        runBank  = r2(runBank  + accountDelta(t, 'Bank'));
        runPetty = r2(runPetty + accountDelta(t, 'Petty'));
        if (t.type === 'In')  { inn += t.amt; vatIn  += t.vat || 0; }
        if (t.type === 'Out') { out += t.amt; vatOut += t.vat || 0; }
        n++;
      } else {
        // entry belongs to an earlier month already summarised — apply to running balances only
        runBank  = r2(runBank  + accountDelta(t, 'Bank'));
        runPetty = r2(runPetty + accountDelta(t, 'Petty'));
      }
      mi++;
    }
    rows.push([m, r2(inn), r2(out), r2(inn - out),
      r2(vatIn), r2(vatOut), r2(vatIn - vatOut),
      openBank, r2(runBank), openPetty, r2(runPetty), n]);
  });
  sh.getRange(2, 1, rows.length, 12).setValues(rows);
}

function rebuildVatQuarters() {
  ensureSetup();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(VAT_SHEET);
  var tx = readTx();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();

  var q = {};
  tx.forEach(function (t) {
    if (!t.vat || t.vat <= 0) return;
    var m = Number(t.date.slice(5, 7));
    var key = t.date.slice(0, 4) + '-Q' + Math.ceil(m / 3);
    if (!q[key]) q[key] = {out: 0, inp: 0, sales: 0, purch: 0};
    if (t.type === 'In')  { q[key].out   += t.vat; q[key].sales += t.net; }
    if (t.type === 'Out') { q[key].inp   += t.vat; q[key].purch += t.net; }
  });
  var keys = Object.keys(q).sort();
  if (!keys.length) return;
  var rows = keys.map(function (k) {
    return [k, r2(q[k].out), r2(q[k].inp), r2(q[k].out - q[k].inp), r2(q[k].sales), r2(q[k].purch)];
  });
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

/* ============================== HEALTH CHECK ============================== */
function healthCheck() {
  ensureSetup();
  var tx = readTx();
  var problems = [];
  var missingReceipts = tx.filter(function (t) { return t.type === 'Out' && !t.rcptFileId; });
  var orphaned = 0;
  tx.forEach(function (t) {
    if (t.rcptFileId) {
      try { DriveApp.getFileById(t.rcptFileId); }
      catch (e) { orphaned++; }
    }
  });
  if (missingReceipts.length) problems.push(missingReceipts.length + ' money-out entries have no receipt attached.');
  if (orphaned) problems.push(orphaned + ' receipt file links point to files that no longer exist in Drive.');
  return {
    entries: tx.length,
    missingReceipts: missingReceipts.length,
    missingList: missingReceipts.slice(0, 20).map(function (t) { return t.date + ' · ' + t.desc + ' · AED ' + t.amt; }),
    orphanedReceipts: orphaned,
    problems: problems
  };
}
function healthCheckUi() {
  var h = healthCheck();
  var msg = h.problems.length ? h.problems.join('\n') : 'All good. ' + h.entries + ' entries, no issues found.';
  SpreadsheetApp.getUi().alert('AVA Cashbook — Health check', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
function openReceiptsFolder() {
  var url = receiptsRoot().getUrl();
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:16px">Receipts folder:<br><br>' +
    '<a href="' + url + '" target="_blank">' + url + '</a></div>')
      .setWidth(420).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, 'AVA Cashbook');
}

/* ============================== BACKUP ============================== */
function backupNow() {
  ensureSetup();
  var root    = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  var backups = getOrCreateFolder(root, 'Backups');
  var name    = 'AVA Cashbook Backup ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  var copy    = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).makeCopy(name, backups);
  // keep only the newest 12 monthly backups
  var files = backups.getFiles(); var all = [];
  while (files.hasNext()) all.push(files.next());
  all.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = 12; i < all.length; i++) {
    try { all[i].setTrashed(true); } catch (e) {}
  }
  logAudit('backup', '', {name: name});
  return copy.getUrl();
}
function backupTriggerInstalled() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'backupNow';
  });
}
function installBackupTrigger() {
  removeBackupTrigger();
  ScriptApp.newTrigger('backupNow').timeBased().onMonthDay(1).atHour(3).create();
  logAudit('backup.trigger.on', '', {});
  return true;
}
function removeBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t);
  });
  logAudit('backup.trigger.off', '', {});
  return true;
}
