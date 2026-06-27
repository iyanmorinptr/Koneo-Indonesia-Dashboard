// ── KONEO MEMBER SYSTEM — Google Apps Script ──────────────────────────────────
// Deploy sebagai Web App: Execute as "Me", Access "Anyone"

var SHEET_NAME = 'Members';
var REWARD_THRESHOLD = 5;

function doGet(e) {
  var cb     = (e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  var action = (e.parameter && e.parameter.action)   ? e.parameter.action   : 'list';

  try {
    // ── Catat kunjungan ───────────────────────────────────────────────────────
    if (action === 'catat') {
      var nama  = (e.parameter.nama  || '').toString().trim();
      var no_hp = (e.parameter.no_hp || '').toString().trim();
      if (!nama || !no_hp) return respond({ status: 'error', message: 'Data tidak lengkap.' }, cb);

      var sheet = getSheet();
      var ts    = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([ts, nama, no_hp]);

      var total = countVisits(sheet, no_hp);
      return respond({ status: 'ok', total_kunjungan: total, bisa_klaim: (total % REWARD_THRESHOLD === 0) }, cb);
    }

    // ── Daftar member ─────────────────────────────────────────────────────────
    if (action === 'list') {
      var sheet = getSheet();
      if (sheet.getLastRow() <= 1) return respond({ status: 'ok', members: [] }, cb);

      var data = sheet.getDataRange().getValues();
      var map  = {};
      for (var i = 1; i < data.length; i++) {
        var ts = data[i][0]; var nama = data[i][1]; var hp = data[i][2];
        if (!hp) continue;
        if (!map[hp]) map[hp] = { nama: nama, no_hp: hp, kunjungan: 0, terakhir: '' };
        map[hp].kunjungan++;
        var tsStr = ts ? ts.toString().substring(0, 10) : '';
        if (!map[hp].terakhir || tsStr > map[hp].terakhir) { map[hp].terakhir = tsStr; map[hp].nama = nama; }
      }
      var members = Object.values(map);
      members.sort(function(a, b) {
        var aK = a.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        var bK = b.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        return bK !== aK ? bK - aK : b.kunjungan - a.kunjungan;
      });
      return respond({ status: 'ok', members: members }, cb);
    }

    // ── Simpan backup ─────────────────────────────────────────────────────────
    if (action === 'backup_save') {
      var jsonStr = (e.parameter.data || '').toString();
      if (!jsonStr) return respond({ status: 'error', message: 'No data.' }, cb);
      var ts = saveBackup(jsonStr);
      return respond({ status: 'ok', timestamp: ts }, cb);
    }

    // ── Ambil backup terbaru ──────────────────────────────────────────────────
    if (action === 'backup_get') {
      var backup = getLatestBackup();
      if (!backup) return respond({ status: 'ok', backup: null }, cb);
      return respond({ status: 'ok', backup: backup }, cb);
    }

    return respond({ status: 'error', message: 'Action tidak dikenal.' }, cb);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, cb);
  }
}

function doPost(e) {
  try {
    var body  = JSON.parse(e.postData.contents);
    var nama  = (body.nama  || '').toString().trim();
    var no_hp = (body.no_hp || '').toString().trim();
    if (!nama || !no_hp) return respond({ status: 'error', message: 'Data tidak lengkap.' }, null);
    var sheet = getSheet();
    var ts    = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, nama, no_hp]);
    var total = countVisits(sheet, no_hp);
    return respond({ status: 'ok', total_kunjungan: total, bisa_klaim: (total % REWARD_THRESHOLD === 0) }, null);
  } catch (err) {
    return respond({ status: 'error', message: err.message }, null);
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); sheet.appendRow(['Timestamp', 'Nama', 'No_HP']); }
  else if (sheet.getLastRow() === 0) { sheet.appendRow(['Timestamp', 'Nama', 'No_HP']); }
  return sheet;
}

function countVisits(sheet, no_hp) {
  var data  = sheet.getDataRange().getValues();
  var total = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][2] === no_hp) total++; }
  return total;
}

function respond(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── BACKUP ────────────────────────────────────────────────────────────────────
var BACKUP_SHEET_NAME = 'DashboardBackup';

function getBackupSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(BACKUP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BACKUP_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'BackupJSON']);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'BackupJSON']);
  }
  return sheet;
}

function saveBackup(jsonStr) {
  var sheet = getBackupSheet();
  var ts    = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
  // Keep only last 10 backups to avoid sheet growing too large
  var lastRow = sheet.getLastRow();
  if (lastRow >= 11) {
    // Remove oldest rows (row 2 is oldest data row)
    sheet.deleteRow(2);
  }
  sheet.appendRow([ts, jsonStr]);
  SpreadsheetApp.flush();
  return ts;
}

function getLatestBackup() {
  var sheet   = getBackupSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  var row = sheet.getRange(lastRow, 1, 1, 2).getValues()[0];
  return { timestamp: row[0].toString(), data: row[1].toString() };
}
