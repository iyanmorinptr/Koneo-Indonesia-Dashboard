// ── KONEO MEMBER SYSTEM — Google Apps Script ──────────────────────────────────
// Deploy sebagai Web App: Execute as "Me", Access "Anyone"
// Sheet name: "Members" dengan kolom: Timestamp | Nama | No_HP

var SHEET_NAME = 'Members';
var REWARD_THRESHOLD = 5;

// ── GET: handle semua request (list & catat kunjungan via GET param) ───────────
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'list';

    // ── Catat kunjungan baru ──────────────────────────────────────────────────
    if (action === 'catat') {
      var nama  = (e.parameter.nama  || '').toString().trim();
      var no_hp = (e.parameter.no_hp || '').toString().trim();

      if (!nama || !no_hp) {
        return jsonResponse({ status: 'error', message: 'Nama dan No HP tidak boleh kosong.' });
      }

      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Timestamp', 'Nama', 'No_HP']);
      }

      var ts = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([ts, nama, no_hp]);

      var data  = sheet.getDataRange().getValues();
      var total = 0;
      for (var i = 1; i < data.length; i++) {
        if (data[i][2] === no_hp) total++;
      }

      var bisa_klaim = (total % REWARD_THRESHOLD === 0);
      return jsonResponse({ status: 'ok', total_kunjungan: total, bisa_klaim: bisa_klaim });
    }

    // ── Daftar member teragregasi ─────────────────────────────────────────────
    if (action === 'list') {
      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() <= 1) {
        return jsonResponse({ status: 'ok', members: [] });
      }

      var data = sheet.getDataRange().getValues();
      var map  = {};
      for (var i = 1; i < data.length; i++) {
        var ts   = data[i][0];
        var nama = data[i][1];
        var hp   = data[i][2];
        if (!hp) continue;
        if (!map[hp]) map[hp] = { nama: nama, no_hp: hp, kunjungan: 0, terakhir: '' };
        map[hp].kunjungan++;
        var tsStr = ts ? ts.toString().substring(0, 10) : '';
        if (!map[hp].terakhir || tsStr > map[hp].terakhir) {
          map[hp].terakhir = tsStr;
          map[hp].nama = nama;
        }
      }

      var members = Object.values(map);
      members.sort(function(a, b) {
        var aK = a.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        var bK = b.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        if (bK !== aK) return bK - aK;
        return b.kunjungan - a.kunjungan;
      });

      return jsonResponse({ status: 'ok', members: members });
    }

    return jsonResponse({ status: 'error', message: 'Action tidak dikenal.' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── POST tetap ada sebagai fallback ───────────────────────────────────────────
function doPost(e) {
  try {
    var body  = JSON.parse(e.postData.contents);
    var nama  = (body.nama  || '').toString().trim();
    var no_hp = (body.no_hp || '').toString().trim();
    if (!nama || !no_hp) return jsonResponse({ status: 'error', message: 'Data tidak lengkap.' });

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp', 'Nama', 'No_HP']);

    var ts = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, nama, no_hp]);

    var data = sheet.getDataRange().getValues();
    var total = 0;
    for (var i = 1; i < data.length; i++) { if (data[i][2] === no_hp) total++; }

    return jsonResponse({ status: 'ok', total_kunjungan: total, bisa_klaim: (total % REWARD_THRESHOLD === 0) });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── Helper ────────────────────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
