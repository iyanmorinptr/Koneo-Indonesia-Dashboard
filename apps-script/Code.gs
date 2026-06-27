// ── KONEO MEMBER SYSTEM — Google Apps Script ──────────────────────────────────
// Deploy sebagai Web App: Execute as "Me", Access "Anyone"
// Sheet name: "Members" dengan kolom: Timestamp | Nama | No_HP

var SHEET_NAME = 'Members';
var REWARD_THRESHOLD = 5; // jumlah kunjungan untuk 1 reward

// ── POST: Catat kunjungan baru ─────────────────────────────────────────────────
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var nama  = (body.nama  || '').toString().trim();
    var no_hp = (body.no_hp || '').toString().trim();

    if (!nama || !no_hp) {
      return jsonResponse({ status: 'error', message: 'Nama dan No HP tidak boleh kosong.' });
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Pastikan header ada
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Nama', 'No_HP']);
    }

    // Catat kunjungan (1 baris = 1 kunjungan)
    var ts = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, nama, no_hp]);

    // Hitung total kunjungan untuk nomor HP ini
    var data  = sheet.getDataRange().getValues();
    var total = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === no_hp) total++;
    }

    // Modulo reward — sisa poin setelah klaim
    var poin_aktif = total % REWARD_THRESHOLD;
    if (poin_aktif === 0) poin_aktif = REWARD_THRESHOLD; // tepat habis = bisa klaim

    return jsonResponse({
      status: 'ok',
      total_kunjungan: total,
      poin_aktif: poin_aktif,
      bisa_klaim: (total % REWARD_THRESHOLD === 0)
    });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── GET: Daftar member teragregasi ─────────────────────────────────────────────
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'list';

    if (action === 'list') {
      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() <= 1) {
        return jsonResponse({ status: 'ok', members: [] });
      }

      var data = sheet.getDataRange().getValues();
      // Agregasi per No_HP
      var map = {}; // { no_hp: { nama, no_hp, kunjungan, terakhir } }
      for (var i = 1; i < data.length; i++) {
        var ts   = data[i][0];
        var nama = data[i][1];
        var hp   = data[i][2];
        if (!hp) continue;
        if (!map[hp]) {
          map[hp] = { nama: nama, no_hp: hp, kunjungan: 0, terakhir: '' };
        }
        map[hp].kunjungan++;
        // Simpan timestamp terbaru sebagai string
        var tsStr = ts ? ts.toString().substring(0, 10) : '';
        if (!map[hp].terakhir || tsStr > map[hp].terakhir) {
          map[hp].terakhir = tsStr;
          map[hp].nama = nama; // update nama ke yang terbaru
        }
      }

      var members = Object.values(map);
      // Urutkan: bisa klaim dulu, lalu kunjungan terbanyak
      members.sort(function(a, b) {
        var aKlaim = a.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        var bKlaim = b.kunjungan >= REWARD_THRESHOLD ? 1 : 0;
        if (bKlaim !== aKlaim) return bKlaim - aKlaim;
        return b.kunjungan - a.kunjungan;
      });

      return jsonResponse({ status: 'ok', members: members });
    }

    return jsonResponse({ status: 'error', message: 'Action tidak dikenal.' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── Helper ─────────────────────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
