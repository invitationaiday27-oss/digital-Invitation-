/**
 * ================================================================
 * GOOGLE APPS SCRIPT — Kad Jemputan Digital Aiman & Izzati
 * ================================================================
 * FUNGSI:
 *   1. Simpan RSVP tetamu          -> Sheet "RSVP"
 *   2. Simpan ucapan tetamu         -> Sheet "Ucapan"
 *   3. Papar semula ucapan tetamu   -> doGet (JSON) ke dalam kad
 *
 * CARA SETUP:
 * 1. Pergi ke sheets.google.com -> buat spreadsheet baru
 *    Namakan: "Kad Jemputan - Aiman & Izzati"
 * 2. Klik Extensions > Apps Script
 * 3. Padam semua kod lama, tampal kod ini sepenuhnya
 * 4. Klik Save (Ctrl+S)
 * 5. Klik Deploy > New Deployment
 *      - Type           : Web app
 *      - Execute as     : Me
 *      - Who has access : Anyone
 * 6. Klik Deploy > beri kebenaran (Allow)
 * 7. Salin URL "exec" -> letak dalam kad-jemputan.html pada:
 *      SHEET_URL: 'URL_ANDA_DI_SINI'
 * ================================================================
 */

// ── Cipta header sheet (auto bila kosong) ─────────────────────────
function setupSheets(ss) {
  // Sheet RSVP
  var rsvp = ss.getSheetByName('RSVP') || ss.insertSheet('RSVP');
  if (!rsvp.getRange('A1').getValue()) {
    var h1 = ['Tarikh & Masa', 'Nama', 'Kehadiran', 'Bilangan'];
    var r1 = rsvp.getRange(1, 1, 1, h1.length);
    r1.setValues([h1]);
    r1.setFontWeight('bold').setBackground('#7d2d3f').setFontColor('#ffffff').setHorizontalAlignment('center');
    rsvp.setFrozenRows(1);
    rsvp.autoResizeColumns(1, h1.length);
  }

  // Sheet Ucapan
  var ucapan = ss.getSheetByName('Ucapan') || ss.insertSheet('Ucapan');
  if (!ucapan.getRange('A1').getValue()) {
    var h2 = ['Tarikh & Masa', 'Nama', 'Ucapan', 'Papar'];
    var r2 = ucapan.getRange(1, 1, 1, h2.length);
    r2.setValues([h2]);
    r2.setFontWeight('bold').setBackground('#b8974f').setFontColor('#ffffff').setHorizontalAlignment('center');
    ucapan.setFrozenRows(1);
    ucapan.autoResizeColumns(1, h2.length);
  }

  return { rsvp: rsvp, ucapan: ucapan };
}

// ── Terima POST dari kad (RSVP atau Ucapan) ───────────────────────
function doPost(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var data   = JSON.parse(e.postData.contents);
    var sheets = setupSheets(ss);
    var now    = new Date().toLocaleString('ms-MY', { timeZone: 'Asia/Kuala_Lumpur' });

    if (data.action === 'rsvp') {
      sheets.rsvp.appendRow([
        now,
        data.nama      || '',
        data.kehadiran || '',
        data.bilangan  || ''
      ]);
      sheets.rsvp.autoResizeColumns(1, 4);

    } else if (data.action === 'ucapan') {
      // 'Papar' = YA secara lalai supaya ucapan terus muncul
      sheets.ucapan.appendRow([
        now,
        data.nama   || 'Tetamu',
        data.ucapan || '',
        'YA'
      ]);
      sheets.ucapan.autoResizeColumns(1, 4);
    }

    return jsonOut({ status: 'OK' });

  } catch (err) {
    console.error('doPost error:', err.toString());
    return jsonOut({ status: 'ERROR', message: err.toString() });
  }
}

// ── GET: pulangkan senarai ucapan (untuk dipaparkan dalam kad) ────
function doGet(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = setupSheets(ss);
    var sheet  = sheets.ucapan;
    var lastRow = sheet.getLastRow();

    var list = [];
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      for (var i = rows.length - 1; i >= 0; i--) { // terbaru di atas
        var papar = (rows[i][3] || 'YA').toString().toUpperCase();
        if (papar === 'TIDAK' || papar === 'NO') continue; // sembunyikan jika ditandakan
        list.push({
          masa:   rows[i][0],
          nama:   rows[i][1],
          ucapan: rows[i][2]
        });
      }
    }

    return jsonOut({ status: 'OK', ucapan: list });

  } catch (err) {
    return jsonOut({ status: 'ERROR', message: err.toString(), ucapan: [] });
  }
}

// ── Helper: output JSON ───────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
