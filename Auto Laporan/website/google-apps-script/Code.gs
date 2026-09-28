/**
 * Rekap Faktur Penjual -> Google Sheets
 *
 * Tempel file ini di Extensions > Apps Script pada Google Sheet tujuan,
 * lalu Deploy sebagai Web app (lihat README.md di folder website).
 *
 * Setiap laporan yang dikirim dari website menjadi 1 tab baru bernama
 * tanggal laporannya (misal "25 September 2026"). Mengirim ulang laporan
 * yang sama menimpa tab dengan nama itu.
 */

const TOKEN_PROPERTY = 'REKAP_TOKEN';

const COLOR_HEADER = '#1f4e78';
const COLOR_TOTAL = '#ddebf7';
const FIRST_ROW = 5; // row 4 = header, data starts at row 5

function doGet() {
  return json_({ ok: true, app: 'Rekap Faktur Penjual' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Data tidak valid.' });
  }

  const expected = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY);
  if (!expected) return json_({ ok: false, error: 'Token belum diatur di Apps Script (Project Settings > Script properties > REKAP_TOKEN).' });
  if (body.token !== expected) return json_({ ok: false, error: 'Token salah. Samakan token di config.js dengan REKAP_TOKEN di Apps Script.' });

  if (!Array.isArray(body.invoices) || !body.invoices.length) {
    return json_({ ok: false, error: 'Tidak ada faktur untuk dikirim.' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return json_(writeReport_(body));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function writeReport_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheetName_(body.sheetName || body.fileName || 'Rekap');

  let sh = ss.getSheetByName(name);
  const replaced = !!sh;
  if (sh) {
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
    sh.clear();
    sh.setFrozenRows(0);
  } else {
    sh = ss.insertSheet(name, 0);
  }

  const invoices = body.invoices;   // [{seller, inv, date, iso}]
  const sellers = body.bySeller;    // [{seller, items: [{inv, date: "dd/mm/yyyy", iso}]}]
  const n = invoices.length;
  const lastDetail = FIRST_ROW + n - 1;
  const tz = ss.getSpreadsheetTimeZone();
  const toDate = function (r) { return r.iso ? Utilities.parseDate(r.iso, tz, 'yyyy-MM-dd') : (r.date || ''); };

  // Rekap has one row per invoice of each seller.
  const rekapRows = sellers.reduce(function (sum, s) { return sum + s.items.length; }, 0);
  const totalRow = FIRST_ROW + rekapRows;

  // Text format first so numbers are not auto-converted; date columns get a real date format.
  sh.getRange(1, 1, Math.max(totalRow, lastDetail) + 1, 10).setNumberFormat('@');
  sh.getRange(FIRST_ROW, 1, rekapRows + 1, 1).setNumberFormat('0');
  sh.getRange(FIRST_ROW, 3, rekapRows + 1, 1).setNumberFormat('0');
  sh.getRange(FIRST_ROW, 4, rekapRows, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, 7, n, 1).setNumberFormat('0');
  sh.getRange(FIRST_ROW, 10, n, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');

  // Title
  sh.getRange('A1').setValue('Rekap Jumlah Faktur per Penjual').setFontSize(13).setFontWeight('bold');
  const sub = [body.period, 'Sumber: ' + (body.fileName || '-'),
    'Dikirim: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')]
    .filter(String).join('  |  ');
  sh.getRange('A2').setValue(sub).setFontSize(9).setFontStyle('italic').setFontColor('#595959');

  // Headers
  sh.getRange(4, 1, 1, 5).setValues([['No', 'Nama Penjual', 'Jumlah Faktur', 'Tgl Faktur', 'No. Faktur']]);
  sh.getRange(4, 7, 1, 4).setValues([['No', 'Nama Penjual', 'No. Faktur', 'Tgl Faktur']]);
  sh.getRangeList(['A4:E4', 'G4:J4']).setBackground(COLOR_HEADER).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('G3').setValue('Detail Faktur').setFontWeight('bold');

  // Detail (one row per unique invoice); the date is a real date so it sorts and filters.
  sh.getRange(FIRST_ROW, 7, n, 4).setValues(invoices.map(function (r, i) {
    return [i + 1, r.seller, r.inv, toDate(r)];
  }));

  // Rekap: one row per invoice (date next to its number); No, seller and count are merged
  // over the seller's rows. The count comes from the detail block, so edits there update the totals.
  const rekapItems = [];
  let row = FIRST_ROW;
  sellers.forEach(function (s, i) {
    const first = row;
    s.items.forEach(function (it) { rekapItems.push([toDate(it), it.inv]); row++; });
    sh.getRange(first, 1, 1, 2).setValues([[i + 1, s.seller]]);
    sh.getRange(first, 3).setFormula('=COUNTIF($H$' + FIRST_ROW + ':$H$' + lastDetail + ',B' + first + ')');
    if (row - first > 1) {
      for (let c = 1; c <= 3; c++) sh.getRange(first, c, row - first, 1).merge();
    }
  });
  if (rekapRows) sh.getRange(FIRST_ROW, 4, rekapRows, 2).setValues(rekapItems);
  sh.getRange(FIRST_ROW, 1, Math.max(rekapRows, 1), 3).setVerticalAlignment('top');

  sh.getRange(totalRow, 2).setValue('TOTAL');
  sh.getRange(totalRow, 3).setFormula('=SUM(C' + FIRST_ROW + ':C' + (totalRow - 1) + ')');
  sh.getRange(totalRow, 1, 1, 5).setBackground(COLOR_TOTAL).setFontWeight('bold');

  // Borders, font, widths
  sh.getRange(4, 1, rekapRows + 2, 5).setBorder(true, true, true, true, true, true, '#bfbfbf', null);
  sh.getRange(4, 7, n + 1, 4).setBorder(true, true, true, true, true, true, '#bfbfbf', null);
  sh.getRange(1, 1, Math.max(totalRow, lastDetail), 10).setFontFamily('Arial');
  sh.getRange(3, 1, Math.max(totalRow, lastDetail) - 2, 10).setFontSize(10);
  [6, 28, 14, 14, 22, 3, 6, 28, 20, 14].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 7); });
  sh.setFrozenRows(4);

  return {
    ok: true,
    sheetName: name,
    replaced: replaced,
    invoices: n,
    sellers: sellers.length,
    url: ss.getUrl() + '#gid=' + sh.getSheetId(),
  };
}

function sheetName_(raw) {
  const clean = String(raw).replace(/[\[\]\*\?\/\\:]/g, ' ').replace(/\s+/g, ' ').trim();
  return (clean || 'Rekap').slice(0, 100);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
