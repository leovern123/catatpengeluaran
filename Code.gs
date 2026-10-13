/**
 * KeuanganKu — Google Apps Script backend
 * Spreadsheet-backed personal finance tracker.
 *
 * First run: run setupKeuanganKu() from the Apps Script editor.
 * Deploy as Web app: Execute as Me; Who has access: Only myself.
 */

const APP = {
  NAME: 'KeuanganKu',
  SHEETS: {
    TRANSACTIONS: 'Transaksi',
    BUDGETS: 'Anggaran',
    SETTINGS: 'Pengaturan',
    CATEGORIES: 'Kategori'
  },
  HEADERS: {
    TRANSACTIONS: ['ID', 'Tanggal', 'Jenis', 'Kategori', 'Nominal', 'Metode', 'Catatan', 'Dibuat', 'Diubah'],
    BUDGETS: ['Bulan', 'Kategori', 'Batas'],
    SETTINGS: ['Kunci', 'Nilai'],
    CATEGORIES: ['Jenis', 'Nama', 'Aktif']
  },
  DEFAULT_CATEGORIES: [
    ['Pengeluaran', 'Makan & Minum', true],
    ['Pengeluaran', 'Bensin & Transportasi', true],
    ['Pengeluaran', 'Pulsa & Internet', true],
    ['Pengeluaran', 'Cicilan', true],
    ['Pengeluaran', 'Kuliah', true],
    ['Pengeluaran', 'Keluarga / Orang Tua', true],
    ['Pengeluaran', 'Kebutuhan Rumah', true],
    ['Pengeluaran', 'Kesehatan', true],
    ['Pengeluaran', 'Hiburan', true],
    ['Pengeluaran', 'KKN / Aktivitas', true],
    ['Pengeluaran', 'Belanja', true],
    ['Pengeluaran', 'Tabungan / Dana Darurat', true],
    ['Pengeluaran', 'Lainnya', true],
    ['Pemasukan', 'Gaji', true],
    ['Pemasukan', 'Bonus / Lembur', true],
    ['Pemasukan', 'Usaha / Freelance', true],
    ['Pemasukan', 'Pengembalian Dana', true],
    ['Pemasukan', 'Lainnya', true]
  ]
};

function doGet() {
  assertAuthorized_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Run once manually from the Apps Script editor.
 * Creates required sheets and basic settings.
 */
function setupKeuanganKu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari spreadsheet yang akan digunakan (Extensions > Apps Script).');

  ensureSheet_(ss, APP.SHEETS.TRANSACTIONS, APP.HEADERS.TRANSACTIONS);
  ensureSheet_(ss, APP.SHEETS.BUDGETS, APP.HEADERS.BUDGETS);
  const settings = ensureSheet_(ss, APP.SHEETS.SETTINGS, APP.HEADERS.SETTINGS);
  const categories = ensureSheet_(ss, APP.SHEETS.CATEGORIES, APP.HEADERS.CATEGORIES);

  if (categories.getLastRow() < 2) {
    categories.getRange(2, 1, APP.DEFAULT_CATEGORIES.length, 3).setValues(APP.DEFAULT_CATEGORIES);
  }

  const current = readSettings_();
  if (!Object.prototype.hasOwnProperty.call(current, 'saldoAwal')) {
    settings.appendRow(['saldoAwal', 0]);
  }
  if (!Object.prototype.hasOwnProperty.call(current, 'namaAplikasi')) {
    settings.appendRow(['namaAplikasi', APP.NAME]);
  }

  formatSheets_();
  SpreadsheetApp.flush();
  return 'Setup selesai. Sheet Transaksi, Anggaran, Pengaturan, dan Kategori siap digunakan.';
}

function getInitialData() {
  assertAuthorized_();
  const now = new Date();
  const month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  return {
    appName: APP.NAME,
    today: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    currentMonth: month,
    settings: readSettings_(),
    categories: getCategories_(),
    budgets: getBudgets_(month),
    summary: getSummary_(month),
    transactions: listTransactions_({ month: month, limit: 200 }),
    monthlyReport: getMonthlyReport_(6),
    categoryReport: getCategoryReport_(month)
  };
}

function getDashboardData(month) {
  assertAuthorized_();
  month = normalizeMonth_(month);
  return {
    month: month,
    summary: getSummary_(month),
    budgets: getBudgets_(month),
    transactions: listTransactions_({ month: month, limit: 200 }),
    monthlyReport: getMonthlyReport_(6),
    categoryReport: getCategoryReport_(month)
  };
}

function saveTransaction(input) {
  assertAuthorized_();
  validateTransaction_(input);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(APP.SHEETS.TRANSACTIONS);
    const now = new Date();
    const id = input.id ? String(input.id) : Utilities.getUuid();
    const existingRow = input.id ? findRowById_(sheet, id) : -1;
    if (input.id && existingRow < 2) throw new Error('Transaksi tidak ditemukan. Muat ulang halaman lalu coba lagi.');

    const date = parseDate_(input.date);
    const row = [
      id,
      date,
      input.type,
      String(input.category).trim(),
      Math.round(Number(input.amount)),
      String(input.method || 'Tunai').trim(),
      String(input.note || '').trim().slice(0, 500),
      existingRow > 1 ? sheet.getRange(existingRow, 8).getValue() : now,
      now
    ];
    if (existingRow > 1) sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);

    return { ok: true, message: input.id ? 'Transaksi diperbarui.' : 'Transaksi disimpan.', id: id };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(id) {
  assertAuthorized_();
  if (!id) throw new Error('ID transaksi tidak valid.');
  const sheet = getSheet_(APP.SHEETS.TRANSACTIONS);
  const row = findRowById_(sheet, String(id));
  if (row < 2) throw new Error('Transaksi tidak ditemukan.');
  sheet.deleteRow(row);
  return { ok: true, message: 'Transaksi dihapus.' };
}

function saveBudget(input) {
  assertAuthorized_();
  const month = normalizeMonth_(input.month);
  const category = String(input.category || '').trim();
  const limit = Number(input.limit);
  if (!category) throw new Error('Pilih kategori.');
  if (!Number.isFinite(limit) || limit < 0 || limit > 1000000000000) throw new Error('Batas anggaran tidak valid.');

  const sheet = getSheet_(APP.SHEETS.BUDGETS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === month && String(rows[i][1]) === category) {
      if (limit === 0) sheet.deleteRow(i + 1);
      else sheet.getRange(i + 1, 3).setValue(Math.round(limit));
      return { ok: true, message: limit === 0 ? 'Anggaran dihapus.' : 'Anggaran disimpan.' };
    }
  }
  if (limit > 0) sheet.appendRow([month, category, Math.round(limit)]);
  return { ok: true, message: limit === 0 ? 'Tidak ada anggaran yang perlu disimpan.' : 'Anggaran disimpan.' };
}

function saveOpeningBalance(value) {
  assertAuthorized_();
  const amount = Number(value);
  if (!Number.isFinite(amount) || Math.abs(amount) > 1000000000000) throw new Error('Saldo awal tidak valid.');
  const sheet = getSheet_(APP.SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === 'saldoAwal') {
      sheet.getRange(i + 1, 2).setValue(Math.round(amount));
      return { ok: true, message: 'Saldo awal disimpan.' };
    }
  }
  sheet.appendRow(['saldoAwal', Math.round(amount)]);
  return { ok: true, message: 'Saldo awal disimpan.' };
}

function listTransactions(filters) {
  assertAuthorized_();
  return listTransactions_(filters || {});
}

function exportTransactionsCsv(month) {
  assertAuthorized_();
  month = month ? normalizeMonth_(month) : '';
  const rows = listTransactions_({ month: month, limit: 10000 });
  const header = ['Tanggal', 'Jenis', 'Kategori', 'Nominal_Rp', 'Metode', 'Catatan'];
  const lines = [header].concat(rows.map(t => [t.date, t.type, t.category, t.amount, t.method, t.note]))
    .map(row => row.map(csvEscape_).join(','));
  return '\uFEFF' + lines.join('\r\n');
}

// -------------------- Internal helpers --------------------

function assertAuthorized_() {
  // Google deployment access is the primary gate ("Only myself").
  // Optional extra allowlist can be set in Project Settings > Script properties:
  // ALLOWED_EMAIL = your-google-account@gmail.com
  const props = PropertiesService.getScriptProperties();
  const allowed = String(props.getProperty('ALLOWED_EMAIL') || '').trim().toLowerCase();
  const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (allowed && active && active !== allowed) throw new Error('Akses ditolak: akun Google tidak diizinkan.');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  return sheet;
}

function formatSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tx = ss.getSheetByName(APP.SHEETS.TRANSACTIONS);
  if (tx && tx.getMaxRows() > 1) tx.getRange(2, 5, tx.getMaxRows() - 1, 1).setNumberFormat('"Rp" #,##0;[Red]-"Rp" #,##0');
  const budgets = ss.getSheetByName(APP.SHEETS.BUDGETS);
  if (budgets && budgets.getMaxRows() > 1) budgets.getRange(2, 3, budgets.getMaxRows() - 1, 1).setNumberFormat('"Rp" #,##0');
  const settings = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (settings && settings.getMaxRows() > 1) settings.getRange(2, 2, settings.getMaxRows() - 1, 1).setNumberFormat('"Rp" #,##0');
  [tx, budgets, settings, ss.getSheetByName(APP.SHEETS.CATEGORIES)].filter(Boolean).forEach(s => s.autoResizeColumns(1, s.getLastColumn()));
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss && ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupKeuanganKu() terlebih dahulu.');
  return sheet;
}

function readSettings_() {
  const sheet = getSheet_(APP.SHEETS.SETTINGS);
  if (sheet.getLastRow() < 2) return {};
  const result = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => {
    if (r[0] !== '') result[String(r[0])] = r[1];
  });
  return result;
}

function getCategories_() {
  const sheet = getSheet_(APP.SHEETS.CATEGORIES);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] && r[1] && (r[2] === true || String(r[2]).toLowerCase() === 'true' || r[2] === 1))
    .map(r => ({ type: String(r[0]), name: String(r[1]) }));
}

function getBudgets_(month) {
  const sheet = getSheet_(APP.SHEETS.BUDGETS);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(r => String(r[0]) === month && r[1])
    .map(r => ({ month: String(r[0]), category: String(r[1]), limit: Number(r[2]) || 0 }));
}

function getSummary_(month) {
  const sheet = getSheet_(APP.SHEETS.TRANSACTIONS);
  const settings = readSettings_();
  const opening = Number(settings.saldoAwal) || 0;
  let income = 0, expense = 0, monthIncome = 0, monthExpense = 0;
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    rows.forEach(r => {
      const type = String(r[2]);
      const amount = Number(r[4]) || 0;
      const date = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (type === 'Pemasukan') income += amount;
      if (type === 'Pengeluaran') expense += amount;
      if (!isNaN(date.getTime()) && Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM') === month) {
        if (type === 'Pemasukan') monthIncome += amount;
        if (type === 'Pengeluaran') monthExpense += amount;
      }
    });
  }
  return {
    openingBalance: opening,
    totalIncome: income,
    totalExpense: expense,
    balance: opening + income - expense,
    monthIncome: monthIncome,
    monthExpense: monthExpense,
    monthNet: monthIncome - monthExpense,
    month: month
  };
}

function listTransactions_(filters) {
  filters = filters || {};
  const month = filters.month ? normalizeMonth_(filters.month) : '';
  const query = String(filters.query || '').trim().toLowerCase();
  const type = String(filters.type || '');
  const limit = Math.min(Math.max(Number(filters.limit) || 200, 1), 10000);
  const sheet = getSheet_(APP.SHEETS.TRANSACTIONS);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.map(r => {
    const date = r[1] instanceof Date ? r[1] : new Date(r[1]);
    return {
      id: String(r[0]),
      date: isNaN(date.getTime()) ? '' : Utilities.formatDate(date, tz, 'yyyy-MM-dd'),
      type: String(r[2]),
      category: String(r[3]),
      amount: Number(r[4]) || 0,
      method: String(r[5] || ''),
      note: String(r[6] || ''),
      created: r[7] instanceof Date ? r[7].toISOString() : '',
      updated: r[8] instanceof Date ? r[8].toISOString() : ''
    };
  }).filter(t => {
    if (month && !t.date.startsWith(month)) return false;
    if (type && t.type !== type) return false;
    if (query && ![t.category, t.method, t.note, t.type].join(' ').toLowerCase().includes(query)) return false;
    return true;
  }).sort((a, b) => b.date.localeCompare(a.date) || b.updated.localeCompare(a.updated)).slice(0, limit);
}

function getMonthlyReport_(monthsBack) {
  const count = Math.min(Math.max(Number(monthsBack) || 6, 1), 24);
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const result = [];
  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const month = Utilities.formatDate(d, tz, 'yyyy-MM');
    const summary = getSummary_(month);
    result.push({ month: month, income: summary.monthIncome, expense: summary.monthExpense, net: summary.monthNet });
  }
  return result;
}

function getCategoryReport_(month) {
  const tx = listTransactions_({ month: month, limit: 10000 });
  const totals = {};
  tx.filter(t => t.type === 'Pengeluaran').forEach(t => {
    totals[t.category] = (totals[t.category] || 0) + t.amount;
  });
  return Object.keys(totals).map(category => ({ category: category, amount: totals[category] }))
    .sort((a, b) => b.amount - a.amount);
}

function validateTransaction_(input) {
  if (!input || typeof input !== 'object') throw new Error('Data transaksi tidak valid.');
  if (!['Pemasukan', 'Pengeluaran'].includes(String(input.type))) throw new Error('Jenis transaksi tidak valid.');
  const amount = Number(input.amount);
  if (!Number.isFinite(amount) || amount <= 0 || amount > 1000000000000) throw new Error('Nominal harus lebih dari Rp0.');
  if (!String(input.date || '').match(/^\d{4}-\d{2}-\d{2}$/)) throw new Error('Tanggal tidak valid.');
  if (!String(input.category || '').trim()) throw new Error('Pilih kategori.');
  if (String(input.category).length > 100) throw new Error('Nama kategori terlalu panjang.');
  if (String(input.method || '').length > 50) throw new Error('Metode pembayaran terlalu panjang.');
}

function parseDate_(dateString) {
  const parts = String(dateString).split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  if (d.getFullYear() !== parts[0] || d.getMonth() !== parts[1] - 1 || d.getDate() !== parts[2]) {
    throw new Error('Tanggal tidak valid.');
  }
  return d;
}

function normalizeMonth_(month) {
  const value = String(month || '');
  if (!/^\d{4}-\d{2}$/.test(value)) throw new Error('Format bulan harus YYYY-MM.');
  const m = Number(value.slice(5, 7));
  if (m < 1 || m > 12) throw new Error('Bulan tidak valid.');
  return value;
}

function findRowById_(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  const found = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}

function csvEscape_(value) {
  const text = String(value == null ? '' : value);
  return '"' + text.replace(/"/g, '""') + '"';
}
