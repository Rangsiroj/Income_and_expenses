// ==========================================
// EXPENSE TRACKER - Google Apps Script
// ==========================================

const SHEET_ID = '1P5EagTgb2btjRjT86rRwszjtwCXURqA-kv6nte4T5Ho';
const SETTINGS_SHEET = 'ตั้งค่า';

const MONTH_NAMES = [
  'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน',
  'พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม',
  'กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
];

const COLUMNS = {
  DATE: 1,
  NAME: 2,
  TYPE: 3,
  CATEGORY: 4,
  WALLET: 5,
  WALLET_FROM: 6,
  WALLET_TO: 7,
  AMOUNT: 8,
  NOTE: 9,
  ID: 10
};

// ==========================================
// WEB APP ENTRY POINT
// ==========================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const callback = e.parameter.callback;
  let result;
  try {
    const action = e.parameter.action;
    const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};

    switch (action) {
      case 'addEntry':       result = addEntry(data); break;
      case 'getEntries':     result = getEntries(data); break;
      case 'updateEntry':    result = updateEntry(data); break;
      case 'deleteEntry':    result = deleteEntry(data); break;
      case 'getSettings':    result = getSettings(); break;
      case 'saveSettings':   result = saveSettings(data); break;
      case 'getMonthSummary':result = getMonthSummary(data); break;
      case 'getAllMonths':    result = getAllMonths(); break;
      case 'closeMonth':     result = closeMonth(data); break;
      default: result = { success: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  const json = JSON.stringify(result);
  const output = callback
    ? ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ==========================================
// SHEET HELPERS
// ==========================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getOrCreateSheet(monthName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(monthName);
  if (!sheet) {
    sheet = ss.insertSheet(monthName);
    // Add header row
    sheet.appendRow([
      'วันที่','ชื่อรายการ','ประเภท','หมวดหมู่',
      'ช่องเงิน','จากช่อง','ไปช่อง','จำนวนเงิน','หมายเหตุ','ID'
    ]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(['key','value']);
    sheet.getRange(1,1,1,2).setFontWeight('bold');
  }
  return sheet;
}

function generateId() {
  return 'id_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,5);
}

function getMonthNameFromDate(dateStr) {
  const d = new Date(dateStr);
  return MONTH_NAMES[d.getMonth()];
}

// ==========================================
// CRUD OPERATIONS
// ==========================================

function addEntry(data) {
  const monthName = getMonthNameFromDate(data.date);
  const sheet = getOrCreateSheet(monthName);
  const id = generateId();

  sheet.appendRow([
    data.date,
    data.name,
    data.type,
    data.category || '',
    data.wallet || '',
    data.walletFrom || '',
    data.walletTo || '',
    parseFloat(data.amount) || 0,
    data.note || '',
    id
  ]);

  return { success: true, id: id, month: monthName };
}

function getEntries(data) {
  const monthName = data.month || MONTH_NAMES[new Date().getMonth()];
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) return { success: true, entries: [] };

  const rows = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[COLUMNS.ID - 1]) continue;
    entries.push({
      date: r[0], name: r[1], type: r[2], category: r[3],
      wallet: r[4], walletFrom: r[5], walletTo: r[6],
      amount: r[7], note: r[8], id: r[9]
    });
  }
  return { success: true, entries: entries };
}

function updateEntry(data) {
  const monthName = getMonthNameFromDate(data.date);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COLUMNS.ID - 1] === data.id) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([[
        data.date, data.name, data.type, data.category || '',
        data.wallet || '', data.walletFrom || '', data.walletTo || '',
        parseFloat(data.amount) || 0, data.note || '', data.id
      ]]);
      return { success: true };
    }
  }
  return { success: false, error: 'Entry not found' };
}

function deleteEntry(data) {
  // Search all month sheets for the entry
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    if (sheet.getName() === SETTINGS_SHEET) continue;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][COLUMNS.ID - 1] === data.id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
  }
  return { success: false, error: 'Entry not found' };
}

// ==========================================
// SETTINGS
// ==========================================

function getSettings() {
  const sheet = getSettingsSheet();
  const rows = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
  }
  return { success: true, settings: settings };
}

function saveSettings(data) {
  const sheet = getSettingsSheet();
  const rows = sheet.getDataRange().getValues();

  for (const key of Object.keys(data)) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, data[key]]);
    }
  }
  return { success: true };
}

// ==========================================
// SUMMARY & ANALYTICS
// ==========================================

function getMonthSummary(data) {
  const monthName = data.month || MONTH_NAMES[new Date().getMonth()];
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(monthName);

  const settingsResult = getSettings();
  const settings = settingsResult.settings;

  const wallets = [
    'เงินสด (ใช้จ่าย)', 'เงินสด (เงินเก็บ)', 'กรุงไทย / KTC',
    '金融卡 (郵政)', 'SCB / Planet', 'EasyCard (บัตรนกสีเขียว)', 'EasyCard (บัตรนักเรียน)'
  ];

  // Get opening balances for this month
  const balances = {};
  for (const w of wallets) {
    const key = `opening_${monthName}_${w}`;
    balances[w] = parseFloat(settings[key] || 0);
  }

  const categoryTotals = { 'ของกิน': 0, 'ของใช้': 0, 'รายจ่ายจำเป็น': 0, 'การเดินทาง': 0, 'จำไม่ได้': 0 };
  const walletIncome = {};
  const walletExpense = {};
  for (const w of wallets) { walletIncome[w] = 0; walletExpense[w] = 0; }

  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[COLUMNS.ID - 1]) continue;
      const type = r[2];
      const wallet = r[4];
      const walletFrom = r[5];
      const walletTo = r[6];
      const amount = parseFloat(r[7]) || 0;
      const category = r[3];

      if (type === 'รายรับ') {
        if (wallets.includes(wallet)) walletIncome[wallet] += amount;
      } else if (type === 'รายจ่าย' || type === 'ชำระหนี้') {
        if (wallets.includes(wallet)) walletExpense[wallet] += amount;
        if (type === 'รายจ่าย' && categoryTotals.hasOwnProperty(category)) {
          categoryTotals[category] += amount;
        }
      } else if (type === 'ย้ายเงิน') {
        if (wallets.includes(walletFrom)) walletExpense[walletFrom] += amount;
        if (wallets.includes(walletTo)) walletIncome[walletTo] += amount;
      } else if (type === 'ยืมเงิน') {
        if (wallets.includes(wallet)) walletIncome[wallet] += amount;
      } else if (type === 'จำไม่ได้') {
        if (wallets.includes(wallet)) walletExpense[wallet] += amount;
        categoryTotals['จำไม่ได้'] = (categoryTotals['จำไม่ได้'] || 0) + amount;
      }
    }
  }

  const walletSummary = {};
  let totalBalance = 0;
  let totalExpense = 0;

  for (const w of wallets) {
    const balance = balances[w] + walletIncome[w] - walletExpense[w];
    walletSummary[w] = {
      opening: balances[w],
      income: walletIncome[w],
      expense: walletExpense[w],
      balance: balance
    };
    totalBalance += balance;
    totalExpense += walletExpense[w];
  }

  return {
    success: true,
    month: monthName,
    walletSummary,
    categoryTotals,
    totalBalance,
    totalExpense
  };
}

function getAllMonths() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const months = [];
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (MONTH_NAMES.includes(name)) months.push(name);
  }
  // Sort by month order
  months.sort((a, b) => MONTH_NAMES.indexOf(a) - MONTH_NAMES.indexOf(b));
  return { success: true, months: months };
}

function closeMonth(data) {
  const monthName = data.month;
  const nextMonthIndex = (MONTH_NAMES.indexOf(monthName) + 1) % 12;
  const nextMonthName = MONTH_NAMES[nextMonthIndex];

  const summary = getMonthSummary({ month: monthName });
  if (!summary.success) return summary;

  const settingsData = {};
  for (const [wallet, info] of Object.entries(summary.walletSummary)) {
    settingsData[`opening_${nextMonthName}_${wallet}`] = info.balance;
  }

  saveSettings(settingsData);
  getOrCreateSheet(nextMonthName);

  return { success: true, nextMonth: nextMonthName, balances: summary.walletSummary };
}
