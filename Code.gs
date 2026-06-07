// ============================================================
// EXPENSE TRACKER — Google Apps Script Backend v3
// เพิ่มใหม่:
//   - addDebtEntry  : บันทึกการยืมเงิน / ให้ยืมเงิน
//   - addPayment    : ชำระหนี้ / รับเงินคืน (อัปเดต remaining)
//   - getAllDebtEntries : ดึงรายการหนี้ทั้งหมดข้ามเดือน
//   - getMonthSummary รองรับ type ใหม่ทั้งหมด
// ============================================================

const SHEET_ID       = 'YOUR_SPREADSHEET_ID_HERE';
const SETTINGS_SHEET = 'ตั้งค่า';
const DEBT_SHEET     = 'หนี้สิน';   // Sheet เก็บรายการหนี้ทั้งหมด

const MONTH_TH = [
  'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน',
  'พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม',
  'กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
];

// คอลัมน์ Sheet รายการปกติ
const COL = {
  DATE:1, NAME:2, TYPE:3, CATEGORY:4,
  WALLET:5, WALLET_FROM:6, WALLET_TO:7,
  AMOUNT:8, NOTE:9, ID:10
};

// คอลัมน์ Sheet หนี้สิน
const DCOL = {
  DATE:1, NAME:2, TYPE:3, CONTACT:4,
  WALLET:5, AMOUNT:6, REMAINING:7,
  NOTE:8, YEAR:9, ID:10
};

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const callback = e.parameter.callback;
  let result;
  try {
    const action = e.parameter.action;
    const data   = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    switch (action) {
      case 'addEntry':          result = addEntry(data);          break;
      case 'getEntries':        result = getEntries(data);        break;
      case 'updateEntry':       result = updateEntry(data);       break;
      case 'deleteEntry':       result = deleteEntry(data);       break;
      case 'getSettings':       result = getSettings();           break;
      case 'saveSettings':      result = saveSettings(data);      break;
      case 'getMonthSummary':   result = getMonthSummary(data);   break;
      case 'closeMonth':        result = closeMonth(data);        break;
      case 'addDebtEntry':      result = addDebtEntry(data);      break;
      case 'addPayment':        result = addPayment(data);        break;
      case 'getAllDebtEntries':  result = getAllDebtEntries();     break;
      case 'deleteDebtEntry':   result = deleteDebtEntry(data);   break;
      default: result = { success:false, error:'Unknown action: '+action };
    }
  } catch(err) {
    result = { success:false, error:err.toString() };
  }
  const json = JSON.stringify(result);
  const out  = callback
    ? ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  return out;
}

// ============================================================
// SHEET HELPERS
// ============================================================
function getSpreadsheet() { return SpreadsheetApp.openById(SHEET_ID); }
function sheetName(month, year) { return `${month} ${year}`; }

function getOrCreateSheet(month, year) {
  const ss = getSpreadsheet();
  const name = sheetName(month, year);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['วันที่','ชื่อรายการ','ประเภท','หมวดหมู่','ช่องเงิน','จากช่อง','ไปช่อง','จำนวนเงิน','หมายเหตุ','ID']);
    sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#e8f5e9');
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

// Sheet หนี้สิน — เก็บรายการยืม/ให้ยืม พร้อม remaining
function getDebtSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(DEBT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DEBT_SHEET);
    sheet.appendRow(['วันที่','ชื่อรายการ','ประเภท','ชื่อคน','ช่องเงิน','จำนวน','คงเหลือ','หมายเหตุ','ปี','ID']);
    sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#fff3e0');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateId() {
  return 'id_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2,6);
}

// ============================================================
// CRUD — รายการปกติ
// ============================================================
function addEntry(data) {
  const month = data.month || MONTH_TH[new Date().getMonth()];
  const year  = parseInt(data.year) || new Date().getFullYear();
  const sheet = getOrCreateSheet(month, year);
  const id    = generateId();
  sheet.appendRow([
    data.date, data.name, data.type, data.category||'',
    data.wallet||'', data.walletFrom||'', data.walletTo||'',
    parseFloat(data.amount)||0, data.note||'', id
  ]);
  return { success:true, id, month, year };
}

function getEntries(data) {
  const month = data.month || MONTH_TH[new Date().getMonth()];
  const year  = parseInt(data.year) || new Date().getFullYear();
  const sheet = getSpreadsheet().getSheetByName(sheetName(month, year));
  if (!sheet) return { success:true, entries:[] };
  const rows = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[9]) continue;
    entries.push({
      date:r[0], name:r[1], type:r[2], category:r[3],
      wallet:r[4], walletFrom:r[5], walletTo:r[6],
      amount:r[7], note:r[8], id:r[9]
    });
  }
  return { success:true, entries };
}

function updateEntry(data) {
  const month = data.month || MONTH_TH[new Date().getMonth()];
  const year  = parseInt(data.year) || new Date().getFullYear();
  const sheet = getSpreadsheet().getSheetByName(sheetName(month, year));
  if (!sheet) return { success:false, error:'Sheet not found' };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][9]) === String(data.id)) {
      sheet.getRange(i+1,1,1,10).setValues([[
        data.date, data.name, data.type, data.category||'',
        data.wallet||'', data.walletFrom||'', data.walletTo||'',
        parseFloat(data.amount)||0, data.note||'', data.id
      ]]);
      return { success:true };
    }
  }
  return { success:false, error:'Entry not found' };
}

function deleteEntry(data) {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    if ([SETTINGS_SHEET, DEBT_SHEET].includes(sheet.getName())) continue;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][9]) === String(data.id)) {
        sheet.deleteRow(i+1);
        return { success:true };
      }
    }
  }
  return { success:false, error:'Entry not found' };
}

// ============================================================
// SETTINGS
// ============================================================
function getSettings() {
  const sheet = getSettingsSheet();
  const rows  = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) settings[String(rows[i][0])] = rows[i][1];
  }
  return { success:true, settings };
}

function saveSettings(data) {
  const sheet = getSettingsSheet();
  const rows  = sheet.getDataRange().getValues();
  for (const key of Object.keys(data)) {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(key)) {
        sheet.getRange(i+1,2).setValue(data[key]);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([key, data[key]]);
  }
  return { success:true };
}

// ============================================================
// DEBT — บันทึกการยืมเงิน / ให้ยืมเงิน
// ============================================================
function addDebtEntry(data) {
  const sheet = getDebtSheet();
  const id    = generateId();
  const year  = parseInt(data.year) || new Date().getFullYear();
  const amount = parseFloat(data.amount) || 0;

  sheet.appendRow([
    data.date, data.name, data.type, data.contact||'',
    data.wallet||'', amount, amount,   // remaining = amount เต็มตอนแรก
    data.note||'', year, id
  ]);

  // ถ้า "ยืมมาเก็บ" → บวกเงินเข้าช่องเงินใน Sheet รายการปกติด้วย
  if (data.type === 'ยืมเงินมา(เก็บไว้)' && data.wallet) {
    const month = data.month || MONTH_TH[new Date().getMonth()];
    addEntry({
      date: data.date, name: 'ยืมเงิน: '+data.name,
      type: 'รายรับ', wallet: data.wallet,
      amount: amount, note: 'ยืมจาก '+data.contact, year
    });
  }

  // ถ้า "ให้ยืมเงิน" → หักเงินออกจากช่องเงินใน Sheet รายการปกติ
  if (data.type === 'ให้ยืมเงิน' && data.wallet) {
    const month = data.month || MONTH_TH[new Date().getMonth()];
    addEntry({
      date: data.date, name: 'ให้ยืม: '+data.name,
      type: 'รายจ่าย', wallet: data.wallet,
      amount: amount, note: 'ให้ '+data.contact+' ยืม', year
    });
  }

  return { success:true, id };
}

// ============================================================
// PAYMENT — ชำระหนี้ / รับเงินคืน
// ============================================================
function addPayment(data) {
  const debtSheet = getDebtSheet();
  const rows = debtSheet.getDataRange().getValues();
  const amount = parseFloat(data.amount) || 0;
  const year = parseInt(data.year) || new Date().getFullYear();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][DCOL.ID-1]) !== String(data.debtId)) continue;

    const currentRemaining = parseFloat(rows[i][DCOL.REMAINING-1]) || 0;

    // กันชำระเกินยอด
    const payAmount = Math.min(amount, currentRemaining);
    const newRemaining = Math.max(0, currentRemaining - payAmount);

    // อัปเดต remaining ใน debt sheet
    debtSheet.getRange(i+1, DCOL.REMAINING).setValue(newRemaining);

    const contactName = rows[i][DCOL.CONTACT-1] || '';
    const debtName    = rows[i][DCOL.NAME-1] || '';
    const debtType    = rows[i][DCOL.TYPE-1] || '';
    const month = data.month || MONTH_TH[new Date().getMonth()];

    if (data.paymentType === 'paydebt') {
      // ชำระหนี้ → หักเงินออกจากช่องเงิน
      addEntry({
        date: data.date, name: 'ชำระหนี้: '+debtName,
        type: 'รายจ่าย', wallet: data.wallet,
        amount: payAmount, note: 'คืนเงิน '+contactName, year
      });
    } else if (data.paymentType === 'repay') {
      // รับเงินคืน → บวกเงินเข้าช่องเงิน
      addEntry({
        date: data.date, name: 'รับเงินคืน: '+debtName,
        type: 'รายรับ', wallet: data.wallet,
        amount: payAmount, note: contactName+' คืนเงิน', year
      });
    }

    return { success:true, paid:payAmount, remaining:newRemaining };
  }
  return { success:false, error:'Debt entry not found' };
}

// ============================================================
// DELETE DEBT ENTRY
// ============================================================
function deleteDebtEntry(data) {
  const sheet = getDebtSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][DCOL.ID-1]) === String(data.id)) {
      sheet.deleteRow(i+1);
      return { success:true };
    }
  }
  return { success:false, error:'Debt entry not found' };
}

// ============================================================
// GET ALL DEBT ENTRIES — ดึงทุกรายการจาก Sheet หนี้สิน
// ============================================================
function getAllDebtEntries() {
  const sheet = getDebtSheet();
  const rows  = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[DCOL.ID-1]) continue;
    entries.push({
      date:      r[0],
      name:      r[1],
      type:      r[2],
      contact:   r[3],
      wallet:    r[4],
      amount:    r[5],
      remaining: r[6],
      note:      r[7],
      year:      r[8],
      id:        r[9]
    });
  }
  return { success:true, entries };
}

// ============================================================
// SUMMARY
// ============================================================
function getMonthSummary(data) {
  const month   = data.month || MONTH_TH[new Date().getMonth()];
  const year    = parseInt(data.year) || new Date().getFullYear();
  const wallets = Array.isArray(data.wallets) && data.wallets.length > 0
    ? data.wallets
    : ['เงินสด (ใช้จ่าย)','เงินสด (เงินเก็บ)','กรุงไทย / KTC','金融卡 (郵政)','SCB / Planet','EasyCard (บัตรนกสีเขียว)','EasyCard (บัตรนักเรียน)'];

  const settings = getSettings().settings;
  const balances = {};
  for (const w of wallets) {
    balances[w] = parseFloat(settings[`opening_${month}_${year}_${w}`] || 0);
  }

  const categoryTotals = {};
  const walletIncome   = {};
  const walletExpense  = {};
  for (const w of wallets) { walletIncome[w] = 0; walletExpense[w] = 0; }

  const sheet = getSpreadsheet().getSheetByName(sheetName(month, year));
  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[9]) continue;
      const type   = r[2], wallet = r[4], walletFrom = r[5], walletTo = r[6];
      const amount = parseFloat(r[7]) || 0;
      const cat    = r[3];

      if (type === 'รายรับ') {
        if (wallets.includes(wallet)) walletIncome[wallet] += amount;

      } else if (type === 'รายจ่าย') {
        if (wallets.includes(wallet)) walletExpense[wallet] += amount;
        if (cat) categoryTotals[cat] = (categoryTotals[cat]||0) + amount;

      } else if (type === 'ย้ายเงิน') {
        if (wallets.includes(walletFrom)) walletExpense[walletFrom] += amount;
        if (wallets.includes(walletTo))   walletIncome[walletTo]   += amount;

      } else if (type === 'จำไม่ได้') {
        if (wallets.includes(wallet)) walletExpense[wallet] += amount;
        categoryTotals['จำไม่ได้'] = (categoryTotals['จำไม่ได้']||0) + amount;
      }
      // หมายเหตุ: ยืมมาเก็บ/ให้ยืม/ชำระหนี้/รับเงินคืน
      // → ถูก addEntry เป็น รายรับ/รายจ่าย ไปแล้ว จึงถูกนับอัตโนมัติ
    }
  }

  const walletSummary = {};
  let totalBalance = 0;
  let totalExpense = 0;
  let transferTotal = 0;

  for (const w of wallets) {
    const balance = balances[w] + walletIncome[w] - walletExpense[w];
    walletSummary[w] = { opening:balances[w], income:walletIncome[w], expense:walletExpense[w], balance };
    totalBalance += balance;
    totalExpense += walletExpense[w];
  }

  // หักรายจ่าย ย้ายเงิน ออกจาก totalExpense (นับซ้ำ)
  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][2] === 'ย้ายเงิน') transferTotal += (parseFloat(rows[i][7])||0);
    }
  }
  totalExpense = Math.max(0, totalExpense - transferTotal);

  return { success:true, month, year, walletSummary, categoryTotals, totalBalance, totalExpense };
}

// ============================================================
// CLOSE MONTH
// ============================================================
function closeMonth(data) {
  const month   = data.month;
  const year    = parseInt(data.year);
  const wallets = Array.isArray(data.wallets) ? data.wallets : [];

  const summary = getMonthSummary({ month, year, wallets });
  if (!summary.success) return summary;

  const curIdx    = MONTH_TH.indexOf(month);
  const nextIdx   = (curIdx + 1) % 12;
  const nextMonth = MONTH_TH[nextIdx];
  const nextYear  = nextIdx === 0 ? year + 1 : year;

  const settingsData = {};
  for (const [w, info] of Object.entries(summary.walletSummary)) {
    settingsData[`opening_${nextMonth}_${nextYear}_${w}`] = info.balance;
  }
  saveSettings(settingsData);
  getOrCreateSheet(nextMonth, nextYear);

  return { success:true, nextMonth, nextYear, balances:summary.walletSummary };
}
