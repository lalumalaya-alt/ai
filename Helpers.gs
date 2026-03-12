/***************************************************
 * HELPER FUNCTIONS
 ***************************************************/
function generateTicketId() {
  const prefix = "TKT";
  const timestamp = new Date().getTime().toString().slice(-6);
  const random = Math.floor(Math.random() * 900 + 100);
  return `${prefix}${timestamp}${random}`;
}

function generateFinanceId() {
  const prefix = "FIN";
  const timestamp = new Date().getTime().toString().slice(-6);
  const random = Math.floor(Math.random() * 900 + 100);
  return `${prefix}${timestamp}${random}`;
}

function generateCheckInId() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const setSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  let nextNum = 1;
  if (setSheet && setSheet.getLastRow() > 1) {
    nextNum = parseInt(setSheet.getRange(2, SET_NEXT_CHECKIN_COL + 1).getValue()) || 1;
    setSheet.getRange(2, SET_NEXT_CHECKIN_COL + 1).setValue(nextNum + 1);
  }
  return "CHK-" + String(nextNum).padStart(4, '0');
}

function generateBillNumber() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const setSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  let nextNum = 1;
  if (setSheet && setSheet.getLastRow() > 1) {
    nextNum = parseInt(setSheet.getRange(2, SET_NEXT_BILL_COL + 1).getValue()) || 1;
    setSheet.getRange(2, SET_NEXT_BILL_COL + 1).setValue(nextNum + 1);
  }
  return "BILL-" + String(nextNum).padStart(6, '0');
}

function generateOrderId() {
  return "ORD-" + new Date().getTime().toString().slice(-6) + Math.floor(Math.random() * 900 + 100);
}

function daysBetween(d1, d2) {
  let diff = d2.getTime() - d1.getTime();
  let days = Math.ceil(diff / (1000 * 3600 * 24));
  return days;
}

/**
 * Sequential ID generator using SETTINGS sheet as counter store.
 * type: 'invoice' → INV-0001, 'quote' → QTN-0001
 */
function getNextSequentialId(type) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    throw new Error("Settings sheet not found. Please run Setup Demo Data first.");
  }

  const prefixMap = { invoice: 'INV', quote: 'QTN' };
  const colMap = { invoice: SET_NEXT_INVOICE_COL, quote: SET_NEXT_QUOTE_COL };

  const prefix = prefixMap[type];
  const col = colMap[type];
  if (!prefix || col === undefined) throw new Error("Invalid sequential ID type: " + type);

  const cell = settingsSheet.getRange(2, col + 1);
  let currentNum = parseInt(cell.getValue()) || 1;
  const id = prefix + '-' + String(currentNum).padStart(4, '0');

  cell.setValue(currentNum + 1);
  SpreadsheetApp.flush();

  return id;
}

/**
 * Finds or creates a Drive folder by name (in root).
 */
function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function numberToWords(num) {
  if (num === 0) return 'Zero';
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  num = Math.round(Math.abs(num));
  if (num < 20) return ones[num];
  if (num < 100) return tens[Math.floor(num / 10)] + (num % 10 ? ' ' + ones[num % 10] : '');
  if (num < 1000) return ones[Math.floor(num / 100)] + ' Hundred' + (num % 100 ? ' ' + numberToWords(num % 100) : '');
  if (num < 100000) return numberToWords(Math.floor(num / 1000)) + ' Thousand' + (num % 1000 ? ' ' + numberToWords(num % 1000) : '');
  if (num < 10000000) return numberToWords(Math.floor(num / 100000)) + ' Lakh' + (num % 100000 ? ' ' + numberToWords(num % 100000) : '');
  return numberToWords(Math.floor(num / 10000000)) + ' Crore' + (num % 10000000 ? ' ' + numberToWords(num % 10000000) : '');
}
