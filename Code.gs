/**
 * Serves the HTML file
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Income & Expense Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Setup function to initialize sheets and headers
 * Run this function once manually from the Apps Script editor
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup Income Sheet
  let incomeSheet = ss.getSheetByName('Income');
  if (!incomeSheet) {
    incomeSheet = ss.insertSheet('Income');
  }
  const incomeHeaders = [
    'Date', 'Sl Number', 'Entry Number', 'Room Number', 'Other',
    'Room Rent', 'Fooding', 'Total', 'Payment Status', 'Mode Of Payment',
    'Entry By', 'Payment Date'
  ];
  incomeSheet.getRange(1, 1, 1, incomeHeaders.length).setValues([incomeHeaders]);
  incomeSheet.getRange(1, 1, 1, incomeHeaders.length).setFontWeight('bold');

  // Setup Expenses Sheet
  let expenseSheet = ss.getSheetByName('Expenses');
  if (!expenseSheet) {
    expenseSheet = ss.insertSheet('Expenses');
  }
  const expenseHeaders = [
    'Date', 'Sl Number', 'Type', 'Description', 'Details', 'Amount',
    'Payment Status', 'Source Of Payment', 'Mode Of Payment',
    'Entry By', 'Payment Date'
  ];
  expenseSheet.getRange(1, 1, 1, expenseHeaders.length).setValues([expenseHeaders]);
  expenseSheet.getRange(1, 1, 1, expenseHeaders.length).setFontWeight('bold');
}

/**
 * Helper to generate an auto-incremented SL Number
 */
function generateSlNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const lastSl = sheet.getRange(lastRow, 2).getValue();
  return (Number(lastSl) || 0) + 1;
}

/**
 * Add an income record
 */
function addIncome(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Income');
    const slNumber = generateSlNumber(sheet);

    // Auto calculate total
    const roomRent = parseFloat(data.roomRent) || 0;
    const fooding = parseFloat(data.fooding) || 0;
    const total = roomRent + fooding;

    const paymentDate = (data.paymentStatus === 'PAID' || data.paymentStatus === 'ADVANCE') ? data.date : '';

    const rowData = [
      data.date,
      slNumber,
      data.entryNumber,
      data.roomNumber,
      data.other,
      roomRent,
      fooding,
      total,
      data.paymentStatus,
      data.modeOfPayment,
      data.entryBy,
      paymentDate
    ];

    sheet.appendRow(rowData);
    return { success: true, message: 'Income added successfully!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Add an expense record
 */
function addExpense(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Expenses');
    const slNumber = generateSlNumber(sheet);

    const amount = parseFloat(data.amount) || 0;
    const paymentDate = (data.paymentStatus === 'PAID' || data.paymentStatus === 'ADVANCE') ? data.date : '';

    const rowData = [
      data.date,
      slNumber,
      data.type,
      data.description,
      data.details,
      amount,
      data.paymentStatus,
      data.sourceOfPayment,
      data.modeOfPayment,
      data.entryBy,
      paymentDate
    ];

    sheet.appendRow(rowData);
    return { success: true, message: 'Expense added successfully!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Helper to get data as array of objects
 */
function getSheetDataAsObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);
  return rows.map((row, index) => {
    let obj = { _rowIndex: index + 2 }; // +2 because array is 0-indexed and header is row 1
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

/**
 * Parses date string (YYYY-MM-DD or standard JS date) into Date object safely
 */
function parseDate(dateValue) {
  if (!dateValue) return null;
  const d = new Date(dateValue);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Dashboard Calculations
 */
function getDashboardData(monthValue = '') {
  try {
    const incomeData = getSheetDataAsObjects('Income');
    const expenseData = getSheetDataAsObjects('Expenses');

    let targetMonth, targetYear;
    if (monthValue) {
      const parts = monthValue.split('-');
      targetYear = parseInt(parts[0], 10);
      targetMonth = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    } else {
      const now = new Date();
      targetMonth = now.getMonth();
      targetYear = now.getFullYear();
    }

    const now = new Date();
    const todayStr = now.toISOString().split('T')[0];

    // Calculate Financial Year (April 1 to March 31) based on current real date
    let fyStartYear = now.getFullYear();
    if (now.getMonth() < 3) { // Jan, Feb, Mar are months 0, 1, 2
      fyStartYear -= 1;
    }
    const fyStartDate = new Date(fyStartYear, 3, 1); // April 1st

    let dashboard = {
      todayIncome: 0,
      todayExpenses: 0,
      cashInCounter: 0,
      monthlyIncome: 0,
      monthlyExpenses: 0,
      netMonthlyRoomSavings: 0,
      netMonthlyFoodSavings: 0,
      totalUnpaidIncome: 0,
      totalUnpaidExpenses: 0,
      totalRoomRentIncome: 0,
      totalFoodingIncome: 0,
      totalRoomExpenses: 0,
      totalFoodExpenses: 0,
      totalFyIncome: 0,
      totalFyExpenses: 0
    };

    let totalCashCollection = 0;
    let cashExpensesFromCounter = 0;

    // Process Income
    incomeData.forEach(row => {
      const d = parseDate(row['Date']);
      const isToday = d && d.toISOString().split('T')[0] === todayStr;
      const isTargetMonth = d && d.getMonth() === targetMonth && d.getFullYear() === targetYear;
      const isFy = d && d >= fyStartDate;

      const amount = parseFloat(row['Total']) || 0;
      const roomRent = parseFloat(row['Room Rent']) || 0;
      const fooding = parseFloat(row['Fooding']) || 0;

      if (isToday) dashboard.todayIncome += amount;
      if (isFy) dashboard.totalFyIncome += amount;

      if (isTargetMonth) {
        dashboard.monthlyIncome += amount;
        dashboard.totalRoomRentIncome += roomRent;
        dashboard.totalFoodingIncome += fooding;
      }

      if (row['Payment Status'] === 'UNPAID') {
        dashboard.totalUnpaidIncome += amount;
      } else if (row['Mode Of Payment'] === 'CASH') {
        totalCashCollection += amount;
      }
    });

    // Process Expenses
    expenseData.forEach(row => {
      const d = parseDate(row['Date']);
      const isToday = d && d.toISOString().split('T')[0] === todayStr;
      const isTargetMonth = d && d.getMonth() === targetMonth && d.getFullYear() === targetYear;
      const isFy = d && d >= fyStartDate;

      const amount = parseFloat(row['Amount']) || 0;
      const type = row['Type'];

      if (isToday) dashboard.todayExpenses += amount;
      if (isFy) dashboard.totalFyExpenses += amount;

      if (isTargetMonth) {
        dashboard.monthlyExpenses += amount;
        if (type === 'ROOM') dashboard.totalRoomExpenses += amount;
        if (type === 'FOOD') dashboard.totalFoodExpenses += amount;
      }

      if (row['Payment Status'] === 'UNPAID') {
        dashboard.totalUnpaidExpenses += amount;
      } else if (row['Mode Of Payment'] === 'CASH' && row['Source Of Payment'] === 'COUNTER') {
        cashExpensesFromCounter += amount;
      }
    });

    dashboard.netMonthlyRoomSavings = dashboard.totalRoomRentIncome - dashboard.totalRoomExpenses;
    dashboard.netMonthlyFoodSavings = dashboard.totalFoodingIncome - dashboard.totalFoodExpenses;
    dashboard.cashInCounter = totalCashCollection - cashExpensesFromCounter;

    return { success: true, data: dashboard };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Get Unpaid Incomes
 */
function getUnpaidIncome(monthValue = '') {
  try {
    const data = getSheetDataAsObjects('Income');

    let targetMonth, targetYear;
    if (monthValue) {
      const parts = monthValue.split('-');
      targetYear = parseInt(parts[0], 10);
      targetMonth = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    }

    const unpaid = data.filter(r => {
      if (r['Payment Status'] !== 'UNPAID') return false;
      if (monthValue) {
        const d = parseDate(r['Date']);
        if (!d || d.getMonth() !== targetMonth || d.getFullYear() !== targetYear) {
          return false;
        }
      }
      return true;
    }).map(r => ({
      rowIndex: r._rowIndex,
      date: r['Date'] ? new Date(r['Date']).toISOString().split('T')[0] : '',
      roomNumber: r['Room Number'],
      total: r['Total'],
      entryBy: r['Entry By']
    }));
    return { success: true, data: unpaid };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Get Unpaid Expenses
 */
function getUnpaidExpenses(monthValue = '') {
  try {
    const data = getSheetDataAsObjects('Expenses');

    let targetMonth, targetYear;
    if (monthValue) {
      const parts = monthValue.split('-');
      targetYear = parseInt(parts[0], 10);
      targetMonth = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    }

    const unpaid = data.filter(r => {
      if (r['Payment Status'] !== 'UNPAID') return false;
      if (monthValue) {
        const d = parseDate(r['Date']);
        if (!d || d.getMonth() !== targetMonth || d.getFullYear() !== targetYear) {
          return false;
        }
      }
      return true;
    }).map(r => ({
      rowIndex: r._rowIndex,
      date: r['Date'] ? new Date(r['Date']).toISOString().split('T')[0] : '',
      description: r['Description'],
      amount: r['Amount'],
      entryBy: r['Entry By']
    }));
    return { success: true, data: unpaid };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mark Income Paid
 * Expected data: { rowIndex, modeOfPayment, paymentDate }
 */
function markIncomePaid(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Income');
    // Columns: Payment Status (9), Mode Of Payment (10), Payment Date (12)
    sheet.getRange(data.rowIndex, 9).setValue('PAID');
    sheet.getRange(data.rowIndex, 10).setValue(data.modeOfPayment);
    sheet.getRange(data.rowIndex, 12).setValue(data.paymentDate);
    return { success: true, message: 'Income marked as paid!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Mark Expense Paid
 * Expected data: { rowIndex, modeOfPayment, paymentDate, sourceOfPayment }
 */
function markExpensePaid(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Expenses');
    // Columns: Payment Status (7), Source Of Payment (8), Mode Of Payment (9), Payment Date (11)
    sheet.getRange(data.rowIndex, 7).setValue('PAID');
    sheet.getRange(data.rowIndex, 8).setValue(data.sourceOfPayment);
    sheet.getRange(data.rowIndex, 9).setValue(data.modeOfPayment);
    sheet.getRange(data.rowIndex, 11).setValue(data.paymentDate);
    return { success: true, message: 'Expense marked as paid!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Generate CSV for a specific month
 */
function exportCSVData(monthValue = '') {
  try {
    const incomeData = getSheetDataAsObjects('Income');
    const expenseData = getSheetDataAsObjects('Expenses');

    let targetMonth, targetYear;
    if (monthValue) {
      const parts = monthValue.split('-');
      targetYear = parseInt(parts[0], 10);
      targetMonth = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    } else {
      const now = new Date();
      targetMonth = now.getMonth();
      targetYear = now.getFullYear();
    }

    let csvContent = "Type,Date,Description/Room,Amount,Payment Status,Mode Of Payment\n";

    incomeData.forEach(row => {
      const d = parseDate(row['Date']);
      if (d && d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
        csvContent += `Income,${d.toISOString().split('T')[0]},Room ${row['Room Number']},${row['Total']},${row['Payment Status']},${row['Mode Of Payment']}\n`;
      }
    });

    expenseData.forEach(row => {
      const d = parseDate(row['Date']);
      if (d && d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
        csvContent += `Expense,${d.toISOString().split('T')[0]},${row['Description']},${row['Amount']},${row['Payment Status']},${row['Mode Of Payment']}\n`;
      }
    });

    return { success: true, csv: csvContent, filename: `Export_${targetYear}_${targetMonth+1}.csv` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
