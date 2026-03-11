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
    'Date', 'Entry Number', 'Room Number', 'Other',
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
    'Date', 'Type', 'Description', 'Details', 'Amount',
    'Payment Status', 'Source Of Payment', 'Mode Of Payment',
    'Entry By', 'Payment Date'
  ];
  expenseSheet.getRange(1, 1, 1, expenseHeaders.length).setValues([expenseHeaders]);
  expenseSheet.getRange(1, 1, 1, expenseHeaders.length).setFontWeight('bold');
}

/**
 * Helper to sort a sheet by the Date column (Column A)
 */
function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow > 1) {
    // Sort everything from row 2 (skipping header) by column 1 (Date)
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    range.sort({ column: 1, ascending: true });
  }
}

/**
 * Add an income record
 */
function addIncome(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Income');

    // Auto calculate total
    const roomRent = parseFloat(data.roomRent) || 0;
    const fooding = parseFloat(data.fooding) || 0;
    const total = roomRent + fooding;

    const paymentDate = (data.paymentStatus === 'PAID' || data.paymentStatus === 'ADVANCE') ? data.date : '';

    const rowData = [
      data.date,
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
    sortSheetByDate(sheet);
    return { success: true, message: 'Saved successfully' };
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

    const amount = parseFloat(data.amount) || 0;
    const paymentDate = (data.paymentStatus === 'PAID' || data.paymentStatus === 'ADVANCE') ? data.date : '';

    const rowData = [
      data.date,
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
    sortSheetByDate(sheet);
    return { success: true, message: 'Saved successfully' };
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
 * Helper to safely format a Date object to YYYY-MM-DD using the script's timezone
 * This prevents the "one day back" bug caused by .toISOString() converting local midnight to previous day UTC
 */
function formatDateToLocal(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return '';
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
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
    const todayStr = formatDateToLocal(now);

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
      totalFyExpenses: 0,
      cashTakenMalaya: 0,
      cashTakenMDSir: 0
    };

    let totalCashCollection = 0;
    let cashExpensesFromCounter = 0;

    // Process Income
    incomeData.forEach(row => {
      const d = parseDate(row['Date']);
      const isToday = d && formatDateToLocal(d) === todayStr;
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
      const isToday = d && formatDateToLocal(d) === todayStr;
      const isTargetMonth = d && d.getMonth() === targetMonth && d.getFullYear() === targetYear;
      const isFy = d && d >= fyStartDate;

      const amount = parseFloat(row['Amount']) || 0;
      const type = row['Type'];
      const description = row['Description'];

      if (isToday) dashboard.todayExpenses += amount;
      if (isFy) dashboard.totalFyExpenses += amount;

      if (isTargetMonth) {
        dashboard.monthlyExpenses += amount;
        if (type === 'ROOM') dashboard.totalRoomExpenses += amount;
        if (type === 'FOOD') dashboard.totalFoodExpenses += amount;
        if (description === 'Malaya') dashboard.cashTakenMalaya += amount;
        if (description === 'MD Sir') dashboard.cashTakenMDSir += amount;
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
      date: r['Date'] ? formatDateToLocal(new Date(r['Date'])) : '',
      entryNumber: r['Entry Number'],
      roomNumber: r['Room Number'],
      roomRent: r['Room Rent'],
      fooding: r['Fooding'],
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
      date: r['Date'] ? formatDateToLocal(new Date(r['Date'])) : '',
      type: r['Type'],
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
 * Get Report Data (All rows based on Category and Month)
 */
function getReportData(category, monthValue, expenseType = '', expenseDesc = '') {
  try {
    const sheetName = category === 'INCOME' ? 'Income' : 'Expenses';
    const data = getSheetDataAsObjects(sheetName);

    let targetMonth, targetYear;
    if (monthValue) {
      const parts = monthValue.split('-');
      targetYear = parseInt(parts[0], 10);
      targetMonth = parseInt(parts[1], 10) - 1; // JS months are 0-indexed
    }

    const filteredData = data.filter(r => {
      if (monthValue) {
        const d = parseDate(r['Date']);
        if (!d || d.getMonth() !== targetMonth || d.getFullYear() !== targetYear) {
          return false;
        }
      }
      if (category === 'EXPENSE') {
        if (expenseType && r['Type'] !== expenseType) return false;
        if (expenseDesc && r['Description'] !== expenseDesc) return false;
      }
      return true;
    }).map(r => {
      // Map all columns for the frontend
      if (category === 'INCOME') {
        return {
          date: r['Date'] ? formatDateToLocal(new Date(r['Date'])) : '',
          entryNumber: r['Entry Number'],
          roomNumber: r['Room Number'],
          other: r['Other'],
          roomRent: r['Room Rent'],
          fooding: r['Fooding'],
          total: r['Total'],
          paymentStatus: r['Payment Status'],
          modeOfPayment: r['Mode Of Payment'],
          entryBy: r['Entry By'],
          paymentDate: r['Payment Date'] ? formatDateToLocal(new Date(r['Payment Date'])) : ''
        };
      } else {
        return {
          date: r['Date'] ? formatDateToLocal(new Date(r['Date'])) : '',
          type: r['Type'],
          description: r['Description'],
          details: r['Details'],
          amount: r['Amount'],
          paymentStatus: r['Payment Status'],
          sourceOfPayment: r['Source Of Payment'],
          modeOfPayment: r['Mode Of Payment'],
          entryBy: r['Entry By'],
          paymentDate: r['Payment Date'] ? formatDateToLocal(new Date(r['Payment Date'])) : ''
        };
      }
    });

    return { success: true, data: filteredData };
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
    // Columns: Payment Status (8), Mode Of Payment (9), Payment Date (11)
    sheet.getRange(data.rowIndex, 8).setValue('PAID');
    sheet.getRange(data.rowIndex, 9).setValue(data.modeOfPayment);
    sheet.getRange(data.rowIndex, 11).setValue(data.paymentDate);
    return { success: true, message: 'Saved successfully' };
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
    // Columns: Payment Status (6), Source Of Payment (7), Mode Of Payment (8), Payment Date (10)
    sheet.getRange(data.rowIndex, 6).setValue('PAID');
    sheet.getRange(data.rowIndex, 7).setValue(data.sourceOfPayment);
    sheet.getRange(data.rowIndex, 8).setValue(data.modeOfPayment);
    sheet.getRange(data.rowIndex, 10).setValue(data.paymentDate);
    return { success: true, message: 'Saved successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Close Financial Year
 * Archives all records with a Date older than April 1 of the CURRENT Financial Year
 * to separate 'Income_Archive' and 'Expenses_Archive' sheets.
 */
function closeFinancialYear() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    // Calculate start of current financial year (April 1st)
    let fyStartYear = now.getFullYear();
    if (now.getMonth() < 3) { // Jan, Feb, Mar
      fyStartYear -= 1;
    }
    const currentFyStartDate = new Date(fyStartYear, 3, 1); // April 1st of current FY

    // Archive logic helper
    function archiveSheet(sheetName, archiveSheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return 0;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1) return 0; // Only headers

      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const data = dataRange.getValues();

      let rowsToArchive = [];
      let rowsToDelete = [];

      // Identify rows to archive (Date < currentFyStartDate)
      for (let i = 0; i < data.length; i++) {
        const rowDateStr = data[i][0]; // Column A is Date
        const rowDate = parseDate(rowDateStr);
        if (rowDate && rowDate < currentFyStartDate) {
          rowsToArchive.push(data[i]);
          rowsToDelete.push(i + 2); // +2 because data array is 0-indexed and row 1 is header
        }
      }

      if (rowsToArchive.length > 0) {
        // Get or create archive sheet
        let archiveSheet = ss.getSheetByName(archiveSheetName);
        if (!archiveSheet) {
          archiveSheet = ss.insertSheet(archiveSheetName);
          // Copy headers
          const headers = sheet.getRange(1, 1, 1, lastCol).getValues();
          archiveSheet.getRange(1, 1, 1, lastCol).setValues(headers);
          archiveSheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');
        }

        // Append all identified rows to archive
        const startRow = archiveSheet.getLastRow() + 1;
        archiveSheet.getRange(startRow, 1, rowsToArchive.length, lastCol).setValues(rowsToArchive);

        // Delete rows from the main sheet (Must iterate backwards to avoid shifting indices)
        for (let j = rowsToDelete.length - 1; j >= 0; j--) {
          sheet.deleteRow(rowsToDelete[j]);
        }
      }
      return rowsToArchive.length;
    }

    const archivedIncomeCount = archiveSheet('Income', 'Income_Archive');
    const archivedExpenseCount = archiveSheet('Expenses', 'Expenses_Archive');

    const totalArchived = archivedIncomeCount + archivedExpenseCount;

    if (totalArchived === 0) {
       return { success: true, message: 'No previous financial year data found to close.' };
    }

    return { success: true, message: `Financial Year closed successfully! Archived ${archivedIncomeCount} Income records and ${archivedExpenseCount} Expense records.` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
