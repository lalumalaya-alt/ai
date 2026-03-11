/**
 * Google Apps Script Backend for Rent, Electricity, and Financial Management System
 */

// Global Config
const CONFIG = {
  sheets: {
    tenants: 'Tenants',
    tenantArchive: 'Tenant_Archive',
    rentCollection: 'Rent_Collection',
    monthlySummary: 'Monthly_Summary',
    foIncome: 'F&O_Income',
    expenses: 'Expenses'
  }
};

/**
 * Standard API Response Generator
 */
function createResponse(status, message, data = null) {
  return JSON.stringify({
    status: status,
    message: message,
    data: data
  });
}

/**
 * Serve the HTML file
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Initialize Sheets if they don't exist
 */
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetDefinitions = {
    [CONFIG.sheets.tenants]: ['TenantID', 'Name', 'Mobile', 'Aadhaar', 'RentAmount', 'EB Per Unit Rate', 'Advance Paid', 'Status', 'Joined Date', 'Left Date', 'Previous Meter Reading'],
    [CONFIG.sheets.tenantArchive]: ['TenantID', 'Name', 'Mobile', 'Aadhaar', 'RentAmount', 'EB Per Unit Rate', 'Advance Paid', 'Status', 'Joined Date', 'Left Date', 'Previous Meter Reading'],
    [CONFIG.sheets.rentCollection]: ['Bill ID', 'TenantID', 'Name', 'Month', 'Rent Amount', 'EB Amount', 'Total Amount', 'Status', 'Payment Mode', 'Payment Date', 'Previous Reading', 'Current Reading', 'Units'],
    [CONFIG.sheets.monthlySummary]: ['Month', 'Total Rent', 'Total EB', 'Total Collection', 'F&O Gross', 'F&O Net', 'F&O Charges', 'Total Expenses', 'Gross PnL', 'Net PnL'],
    [CONFIG.sheets.foIncome]: ['Date', 'Broker', 'Trade Type', 'Gross PnL', 'Net PnL', 'Charges'],
    [CONFIG.sheets.expenses]: ['Date', 'Category', 'Subcategory', 'Purpose', 'Amount', 'MOP', 'Account']
  };

  try {
    for (const [sheetName, headers] of Object.entries(sheetDefinitions)) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      }
    }
    return createResponse('success', 'System initialized successfully');
  } catch (error) {
    return createResponse('error', 'Initialization failed: ' + error.message);
  }
}

/**
 * Helper to get a sheet by name
 */
function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * TENANT MANAGEMENT
 */

/**
 * Fetch all active tenants
 */
function getTenants() {
  try {
    const sheet = getSheet(CONFIG.sheets.tenants);
    if (!sheet) return createResponse('error', 'Tenants sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const tenants = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[7] !== 'Vacant') { // TenantID exists and Status is not Vacant
        let tenant = {};
        for (let j = 0; j < headers.length; j++) {
          tenant[headers[j]] = row[j];
        }
        tenants.push(tenant);
      }
    }

    return createResponse('success', 'Tenants retrieved successfully', tenants);
  } catch (e) {
    return createResponse('error', 'Error fetching tenants: ' + e.message);
  }
}

/**
 * Add a new tenant
 */
function addTenant(tenantData) {
  try {
    const sheet = getSheet(CONFIG.sheets.tenants);
    if (!sheet) return createResponse('error', 'Tenants sheet not found');

    // Generate Tenant ID (e.g., T-001)
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const tenantId = `T-${(lastRow).toString().padStart(3, '0')}`;

    const newRow = [
      tenantId,
      tenantData.name,
      tenantData.mobile,
      tenantData.aadhaar,
      tenantData.rentAmount,
      tenantData.ebRate,
      tenantData.advancePaid,
      tenantData.status || 'Active',
      tenantData.joinedDate || new Date().toISOString().split('T')[0],
      '', // Left Date empty initially
      tenantData.previousReading || 0
    ];

    sheet.appendRow(newRow);
    return createResponse('success', 'Tenant added successfully', { tenantId: tenantId });
  } catch (e) {
    return createResponse('error', 'Error adding tenant: ' + e.message);
  }
}

/**
 * Edit/Update a tenant, and handle archiving if Status changes to "Vacant"
 */
function updateTenant(tenantData) {
  try {
    const sheet = getSheet(CONFIG.sheets.tenants);
    if (!sheet) return createResponse('error', 'Tenants sheet not found');

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tenantData.tenantId) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      return createResponse('error', 'Tenant not found');
    }

    // Check if status is changed to Vacant
    if (tenantData.status === 'Vacant') {
      const rowData = data[rowIndex - 1];
      // Update specific fields before archiving
      rowData[1] = tenantData.name;
      rowData[2] = tenantData.mobile;
      rowData[3] = tenantData.aadhaar;
      rowData[4] = tenantData.rentAmount;
      rowData[5] = tenantData.ebRate;
      rowData[6] = tenantData.advancePaid;
      rowData[7] = 'Vacant';
      rowData[9] = tenantData.leftDate || new Date().toISOString().split('T')[0]; // Set Left Date
      rowData[10] = tenantData.previousReading;

      const archiveSheet = getSheet(CONFIG.sheets.tenantArchive);
      if (archiveSheet) {
        archiveSheet.appendRow(rowData);
        sheet.deleteRow(rowIndex);
        return createResponse('success', 'Tenant archived successfully');
      } else {
        return createResponse('error', 'Archive sheet not found');
      }
    } else {
      // Normal Update
      sheet.getRange(rowIndex, 2).setValue(tenantData.name);
      sheet.getRange(rowIndex, 3).setValue(tenantData.mobile);
      sheet.getRange(rowIndex, 4).setValue(tenantData.aadhaar);
      sheet.getRange(rowIndex, 5).setValue(tenantData.rentAmount);
      sheet.getRange(rowIndex, 6).setValue(tenantData.ebRate);
      sheet.getRange(rowIndex, 7).setValue(tenantData.advancePaid);
      sheet.getRange(rowIndex, 8).setValue(tenantData.status);
      sheet.getRange(rowIndex, 9).setValue(tenantData.joinedDate);
      sheet.getRange(rowIndex, 10).setValue(tenantData.leftDate || '');
      sheet.getRange(rowIndex, 11).setValue(tenantData.previousReading);

      return createResponse('success', 'Tenant updated successfully');
    }
  } catch (e) {
    return createResponse('error', 'Error updating tenant: ' + e.message);
  }
}

/**
 * METER READING & BILL GENERATION
 */

/**
 * Record meter reading, calculate bill, and append to Rent_Collection
 */
function recordMeterReading(readingData) {
  try {
    const tenantSheet = getSheet(CONFIG.sheets.tenants);
    const rentSheet = getSheet(CONFIG.sheets.rentCollection);

    if (!tenantSheet || !rentSheet) return createResponse('error', 'Required sheets not found');

    const { tenantId, currentReading, monthYear } = readingData; // monthYear format: YYYY-MM

    // 1. Fetch Tenant info and previous reading
    const data = tenantSheet.getDataRange().getValues();
    let tenantRow = -1;
    let tenantInfo = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tenantId) {
        tenantRow = i + 1;
        tenantInfo = {
          name: data[i][1],
          rentAmount: parseFloat(data[i][4]) || 0,
          ebRate: parseFloat(data[i][5]) || 0,
          previousReading: parseFloat(data[i][10]) || 0
        };
        break;
      }
    }

    if (tenantRow === -1 || !tenantInfo) return createResponse('error', 'Tenant not found');

    const currReadingNum = parseFloat(currentReading);
    if (currReadingNum < tenantInfo.previousReading) {
      return createResponse('error', 'Current reading cannot be less than previous reading');
    }

    // 2. Calculations
    const unitsConsumed = currReadingNum - tenantInfo.previousReading;
    const ebAmount = unitsConsumed * tenantInfo.ebRate;
    const totalAmount = tenantInfo.rentAmount + ebAmount;

    // 3. Generate BILL-YYYYMM-SEQ ID
    // Look at Rent Collection to find next sequence for the month
    const rentData = rentSheet.getDataRange().getValues();
    let seq = 1;
    const prefix = `BILL-${monthYear.replace('-', '')}-`; // e.g., BILL-202310-

    for (let i = 1; i < rentData.length; i++) {
      const id = rentData[i][0];
      if (id && id.toString().startsWith(prefix)) {
        const parts = id.toString().split('-');
        if (parts.length === 3) {
          const num = parseInt(parts[2], 10);
          if (num >= seq) seq = num + 1;
        }
      }
    }
    const billId = `${prefix}${seq.toString().padStart(3, '0')}`;

    // 4. Append to Rent_Collection (Status: "Unpaid")
    // Columns: ['Bill ID', 'TenantID', 'Name', 'Month', 'Rent Amount', 'EB Amount', 'Total Amount', 'Status', 'Payment Mode', 'Payment Date']
    const newRentRow = [
      billId,
      tenantId,
      tenantInfo.name,
      monthYear,
      tenantInfo.rentAmount,
      ebAmount,
      totalAmount,
      'Unpaid',
      '',
      '',
      tenantInfo.previousReading,
      currReadingNum,
      unitsConsumed
    ];
    rentSheet.appendRow(newRentRow);

    // 5. Update Column K (Previous Reading) in Tenant sheet
    tenantSheet.getRange(tenantRow, 11).setValue(currReadingNum);

    return createResponse('success', 'Bill generated successfully', {
      billId: billId,
      units: unitsConsumed,
      ebAmount: ebAmount,
      totalAmount: totalAmount,
      previousReading: tenantInfo.previousReading,
      currentReading: currReadingNum
    });

  } catch (e) {
    return createResponse('error', 'Error generating bill: ' + e.message);
  }
}

/**
 * PAYMENT PROCESSING
 */

/**
 * Mark a bill as "Paid"
 */
function processPayment(paymentData) {
  try {
    const sheet = getSheet(CONFIG.sheets.rentCollection);
    if (!sheet) return createResponse('error', 'Rent Collection sheet not found');

    const { billId, month, paymentMode, paymentDate } = paymentData;

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === billId && data[i][3] === month) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      return createResponse('error', 'Bill not found');
    }

    // Update Rent_Collection row
    // Columns: ['Bill ID', 'TenantID', 'Name', 'Month', 'Rent Amount', 'EB Amount', 'Total Amount', 'Status', 'Payment Mode', 'Payment Date']
    sheet.getRange(rowIndex, 8).setValue('Paid');
    sheet.getRange(rowIndex, 9).setValue(paymentMode);
    sheet.getRange(rowIndex, 10).setValue(paymentDate);

    return createResponse('success', 'Payment processed successfully');
  } catch (e) {
    return createResponse('error', 'Error processing payment: ' + e.message);
  }
}

/**
 * Get all bills (paid and unpaid)
 */
function getBills() {
  try {
    const sheet = getSheet(CONFIG.sheets.rentCollection);
    if (!sheet) return createResponse('error', 'Rent Collection sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const bills = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // Bill ID exists
        let bill = {};
        for (let j = 0; j < headers.length; j++) {
          bill[headers[j]] = data[i][j];
        }
        bills.push(bill);
      }
    }

    return createResponse('success', 'Bills retrieved successfully', bills);
  } catch (e) {
    return createResponse('error', 'Error fetching bills: ' + e.message);
  }
}

/**
 * FINANCIAL TRACKING
 */

/**
 * Add F&O Trade Income/Loss
 */
function addFOIncome(tradeData) {
  try {
    const sheet = getSheet(CONFIG.sheets.foIncome);
    if (!sheet) return createResponse('error', 'F&O Income sheet not found');

    const { date, broker, tradeType, grossPnL, netPnL } = tradeData;
    const gross = parseFloat(grossPnL) || 0;
    const net = parseFloat(netPnL) || 0;
    const charges = gross - net; // Auto-calculate charges

    sheet.appendRow([date, broker, tradeType, gross, net, charges]);

    return createResponse('success', 'Trade recorded successfully');
  } catch (e) {
    return createResponse('error', 'Error recording trade: ' + e.message);
  }
}

/**
 * Get F&O Trades
 */
function getFOTrades() {
  try {
    const sheet = getSheet(CONFIG.sheets.foIncome);
    if (!sheet) return createResponse('error', 'F&O Income sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const trades = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        let trade = {};
        for (let j = 0; j < headers.length; j++) {
          trade[headers[j]] = data[i][j];
        }
        trades.push(trade);
      }
    }
    return createResponse('success', 'Trades retrieved successfully', trades);
  } catch (e) {
    return createResponse('error', 'Error fetching trades: ' + e.message);
  }
}

/**
 * Add Expense
 */
function addExpense(expenseData) {
  try {
    const sheet = getSheet(CONFIG.sheets.expenses);
    if (!sheet) return createResponse('error', 'Expenses sheet not found');

    const { date, category, subcategory, purpose, amount, mop, account } = expenseData;

    sheet.appendRow([date, category, subcategory, purpose, parseFloat(amount), mop, account]);

    return createResponse('success', 'Expense recorded successfully');
  } catch (e) {
    return createResponse('error', 'Error recording expense: ' + e.message);
  }
}

/**
 * Get Expenses
 */
function getExpenses() {
  try {
    const sheet = getSheet(CONFIG.sheets.expenses);
    if (!sheet) return createResponse('error', 'Expenses sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const expenses = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        let expense = {};
        for (let j = 0; j < headers.length; j++) {
          expense[headers[j]] = data[i][j];
        }
        expenses.push(expense);
      }
    }
    return createResponse('success', 'Expenses retrieved successfully', expenses);
  } catch (e) {
    return createResponse('error', 'Error fetching expenses: ' + e.message);
  }
}

/**
 * MONTHLY SUMMARY AUTOMATION
 */

/**
 * Rebuild the Monthly Summary sheet
 */
function rebuildMonthlySummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName(CONFIG.sheets.monthlySummary);
    const rentSheet = ss.getSheetByName(CONFIG.sheets.rentCollection);
    const foSheet = ss.getSheetByName(CONFIG.sheets.foIncome);
    const expSheet = ss.getSheetByName(CONFIG.sheets.expenses);

    if (!summarySheet || !rentSheet || !foSheet || !expSheet) {
      return createResponse('error', 'One or more required sheets missing');
    }

    // Monthly data structure: { 'YYYY-MM': { totalRent: 0, totalEB: 0, totalCol: 0, foGross: 0, foNet: 0, foCharges: 0, totalExp: 0 } }
    const monthlyData = {};

    // Process Rent Collection (Only Paid ones count towards Collection)
    const rentData = rentSheet.getDataRange().getValues();
    for (let i = 1; i < rentData.length; i++) {
      const month = rentData[i][3]; // e.g., '2023-10'
      if (!month) continue;

      if (!monthlyData[month]) {
        monthlyData[month] = { totalRent: 0, totalEB: 0, totalCol: 0, foGross: 0, foNet: 0, foCharges: 0, totalExp: 0 };
      }

      const rentAmt = parseFloat(rentData[i][4]) || 0;
      const ebAmt = parseFloat(rentData[i][5]) || 0;
      const totalAmt = parseFloat(rentData[i][6]) || 0;
      const status = rentData[i][7];

      monthlyData[month].totalRent += rentAmt;
      monthlyData[month].totalEB += ebAmt;

      if (status === 'Paid') {
        monthlyData[month].totalCol += totalAmt;
      }
    }

    // Process F&O Income
    const foData = foSheet.getDataRange().getValues();
    for (let i = 1; i < foData.length; i++) {
      let dateVal = foData[i][0];
      if (!dateVal) continue;

      // Handle Date object or string
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      const monthStr = dateVal.toString().substring(0, 7); // 'YYYY-MM'

      if (!monthlyData[monthStr]) {
        monthlyData[monthStr] = { totalRent: 0, totalEB: 0, totalCol: 0, foGross: 0, foNet: 0, foCharges: 0, totalExp: 0 };
      }

      monthlyData[monthStr].foGross += parseFloat(foData[i][3]) || 0;
      monthlyData[monthStr].foNet += parseFloat(foData[i][4]) || 0;
      monthlyData[monthStr].foCharges += parseFloat(foData[i][5]) || 0;
    }

    // Process Expenses
    const expData = expSheet.getDataRange().getValues();
    for (let i = 1; i < expData.length; i++) {
      let dateVal = expData[i][0];
      if (!dateVal) continue;

      // Handle Date object or string
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      const monthStr = dateVal.toString().substring(0, 7);

      if (!monthlyData[monthStr]) {
        monthlyData[monthStr] = { totalRent: 0, totalEB: 0, totalCol: 0, foGross: 0, foNet: 0, foCharges: 0, totalExp: 0 };
      }

      monthlyData[monthStr].totalExp += parseFloat(expData[i][4]) || 0;
    }

    // Update Summary Sheet
    // Clear old data except headers
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, summarySheet.getLastColumn()).clearContent();
    }

    const sortedMonths = Object.keys(monthlyData).sort().reverse(); // Most recent first

    if (sortedMonths.length > 0) {
      const outputData = [];
      for (const month of sortedMonths) {
        const d = monthlyData[month];
        const grossPnL = d.totalCol + d.foGross - d.totalExp;
        const netPnL = d.totalCol + d.foNet - d.totalExp;

        outputData.push([
          month,
          d.totalRent,
          d.totalEB,
          d.totalCol,
          d.foGross,
          d.foNet,
          d.foCharges,
          d.totalExp,
          grossPnL,
          netPnL
        ]);
      }

      summarySheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    }

    return createResponse('success', 'Monthly Summary rebuilt successfully');
  } catch (e) {
    return createResponse('error', 'Error rebuilding summary: ' + e.message);
  }
}

/**
 * Automatically set up sheets when the Google Sheet is opened
 */
function onOpen() {
  initializeSystem();
}

/**
 * Triggers - Setup from Apps Script dashboard (Edit -> Current project's triggers)
 * We can also provide a programmatic way to install them.
 */
function onEditTrigger(e) {
  // Simple trigger wrapper. If editing Rent, F&O, or Expenses, rebuild summary.
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();
  if ([CONFIG.sheets.rentCollection, CONFIG.sheets.foIncome, CONFIG.sheets.expenses].includes(sheetName)) {
    rebuildMonthlySummary();
  }
}

/**
 * Central router for google.script.run calls from frontend
 */
function processRequest(action, data) {
  try {
    switch (action) {
      case 'init': return initializeSystem();
      case 'getTenants': return getTenants();
      case 'addTenant': return addTenant(data);
      case 'updateTenant': return updateTenant(data);
      case 'recordReading': return recordMeterReading(data);
      case 'processPayment': return processPayment(data);
      case 'getBills': return getBills();
      case 'addFOIncome': return addFOIncome(data);
      case 'getFOTrades': return getFOTrades();
      case 'addExpense': return addExpense(data);
      case 'getExpenses': return getExpenses();
      case 'rebuildSummary': return rebuildMonthlySummary();
      default: return createResponse('error', 'Unknown action');
    }
  } catch (error) {
    return createResponse('error', 'Server error: ' + error.message);
  }
}
