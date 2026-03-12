/**
 * Google Apps Script Backend for Rent & Financial Management System
 */

// Helper function to return standardized JSON
function createResponse(status, message, data) {
  return JSON.stringify({
    status: status,
    message: message,
    data: data || {}
  });
}

// -----------------------------------------------------------------------------
// INITIALIZATION
// -----------------------------------------------------------------------------

function initializeSystem() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = [
      "Tenants",
      "Tenant_Archive",
      "Rent_Collection",
      "Monthly_Summary",
      "F&O_Income",
      "Expenses"
    ];

    sheets.forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
      }
    });

    // Set headers for Tenants
    const tenantHeaders = ["TenantID", "Name", "Mobile", "Aadhaar", "RentAmount", "EB Per Unit Rate", "Advance Paid", "Status", "Joined Date", "Left Date", "Previous Meter Reading"];
    const tenantSheet = ss.getSheetByName("Tenants");
    if (tenantSheet.getLastRow() === 0) {
      tenantSheet.appendRow(tenantHeaders);
    }

    // Set headers for Tenant_Archive
    const archiveSheet = ss.getSheetByName("Tenant_Archive");
    if (archiveSheet.getLastRow() === 0) {
      archiveSheet.appendRow(tenantHeaders);
    }

    // Set headers for Rent_Collection
    const rentHeaders = ["Bill ID", "TenantID", "Name", "Month", "Previous Reading", "Current Reading", "Units Consumed", "EB Amount", "Rent Amount", "Total Amount", "Status", "Payment Mode", "Payment Date"];
    const rentSheet = ss.getSheetByName("Rent_Collection");
    if (rentSheet.getLastRow() === 0) {
      rentSheet.appendRow(rentHeaders);
    }

    // Set headers for F&O_Income
    const foHeaders = ["Date", "Broker", "Gross", "Net", "Charges"];
    const foSheet = ss.getSheetByName("F&O_Income");
    if (foSheet.getLastRow() === 0) {
      foSheet.appendRow(foHeaders);
    }

    // Set headers for Expenses
    const expenseHeaders = ["Date", "Category", "Subcategory", "Purpose", "Amount", "MOP", "Account"];
    const expenseSheet = ss.getSheetByName("Expenses");
    if (expenseSheet.getLastRow() === 0) {
      expenseSheet.appendRow(expenseHeaders);
    }

    // Set headers for Monthly_Summary
    const summaryHeaders = ["Month", "Total Rent", "Total EB", "Total Collection", "Gross PnL", "Net PnL", "Expenses", "Net Savings"];
    const summarySheet = ss.getSheetByName("Monthly_Summary");
    if (summarySheet.getLastRow() === 0) {
      summarySheet.appendRow(summaryHeaders);
    }

    return createResponse("success", "System initialized successfully.");
  } catch (e) {
    return createResponse("error", "Failed to initialize system: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// TENANT MANAGEMENT
// -----------------------------------------------------------------------------

function getTenants() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenants");
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse("success", "No tenants found.", []);

    const headers = data[0];
    const tenants = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tenant = {};
      headers.forEach((header, index) => {
        tenant[header] = row[index];
      });
      tenant.rowNumber = i + 1;
      tenants.push(tenant);
    }

    return createResponse("success", "Tenants retrieved successfully.", tenants);
  } catch (e) {
    return createResponse("error", "Error getting tenants: " + e.message);
  }
}

function addTenant(tenantData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenants");
    const tenantId = "T-" + new Date().getTime();

    // Columns: "TenantID", "Name", "Mobile", "Aadhaar", "RentAmount", "EB Per Unit Rate", "Advance Paid", "Status", "Joined Date", "Left Date", "Previous Meter Reading"
    const newRow = [
      tenantId,
      tenantData.name,
      tenantData.mobile,
      tenantData.aadhaar,
      tenantData.rentAmount,
      tenantData.ebRate,
      tenantData.advancePaid,
      "Occupied",
      tenantData.joinedDate || new Date(),
      "", // Left Date
      tenantData.previousMeterReading || 0
    ];

    sheet.appendRow(newRow);
    return createResponse("success", "Tenant added successfully.", { tenantId: tenantId });
  } catch (e) {
    return createResponse("error", "Error adding tenant: " + e.message);
  }
}

function editTenant(tenantData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenants");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tenantData.tenantId) {
        const row = i + 1;
        // Update columns
        sheet.getRange(row, 2).setValue(tenantData.name);
        sheet.getRange(row, 3).setValue(tenantData.mobile);
        sheet.getRange(row, 4).setValue(tenantData.aadhaar);
        sheet.getRange(row, 5).setValue(tenantData.rentAmount);
        sheet.getRange(row, 6).setValue(tenantData.ebRate);
        sheet.getRange(row, 7).setValue(tenantData.advancePaid);
        sheet.getRange(row, 8).setValue(tenantData.status);

        if (tenantData.status === "Vacant") {
          return archiveTenant(tenantData.tenantId, tenantData.leftDate || new Date());
        }

        return createResponse("success", "Tenant updated successfully.");
      }
    }
    return createResponse("error", "Tenant not found.");
  } catch (e) {
    return createResponse("error", "Error editing tenant: " + e.message);
  }
}

function archiveTenant(tenantId, leftDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tenantSheet = ss.getSheetByName("Tenants");
    const archiveSheet = ss.getSheetByName("Tenant_Archive");

    const data = tenantSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tenantId) {
        const rowData = data[i];
        rowData[7] = "Vacant"; // Status
        rowData[9] = leftDate || new Date(); // Left Date

        archiveSheet.appendRow(rowData);
        tenantSheet.deleteRow(i + 1);

        return createResponse("success", "Tenant archived successfully.");
      }
    }
    return createResponse("error", "Tenant not found for archiving.");
  } catch (e) {
    return createResponse("error", "Error archiving tenant: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// METER READING & BILLING
// -----------------------------------------------------------------------------

function getUnpaidBills() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rent_Collection");
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse("success", "No bills found.", []);

    const headers = data[0];
    const bills = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][10] === "Unpaid") { // Status column
        const bill = {};
        headers.forEach((header, index) => {
          bill[header] = data[i][index];
        });
        bills.push(bill);
      }
    }

    return createResponse("success", "Unpaid bills retrieved successfully.", bills);
  } catch (e) {
    return createResponse("error", "Error getting bills: " + e.message);
  }
}

function recordMeterReading(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tenantSheet = ss.getSheetByName("Tenants");
    const rentSheet = ss.getSheetByName("Rent_Collection");

    const tenantData = tenantSheet.getDataRange().getValues();
    let tenantRowIdx = -1;
    let tenantRow = null;

    for (let i = 1; i < tenantData.length; i++) {
      if (tenantData[i][0] === data.tenantId) {
        tenantRowIdx = i + 1;
        tenantRow = tenantData[i];
        break;
      }
    }

    if (!tenantRow) return createResponse("error", "Tenant not found.");

    // Calculate Bill
    const previousReading = tenantRow[10] || 0; // Column K
    const ebRate = tenantRow[5] || 0;
    const rentAmount = tenantRow[4] || 0;
    const currentReading = parseFloat(data.currentReading);

    if (currentReading < previousReading) {
      return createResponse("error", "Current reading cannot be less than previous reading.");
    }

    const unitsConsumed = currentReading - previousReading;
    const ebAmount = unitsConsumed * ebRate;
    const totalAmount = rentAmount + ebAmount;

    // Generate Bill ID: BILL-YYYYMM-SEQ
    const dateObj = new Date();
    const yyyymm = dateObj.getFullYear() + String(dateObj.getMonth() + 1).padStart(2, '0');
    // Basic SEQ generation based on timestamp for uniqueness
    const seq = String(dateObj.getTime()).slice(-4);
    const billId = `BILL-${yyyymm}-${seq}`;
    const month = dateObj.toLocaleString('default', { month: 'long', year: 'numeric' });

    // Rent_Collection Headers: "Bill ID", "TenantID", "Name", "Month", "Previous Reading", "Current Reading", "Units Consumed", "EB Amount", "Rent Amount", "Total Amount", "Status", "Payment Mode", "Payment Date"
    const newBillRow = [
      billId,
      data.tenantId,
      tenantRow[1], // Name
      month,
      previousReading,
      currentReading,
      unitsConsumed,
      ebAmount,
      rentAmount,
      totalAmount,
      "Unpaid",
      "", // Payment Mode
      ""  // Payment Date
    ];

    rentSheet.appendRow(newBillRow);

    // Update Previous Reading in Tenant Sheet
    tenantSheet.getRange(tenantRowIdx, 11).setValue(currentReading);

    return createResponse("success", "Bill generated successfully.", { billId: billId, totalAmount: totalAmount });
  } catch (e) {
    return createResponse("error", "Error recording reading: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// PAYMENT PROCESSING
// -----------------------------------------------------------------------------

function processPayment(paymentData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rent_Collection");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === paymentData.billId) { // Bill ID match
        const row = i + 1;

        sheet.getRange(row, 11).setValue("Paid");
        sheet.getRange(row, 12).setValue(paymentData.paymentMode);
        sheet.getRange(row, 13).setValue(paymentData.paymentDate || new Date());

        // Rebuild summary after payment
        rebuildMonthlySummary();

        return createResponse("success", "Payment processed successfully.");
      }
    }

    return createResponse("error", "Bill not found.");
  } catch (e) {
    return createResponse("error", "Error processing payment: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// FINANCIAL TRACKING
// -----------------------------------------------------------------------------

function addFOIncome(foData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("F&O_Income");
    const gross = parseFloat(foData.gross) || 0;
    const net = parseFloat(foData.net) || 0;
    const charges = gross - net;

    // Headers: "Date", "Broker", "Gross", "Net", "Charges"
    const newRow = [
      foData.date || new Date(),
      foData.broker,
      gross,
      net,
      charges
    ];

    sheet.appendRow(newRow);
    rebuildMonthlySummary();
    return createResponse("success", "F&O income added successfully.");
  } catch (e) {
    return createResponse("error", "Error adding F&O income: " + e.message);
  }
}

function addExpense(expenseData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
    const amount = parseFloat(expenseData.amount) || 0;

    // Headers: "Date", "Category", "Subcategory", "Purpose", "Amount", "MOP", "Account"
    const newRow = [
      expenseData.date || new Date(),
      expenseData.category, // Personal/Trading
      expenseData.subcategory,
      expenseData.purpose,
      amount,
      expenseData.mop,
      expenseData.account
    ];

    sheet.appendRow(newRow);
    rebuildMonthlySummary();
    return createResponse("success", "Expense added successfully.");
  } catch (e) {
    return createResponse("error", "Error adding expense: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// MONTHLY SUMMARY AUTOMATION
// -----------------------------------------------------------------------------

function getMonthYearString(dateVal) {
  if (!dateVal) return "";
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return "";
  return d.toLocaleString('default', { month: 'long', year: 'numeric' });
}

function rebuildMonthlySummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const rentSheet = ss.getSheetByName("Rent_Collection");
    const foSheet = ss.getSheetByName("F&O_Income");
    const expSheet = ss.getSheetByName("Expenses");
    const sumSheet = ss.getSheetByName("Monthly_Summary");

    const rentData = rentSheet.getDataRange().getValues();
    const foData = foSheet.getDataRange().getValues();
    const expData = expSheet.getDataRange().getValues();

    const summaryMap = {}; // Key: "Month Year", Value: { rent, eb, coll, grossPnl, netPnl, exp }

    // Process Rent Collection
    for (let i = 1; i < rentData.length; i++) {
      const row = rentData[i];
      const month = row[3]; // Month String
      const status = row[10];
      const eb = parseFloat(row[7]) || 0;
      const rent = parseFloat(row[8]) || 0;
      const total = parseFloat(row[9]) || 0;

      if (!summaryMap[month]) summaryMap[month] = { rent:0, eb:0, coll:0, grossPnl:0, netPnl:0, exp:0 };

      summaryMap[month].rent += rent;
      summaryMap[month].eb += eb;
      if (status === "Paid") {
        summaryMap[month].coll += total;
      }
    }

    // Process F&O Income
    for (let i = 1; i < foData.length; i++) {
      const row = foData[i];
      const month = getMonthYearString(row[0]);
      const gross = parseFloat(row[2]) || 0;
      const net = parseFloat(row[3]) || 0;

      if (month) {
        if (!summaryMap[month]) summaryMap[month] = { rent:0, eb:0, coll:0, grossPnl:0, netPnl:0, exp:0 };
        summaryMap[month].grossPnl += gross;
        summaryMap[month].netPnl += net;
      }
    }

    // Process Expenses
    for (let i = 1; i < expData.length; i++) {
      const row = expData[i];
      const month = getMonthYearString(row[0]);
      const amount = parseFloat(row[4]) || 0;

      if (month) {
        if (!summaryMap[month]) summaryMap[month] = { rent:0, eb:0, coll:0, grossPnl:0, netPnl:0, exp:0 };
        summaryMap[month].exp += amount;
      }
    }

    // Clear existing summary
    const sumRange = sumSheet.getDataRange();
    if (sumRange.getNumRows() > 1) {
      sumSheet.getRange(2, 1, sumRange.getNumRows() - 1, sumRange.getNumColumns()).clearContent();
    }

    // Rebuild rows
    for (const month in summaryMap) {
      const s = summaryMap[month];
      const netSavings = s.coll + s.netPnl - s.exp; // Simplified net savings calc

      // Headers: ["Month", "Total Rent", "Total EB", "Total Collection", "Gross PnL", "Net PnL", "Expenses", "Net Savings"]
      sumSheet.appendRow([
        month, s.rent, s.eb, s.coll, s.grossPnl, s.netPnl, s.exp, netSavings
      ]);
    }

    return createResponse("success", "Summary rebuilt successfully.");
  } catch (e) {
    return createResponse("error", "Error rebuilding summary: " + e.message);
  }
}

// Triggers for automatic sync
function onEditTrigger(e) {
  // Can add specific sheet checks, but for now we'll rebuild generally
  rebuildMonthlySummary();
}

// -----------------------------------------------------------------------------
// DASHBOARD DATA
// -----------------------------------------------------------------------------

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tenantSheet = ss.getSheetByName("Tenants");
    const summarySheet = ss.getSheetByName("Monthly_Summary");
    const rentSheet = ss.getSheetByName("Rent_Collection");

    const tenantData = tenantSheet.getDataRange().getValues();
    const summaryData = summarySheet.getDataRange().getValues();
    const rentData = rentSheet.getDataRange().getValues();

    let occupied = 0;
    let vacant = 0; // Historically tracked if they stay in main sheet, otherwise total houses = occupied

    for (let i = 1; i < tenantData.length; i++) {
      if (tenantData[i][7] === "Occupied") occupied++;
      if (tenantData[i][7] === "Vacant") vacant++;
    }

    let pendingBills = 0;
    for (let i = 1; i < rentData.length; i++) {
      if (rentData[i][10] === "Unpaid") pendingBills++;
    }

    // Get current month summary stats (last row ideally, or match current month)
    const currentMonthString = new Date().toLocaleString('default', { month: 'long', year: 'numeric' });
    let rentReceived = 0, tradingPnl = 0, expenses = 0, netSavings = 0;

    for (let i = 1; i < summaryData.length; i++) {
      if (summaryData[i][0] === currentMonthString) {
        rentReceived = summaryData[i][3]; // Total Collection
        tradingPnl = summaryData[i][5];   // Net PnL
        expenses = summaryData[i][6];     // Expenses
        netSavings = summaryData[i][7];   // Net Savings
        break;
      }
    }

    const data = {
      kpi: {
        totalHouses: occupied + vacant, // Approximate logic
        occupied: occupied,
        vacant: vacant,
        pending: pendingBills
      },
      stats: {
        rentReceived: rentReceived,
        tradingPnl: tradingPnl,
        expenses: expenses,
        netSavings: netSavings
      },
      recentActivity: [] // Could be populated from last 5 rows of various sheets
    };

    return createResponse("success", "Dashboard data retrieved successfully.", data);
  } catch (e) {
    return createResponse("error", "Error getting dashboard data: " + e.message);
  }
}

// -----------------------------------------------------------------------------
// API / UI RENDER
// -----------------------------------------------------------------------------

function doGet() {
  // Ensure the system is initialized on load just in case
  initializeSystem();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Rent & Financial Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}
