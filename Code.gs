/*************************************************
 RENT + ELECTRICITY MANAGEMENT SYSTEM
 PRODUCTION READY BACKEND - FIXED VERSION
 Last Updated: 2026-03-04
 Feature: Tenant Sheet Column K Integration + Payment Date
*************************************************/

const SHEETS = {
  TENANTS: "Tenants",
  TENANT_ARCHIVE: "Tenant_Archive",
  RENT: "Rent_Collection",
  SUMMARY: "Monthly_Summary",
  FO_INCOME: "F&O_Income",
  EXPENSES: "Expenses",
  STAFF: "Staff",
  SALARY: "Salary_Payouts",
  STAFF_ADVANCES: "Staff_Advances"
};

const STAFF_COLUMNS = {
  STAFF_ID: 0,
  NAME: 1,
  COMPANY: 2,
  SALARY: 3,
  BANK_NAME: 4,
  ACCOUNT_NAME: 5,
  ACCOUNT_NUMBER: 6,
  ADVANCE_BALANCE: 7,
  STATUS: 8,
  JOINED_DATE: 9,
  LEFT_DATE: 10
};

const STAFF_HEADER = [
  "StaffID",
  "Name",
  "Company",
  "Salary",
  "Bank Name",
  "Account Name",
  "Account Number",
  "Advance Balance",
  "Status",
  "Joined Date",
  "Left Date"
];

const SALARY_COLUMNS = {
  DATE: 0,
  TRANSACTION_ID: 1,
  STAFF_ID: 2,
  NAME: 3,
  MONTH: 4,
  BASE_SALARY: 5,
  DAYS_WORKED: 6,
  EARNED_SALARY: 7,
  BONUS: 8,
  ADVANCE_DEDUCTED: 9,
  NET_PAID: 10,
  MOP: 11,
  SOP: 12
};

const SALARY_HEADER = [
  "Date",
  "TransactionID",
  "StaffID",
  "Name",
  "Month",
  "Base Salary",
  "Days Worked",
  "Earned Salary",
  "Bonus",
  "Advance Deducted",
  "Net Paid",
  "MOP",
  "SOP"
];

const ADVANCE_COLUMNS = {
  DATE: 0,
  TRANSACTION_ID: 1,
  STAFF_ID: 2,
  NAME: 3,
  TYPE: 4,
  AMOUNT: 5,
  MOP: 6,
  DESCRIPTION: 7
};

const ADVANCE_HEADER = [
  "Date",
  "TransactionID",
  "StaffID",
  "Name",
  "Type",
  "Amount",
  "MOP",
  "Description"
];

const TENANT_COLUMNS = {
  TENANT_ID: 0,
  NAME: 1,
  MOBILE: 2,
  AADHAAR: 3,
  RENT_AMOUNT: 4,
  EB_RATE: 5,
  ADVANCE_PAID: 6,
  STATUS: 7,
  JOINED_DATE: 8,
  LEFT_DATE: 9,
  PREVIOUS_METER_READING: 10  // Column K
};

const EXPENSE_COLUMNS = {
  DATE: 0,
  CATEGORY: 1,
  SUBCATEGORY: 2,
  PURPOSE: 3,
  AMOUNT: 4,
  MOP: 5,
  SOP: 6
};

const EXPENSE_HEADER = [
  "Date",
  "Category",
  "Subcategory",
  "Purpose",
  "Amount",
  "MOP",
  "SOP"
];

const EXPENSE_SUBCATEGORIES = {
  Trading: [
    "Internet Bills",
    "Tradetron Subscription",
    "Monthly PF Sharing",
    "Laptop Purchase",
    "Mobile Recharge",
    "Loan EMI"
  ],
  Personal: [
    "Electricity",
    "Maintenance",
    "Vehicle Service",
    "Insurance",
    "Education",
    "Other"
  ]
};

const FO_COLUMNS = {
  DATE: 0,
  BROKER: 1,
  GROSS_NFO: 2,
  CHARGES_NFO: 3,
  NET_NFO: 4,
  GROSS_MCX: 5,
  CHARGES_MCX: 6,
  NET_MCX: 7,
  TOTAL_GROSS: 8,
  TOTAL_CHARGES: 9,
  TOTAL_NET_PNL: 10
};

const FO_HEADER = [
  "Date",
  "Broker",
  "Gross NFO",
  "Charges NFO",
  "Net NFO",
  "Gross MCX",
  "Charges MCX",
  "Net MCX",
  "Total Gross",
  "Total Charges",
  "Total Net PnL"
];

const RENT_COLUMNS = {
  DATE: 0,
  BILL_ID: 1,
  TENANT_ID: 2,
  NAME: 3,
  MONTH: 4,
  RENT_AMOUNT: 5,
  PREVIOUS_READING: 6,
  CURRENT_READING: 7,
  UNITS: 8,
  EB_AMOUNT: 9,
  TOTAL_AMOUNT: 10,
  MOP: 11,
  STATUS: 12
};

const RENT_HEADER = [
  "Date",
  "BillID",
  "Tenant ID",
  "Name",
  "Month",
  "Rent Amount",
  "Previous Reading",
  "Current Reading",
  "Units",
  "EB Amount",
  "Total Amount",
  "MOP",
  "Status"
];

const TENANT_HEADER = [
  "TenantID",
  "Name",
  "Mobile",
  "Aadhaar",
  "RentAmount",
  "EB Per Unit Rate",
  "Advance Paid",
  "Status",
  "Joined (Date)",
  "Left (Date)",
  "Previous Meter Reading"  // Column K
];

const TENANT_ARCHIVE_HEADER = [
  "TenantID",
  "Name",
  "Mobile",
  "Aadhaar",
  "RentAmount",
  "EB Per Unit Rate",
  "Advance Paid",
  "Status",
  "Joined (Date)",
  "Left (Date)",
  "Archived Date"
];

const SUMMARY_HEADER = [
  "Month",
  "Total Rent",
  "Total EB",
  "Total Collection",
  "Total Gross",
  "Total Charges",
  "Total Net PnL",
  "Total expenses"
];

/*************************************************
 LOAD HTML
*************************************************/
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Rent Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/*************************************************
 UTILITIES
*************************************************/
function getSheet(name) {
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function jsonResponse(status, message, data = null) {
  return { status, message, data };
}

function normalizeMonthValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const ymMatch = raw.match(/^(\d{4})-(\d{2})$/);
  if (ymMatch) return `${ymMatch[1]}-${ymMatch[2]}`;

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  return raw;
}

function normalizeDateValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return "";

  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const day = String(parsed.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function ensureTenantsHeader() {
  const tenantSheet = getSheet(SHEETS.TENANTS);
  if (!tenantSheet) return;

  const header = tenantSheet.getRange(1, 1, 1, TENANT_HEADER.length).getValues()[0];
  const isMatch = TENANT_HEADER.every((col, i) => String(header[i] || "").trim() === col);

  if (!isMatch) {
    tenantSheet.getRange(1, 1, 1, TENANT_HEADER.length).setValues([TENANT_HEADER]);
  }
}

function ensureTenantArchiveSheet() {
  const ss = SpreadsheetApp.getActive();
  let archiveSheet = null;

  try {
    archiveSheet = ss.getSheetByName(SHEETS.TENANT_ARCHIVE);
  } catch (e) {
    // Sheet doesn't exist, create it
  }

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(SHEETS.TENANT_ARCHIVE);
    archiveSheet.getRange(1, 1, 1, TENANT_ARCHIVE_HEADER.length).setValues([TENANT_ARCHIVE_HEADER]);
  }
}

function ensureStaffSheet() {
  const ss = SpreadsheetApp.getActive();
  let staffSheet = null;

  try {
    staffSheet = ss.getSheetByName(SHEETS.STAFF);
  } catch (e) {
    // Sheet doesn't exist, create it
  }

  if (!staffSheet) {
    staffSheet = ss.insertSheet(SHEETS.STAFF);
    staffSheet.getRange(1, 1, 1, STAFF_HEADER.length).setValues([STAFF_HEADER]);
  } else {
    // Check if Company column exists, if not insert it
    const lastCol = staffSheet.getLastColumn();
    if (lastCol > 0) {
      const header = staffSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      // Index 2 is the 3rd column (Company)
      if (String(header[2] || "").trim() !== "Company") {
        staffSheet.insertColumnBefore(3);
      }
    }
    // Always ensure headers match
    staffSheet.getRange(1, 1, 1, STAFF_HEADER.length).setValues([STAFF_HEADER]);
  }
}

function ensureSalarySheet() {
  const ss = SpreadsheetApp.getActive();
  let salarySheet = null;

  try {
    salarySheet = ss.getSheetByName(SHEETS.SALARY);
  } catch (e) {
    // Sheet doesn't exist, create it
  }

  if (!salarySheet) {
    salarySheet = ss.insertSheet(SHEETS.SALARY);
    salarySheet.getRange(1, 1, 1, SALARY_HEADER.length).setValues([SALARY_HEADER]);
  } else {
    // Since SOP is at the very end, we just overwrite headers (it naturally expands)
    salarySheet.getRange(1, 1, 1, SALARY_HEADER.length).setValues([SALARY_HEADER]);
  }
}

function ensureStaffAdvancesSheet() {
  const ss = SpreadsheetApp.getActive();
  let advanceSheet = null;

  try {
    advanceSheet = ss.getSheetByName(SHEETS.STAFF_ADVANCES);
  } catch (e) {
    // Sheet doesn't exist, create it
  }

  if (!advanceSheet) {
    advanceSheet = ss.insertSheet(SHEETS.STAFF_ADVANCES);
    advanceSheet.getRange(1, 1, 1, ADVANCE_HEADER.length).setValues([ADVANCE_HEADER]);
  } else {
    const lastCol = advanceSheet.getLastColumn();
    if (lastCol > 0) {
      const header = advanceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      // Index 6 is the 7th column (MOP)
      if (String(header[6] || "").trim() !== "MOP") {
        advanceSheet.insertColumnBefore(7);
      }
    }
    // Always ensure headers match
    advanceSheet.getRange(1, 1, 1, ADVANCE_HEADER.length).setValues([ADVANCE_HEADER]);
  }
}

function ensureRentCollectionHeader() {
  const rentSheet = getSheet(SHEETS.RENT);
  if (!rentSheet) return;

  const header = rentSheet.getRange(1, 1, 1, RENT_HEADER.length).getValues()[0];
  const isMatch = RENT_HEADER.every((col, i) => String(header[i] || "").trim() === col);

  if (!isMatch) {
    rentSheet.getRange(1, 1, 1, RENT_HEADER.length).setValues([RENT_HEADER]);
  }
}

function ensureFoIncomeHeader() {
  const foSheet = getSheet(SHEETS.FO_INCOME);
  if (!foSheet) return;

  const header = foSheet.getRange(1, 1, 1, FO_HEADER.length).getValues()[0];
  const isMatch = FO_HEADER.every((col, i) => String(header[i] || "").trim() === col);

  if (!isMatch) {
    foSheet.getRange(1, 1, 1, FO_HEADER.length).setValues([FO_HEADER]);
  }
}

function ensureExpensesHeader() {
  const expenseSheet = getSheet(SHEETS.EXPENSES);
  if (!expenseSheet) return;

  const header = expenseSheet.getRange(1, 1, 1, EXPENSE_HEADER.length).getValues()[0];
  const isMatch = EXPENSE_HEADER.every((col, i) => String(header[i] || "").trim() === col);

  if (!isMatch) {
    expenseSheet.getRange(1, 1, 1, EXPENSE_HEADER.length).setValues([EXPENSE_HEADER]);
  }
}

function ensureSummaryHeader() {
  const summarySheet = getSheet(SHEETS.SUMMARY);
  if (!summarySheet) return;

  const header = summarySheet.getRange(1, 1, 1, SUMMARY_HEADER.length).getValues()[0];
  const isMatch = SUMMARY_HEADER.every((col, i) => String(header[i] || "").trim() === col);

  if (!isMatch) {
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADER.length).setValues([SUMMARY_HEADER]);
  }
}

function generateBillId(month) {
  const monthKey = normalizeMonthValue(month);
  const yyyymm = monthKey.replace("-", "");
  const rentRows = getSheet(SHEETS.RENT).getDataRange().getValues().slice(1);

  const prefix = `BILL-${yyyymm}-`;
  let maxSeq = 0;

  rentRows.forEach(row => {
    const billId = String(row[RENT_COLUMNS.BILL_ID] || "").trim();
    if (!billId.startsWith(prefix)) return;

    const seq = Number(billId.split("-").pop());
    if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
  });

  const nextSeq = String(maxSeq + 1).padStart(3, "0");
  return `${prefix}${nextSeq}`;
}

/*************************************************
 SETUP INSTALLABLE TRIGGERS (RUN ONCE)
*************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Rent System")
    .addItem("Rebuild Monthly Summary", "rebuildMonthlySummary")
    .addItem("Setup Summary Sync Trigger", "setupSummarySyncTriggers")
    .addToUi();

  ensureTenantsHeader();
  ensureRentCollectionHeader();
  ensureFoIncomeHeader();
  ensureExpensesHeader();
  ensureSummaryHeader();
  ensureTenantArchiveSheet();
  ensureStaffSheet();
  ensureSalarySheet();
  ensureStaffAdvancesSheet();
  ensureSummarySyncTrigger();
}

function ensureSummarySyncTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const hasOnChange = triggers.some(trigger =>
      trigger.getHandlerFunction() === "onChange"
    );
    const hasOnEdit = triggers.some(trigger =>
      trigger.getHandlerFunction() === "onEdit"
    );

    if (!hasOnChange || !hasOnEdit) {
      setupSummarySyncTriggers();
    }
  } catch (e) {
    Logger.log("ensureSummarySyncTrigger error: " + e.message);
  }
}

function setupSummarySyncTriggers() {
  const ss = SpreadsheetApp.getActive();
  const existing = ScriptApp.getProjectTriggers();

  existing.forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === "onChange" || fn === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();

  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  return jsonResponse("success", "Summary sync triggers configured");
}

/*************************************************
 DASHBOARD
*************************************************/

function getMonthFromDateValue(value) {
  const normalizedDate = normalizeDateValue(value);
  if (!normalizedDate) return "";
  return normalizedDate.slice(0, 7);
}

function getDashboardMonths() {
  try {
    const months = new Set();

    const rentSheet = getSheet(SHEETS.RENT);
    if (rentSheet) {
      rentSheet.getDataRange().getValues().slice(1).forEach(row => {
        const month = normalizeMonthValue(row[RENT_COLUMNS.MONTH]);
        if (month) months.add(month);
      });
    }

    const foSheet = getSheet(SHEETS.FO_INCOME);
    if (foSheet) {
      foSheet.getDataRange().getValues().slice(1).forEach(row => {
        const month = getMonthFromDateValue(row[FO_COLUMNS.DATE]);
        if (month) months.add(month);
      });
    }

    const expenseSheet = getSheet(SHEETS.EXPENSES);
    if (expenseSheet) {
      expenseSheet.getDataRange().getValues().slice(1).forEach(row => {
        const month = getMonthFromDateValue(row[EXPENSE_COLUMNS.DATE]);
        if (month) months.add(month);
      });
    }

    return Array.from(months).sort();
  } catch (e) {
    return [];
  }
}

function dashboard(selectedMonth) {
  try {
    const tenants = getSheet(SHEETS.TENANTS).getDataRange().getValues().slice(1);
    const rent = getSheet(SHEETS.RENT).getDataRange().getValues().slice(1);
    const foSheet = getSheet(SHEETS.FO_INCOME);
    const foRows = foSheet ? foSheet.getDataRange().getValues().slice(1) : [];
    const expenseSheet = getSheet(SHEETS.EXPENSES);
    const expenseRows = expenseSheet ? expenseSheet.getDataRange().getValues().slice(1) : [];

    const monthFilter = normalizeMonthValue(selectedMonth);

    const paidRentRows = rent.filter(r => String(r[RENT_COLUMNS.STATUS]).trim() === "Paid");
    const paidRentRowsByMonth = paidRentRows.filter(r => {
      if (!monthFilter) return true;
      return normalizeMonthValue(r[RENT_COLUMNS.MONTH]) === monthFilter;
    });

    const monthlyRentReceived = paidRentRowsByMonth.reduce((sum, row) => sum + (Number(row[RENT_COLUMNS.RENT_AMOUNT]) || 0), 0);

    const foRowsByMonth = foRows.filter(row => {
      if (!monthFilter) return true;
      return getMonthFromDateValue(row[FO_COLUMNS.DATE]) === monthFilter;
    });
    const tradingMonthlyPnl = foRowsByMonth.reduce((sum, row) => sum + (Number(row[FO_COLUMNS.TOTAL_NET_PNL]) || 0), 0);

    const expenseRowsByMonth = expenseRows.filter(row => {
      if (!monthFilter) return true;
      return getMonthFromDateValue(row[EXPENSE_COLUMNS.DATE]) === monthFilter;
    });
    const totalMonthlyExpenses = expenseRowsByMonth.reduce((sum, row) => sum + (Number(row[EXPENSE_COLUMNS.AMOUNT]) || 0), 0);

    const netMonthlySavings = monthlyRentReceived + tradingMonthlyPnl - totalMonthlyExpenses;

    const occupied = tenants.filter(r => String(r[TENANT_COLUMNS.STATUS]).trim() === "Active" && String(r[TENANT_COLUMNS.NAME]).trim() !== "").length;
    const vacant = tenants.filter(r => String(r[TENANT_COLUMNS.STATUS]).trim() === "Vacant" || String(r[TENANT_COLUMNS.NAME]).trim() === "").length;

    return {
      totalHouses: tenants.length,
      occupied: occupied,
      vacant: vacant,
      pending: rent.filter(r => String(r[RENT_COLUMNS.STATUS]).trim() === "Unpaid").length,
      foNetTotal: foRows.reduce((sum, row) => sum + (Number(row[FO_COLUMNS.TOTAL_NET_PNL]) || 0), 0),
      monthlyRentReceived,
      tradingMonthlyPnl,
      totalMonthlyExpenses,
      netMonthlySavings,
      monthFilter: monthFilter || "All"
    };
  } catch (e) {
    Logger.log("Dashboard error: " + e.message);
    return {
      totalHouses: 0,
      occupied: 0,
      vacant: 0,
      pending: 0,
      foNetTotal: 0,
      monthlyRentReceived: 0,
      tradingMonthlyPnl: 0,
      totalMonthlyExpenses: 0,
      netMonthlySavings: 0,
      monthFilter: "All"
    };
  }
}

function toNumber(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function upsertFoIncomeRow(data, segment) {
  ensureFoIncomeHeader();

  const foSheet = getSheet(SHEETS.FO_INCOME);
  if (!foSheet) return jsonResponse("error", "F&O_Income sheet not found");

  const date = normalizeDateValue(data.date);
  if (!date) return jsonResponse("error", "Valid Date is required");

  const broker = String(data.broker || "").trim();
  if (!["Rmoney", "IIFL"].includes(broker)) {
    return jsonResponse("error", "Broker must be Rmoney or IIFL");
  }

  const values = foSheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < values.length; i++) {
    const rowDate = normalizeDateValue(values[i][FO_COLUMNS.DATE]);
    const rowBroker = String(values[i][FO_COLUMNS.BROKER] || "").trim();
    if (rowDate === date && rowBroker === broker) {
      targetRow = i + 1;
      break;
    }
  }

  const rowData = targetRow > 0
    ? foSheet.getRange(targetRow, 1, 1, FO_HEADER.length).getValues()[0]
    : [date, broker, "", "", "", "", "", "", "", "", ""];

  const grossNfo = toNumber(rowData[FO_COLUMNS.GROSS_NFO]);
  const chargesNfo = toNumber(rowData[FO_COLUMNS.CHARGES_NFO]);
  const netNfo = toNumber(rowData[FO_COLUMNS.NET_NFO]);
  const grossMcx = toNumber(rowData[FO_COLUMNS.GROSS_MCX]);
  const chargesMcx = toNumber(rowData[FO_COLUMNS.CHARGES_MCX]);
  const netMcx = toNumber(rowData[FO_COLUMNS.NET_MCX]);

  let nextGrossNfo = grossNfo;
  let nextChargesNfo = chargesNfo;
  let nextNetNfo = netNfo;
  let nextGrossMcx = grossMcx;
  let nextChargesMcx = chargesMcx;
  let nextNetMcx = netMcx;

  if (segment === "NFO") {
    nextGrossNfo = toNumber(data.grossNfo);
    nextNetNfo = toNumber(data.netNfo);
    nextChargesNfo = nextGrossNfo - nextNetNfo;
  }

  if (segment === "MCX") {
    nextGrossMcx = toNumber(data.grossMcx);
    nextNetMcx = toNumber(data.netMcx);
    nextChargesMcx = nextGrossMcx - nextNetMcx;
  }

  const totalGross = nextGrossNfo + nextGrossMcx;
  const totalCharges = nextChargesNfo + nextChargesMcx;
  const totalNet = nextNetNfo + nextNetMcx;

  const finalRow = [
    date,
    broker,
    nextGrossNfo,
    nextChargesNfo,
    nextNetNfo,
    nextGrossMcx,
    nextChargesMcx,
    nextNetMcx,
    totalGross,
    totalCharges,
    totalNet
  ];

  if (targetRow > 0) {
    foSheet.getRange(targetRow, 1, 1, FO_HEADER.length).setValues([finalRow]);
  } else {
    foSheet.appendRow(finalRow);
  }

  return jsonResponse("success", `${segment} details saved successfully`);
}

function submitNfoIncome(data) {
  try {
    const res = upsertFoIncomeRow(data, "NFO");
    if (res.status === "success") updateMonthlySummary(getMonthFromDateValue(data.date));
    return res;
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function submitMcxIncome(data) {
  try {
    const res = upsertFoIncomeRow(data, "MCX");
    if (res.status === "success") updateMonthlySummary(getMonthFromDateValue(data.date));
    return res;
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getExpenseSubcategories(category) {
  const key = String(category || "").trim();
  return EXPENSE_SUBCATEGORIES[key] || [];
}

function addExpenseEntry(data) {
  try {
    ensureExpensesHeader();

    const expenseSheet = getSheet(SHEETS.EXPENSES);
    if (!expenseSheet) return jsonResponse("error", "Expenses sheet not found");

    const date = normalizeDateValue(data.date);
    if (!date) return jsonResponse("error", "Valid Date is required");

    const category = String(data.category || "").trim();
    const subcategory = String(data.subcategory || "").trim();
    const purpose = String(data.purpose || "").trim();
    const amount = Number(data.amount);
    const mop = String(data.mop || "").trim();
    const sop = String(data.sop || "").trim();

    if (!Object.keys(EXPENSE_SUBCATEGORIES).includes(category)) {
      return jsonResponse("error", "Category must be Personal or Trading");
    }

    if (!EXPENSE_SUBCATEGORIES[category].includes(subcategory)) {
      return jsonResponse("error", "Invalid Subcategory for selected Category");
    }

    if (isNaN(amount)) return jsonResponse("error", "Valid Amount is required");
    if (!mop) return jsonResponse("error", "MOP is required");
    if (!sop) return jsonResponse("error", "SOP is required");

    expenseSheet.appendRow([date, category, subcategory, purpose, amount, mop, sop]);

    updateMonthlySummary(getMonthFromDateValue(date));
    return jsonResponse("success", "Expense saved successfully");
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getExpensesByCategory(category) {
  try {
    ensureExpensesHeader();

    const expenseSheet = getSheet(SHEETS.EXPENSES);
    if (!expenseSheet) return [];

    const rows = expenseSheet.getDataRange().getValues().slice(1);
    const selected = String(category || "All").trim();

    return rows
      .filter(r => {
        if (selected === "All" || !selected) return true;
        return String(r[EXPENSE_COLUMNS.CATEGORY]).trim() === selected;
      })
      .map(r => ({
        date: normalizeDateValue(r[EXPENSE_COLUMNS.DATE]),
        category: r[EXPENSE_COLUMNS.CATEGORY],
        subcategory: r[EXPENSE_COLUMNS.SUBCATEGORY],
        purpose: r[EXPENSE_COLUMNS.PURPOSE],
        amount: Number(r[EXPENSE_COLUMNS.AMOUNT]) || 0,
        mop: r[EXPENSE_COLUMNS.MOP],
        sop: r[EXPENSE_COLUMNS.SOP]
      }));
  } catch (e) {
    return [];
  }
}

/*************************************************
 TENANT MANAGEMENT
*************************************************/

function archiveTenantData(tenantRow, leftDate) {
  ensureTenantArchiveSheet();

  const archiveSheet = getSheet(SHEETS.TENANT_ARCHIVE);
  if (!archiveSheet) return false;

  const archiveRow = [
    tenantRow[TENANT_COLUMNS.TENANT_ID],
    tenantRow[TENANT_COLUMNS.NAME],
    tenantRow[TENANT_COLUMNS.MOBILE],
    tenantRow[TENANT_COLUMNS.AADHAAR],
    tenantRow[TENANT_COLUMNS.RENT_AMOUNT],
    tenantRow[TENANT_COLUMNS.EB_RATE],
    tenantRow[TENANT_COLUMNS.ADVANCE_PAID],
    tenantRow[TENANT_COLUMNS.STATUS],
    tenantRow[TENANT_COLUMNS.JOINED_DATE],
    leftDate || "",
    new Date()
  ];

  archiveSheet.appendRow(archiveRow);
  return true;
}

function addTenant(data) {
  try {
    ensureTenantsHeader();

    const sheet = getSheet(SHEETS.TENANTS);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][TENANT_COLUMNS.TENANT_ID]).trim() === String(data.tenantId).trim()) {
        const isReusable = !String(values[i][TENANT_COLUMNS.NAME] || "").trim() || String(values[i][TENANT_COLUMNS.STATUS] || "").trim() === "Vacant";
        if (!isReusable) {
          return jsonResponse("error", "Tenant ID already exists");
        }

        sheet.getRange(i + 1, 1, 1, TENANT_HEADER.length).setValues([[
          data.tenantId,
          data.name,
          data.mobile,
          data.aadhaar,
          Number(data.rentAmount) || 0,
          Number(data.ebRate) || 0,
          Number(data.advance) || 0,
          data.status || "Active",
          data.joined || new Date(),
          data.left || "",
          0
        ]]);

        return jsonResponse("success", "Tenant Added Successfully");
      }
    }

    sheet.appendRow([
      data.tenantId,
      data.name,
      data.mobile,
      data.aadhaar,
      Number(data.rentAmount) || 0,
      Number(data.ebRate) || 0,
      Number(data.advance) || 0,
      data.status || "Active",
      data.joined || new Date(),
      data.left || "",
      0
    ]);

    return jsonResponse("success", "Tenant Added Successfully");
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function updateTenant(data) {
  try {
    ensureTenantsHeader();

    const sheet = getSheet(SHEETS.TENANTS);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][TENANT_COLUMNS.TENANT_ID]).trim() === String(data.tenantId).trim()) {
        const rowNumber = i + 1;
        const nextStatus = String(data.status || "").trim();
        const leftDate = data.left ? normalizeDateValue(data.left) : "";

        if (nextStatus === "Vacant") {
          archiveTenantData(values[i], leftDate);
          sheet.getRange(rowNumber, 2, 1, 9).clearContent();
          sheet.getRange(rowNumber, 8).setValue("Vacant");
          return jsonResponse("success", "Tenant marked as Vacant and archived");
        }

        sheet.getRange(rowNumber, 1, 1, TENANT_HEADER.length).setValues([[
          data.tenantId,
          data.name,
          data.mobile,
          data.aadhaar,
          Number(data.rentAmount) || 0,
          Number(data.ebRate) || 0,
          Number(data.advance) || 0,
          nextStatus || values[i][TENANT_COLUMNS.STATUS] || "Active",
          data.joined || values[i][TENANT_COLUMNS.JOINED_DATE] || "",
          leftDate || values[i][TENANT_COLUMNS.LEFT_DATE] || "",
          values[i][TENANT_COLUMNS.PREVIOUS_METER_READING] || 0
        ]]);

        return jsonResponse("success", "Tenant Updated Successfully");
      }
    }

    return jsonResponse("error", "Tenant Not Found");
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getTenantById(tenantId) {
  try {
    const sheet = getSheet(SHEETS.TENANTS);
    const data = sheet.getDataRange().getValues().slice(1);
    const tenant = data.find(r => String(r[TENANT_COLUMNS.TENANT_ID]).trim() === String(tenantId).trim());

    if (!tenant) return null;

    return {
      tenantId: tenant[TENANT_COLUMNS.TENANT_ID],
      name: tenant[TENANT_COLUMNS.NAME],
      mobile: tenant[TENANT_COLUMNS.MOBILE],
      aadhaar: tenant[TENANT_COLUMNS.AADHAAR],
      rentAmount: Number(tenant[TENANT_COLUMNS.RENT_AMOUNT]) || 0,
      ebRate: Number(tenant[TENANT_COLUMNS.EB_RATE]) || 0,
      advance: Number(tenant[TENANT_COLUMNS.ADVANCE_PAID]) || 0,
      status: tenant[TENANT_COLUMNS.STATUS],
      joined: normalizeDateValue(tenant[TENANT_COLUMNS.JOINED_DATE]),
      left: normalizeDateValue(tenant[TENANT_COLUMNS.LEFT_DATE]),
      previousMeterReading: Number(tenant[TENANT_COLUMNS.PREVIOUS_METER_READING]) || 0
    };
  } catch (e) {
    return null;
  }
}

function getAllTenantsDropdown() {
  const sheet = getSheet(SHEETS.TENANTS);
  const data = sheet.getDataRange().getValues().slice(1);
  return data
    .filter(r => String(r[TENANT_COLUMNS.TENANT_ID]).trim())
    .map(r => ({
      id: r[TENANT_COLUMNS.TENANT_ID],
      name: String(r[TENANT_COLUMNS.NAME] || "").trim() || "(Vacant - Available)"
    }));
}

function getArchivedTenants() {
  try {
    ensureTenantArchiveSheet();
    const archiveSheet = getSheet(SHEETS.TENANT_ARCHIVE);
    if (!archiveSheet) return [];
    return archiveSheet.getDataRange().getValues().slice(1);
  } catch (e) {
    Logger.log("Error getting archived tenants: " + e.message);
    return [];
  }
}

/*************************************************
 ELECTRICITY + RENT
*************************************************/

function getPreviousMeterReadingFromTenant(tenantId) {
  try {
    const sheet = getSheet(SHEETS.TENANTS);
    const data = sheet.getDataRange().getValues().slice(1);
    const tenant = data.find(r => String(r[TENANT_COLUMNS.TENANT_ID]).trim() === String(tenantId).trim());

    if (!tenant) return 0;

    const reading = Number(tenant[TENANT_COLUMNS.PREVIOUS_METER_READING]) || 0;
    return reading;
  } catch (e) {
    Logger.log("Error getting previous meter reading: " + e.message);
    return 0;
  }
}

function updatePreviousMeterReadingInTenant(tenantId, newReading) {
  try {
    const sheet = getSheet(SHEETS.TENANTS);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][TENANT_COLUMNS.TENANT_ID]).trim() === String(tenantId).trim()) {
        sheet.getRange(i + 1, TENANT_COLUMNS.PREVIOUS_METER_READING + 1).setValue(Number(newReading) || 0);
        return true;
      }
    }

    return false;
  } catch (e) {
    Logger.log("Error updating previous meter reading: " + e.message);
    return false;
  }
}

function recordMeter(data) {
  try {
    ensureRentCollectionHeader();

    const rentSheet = getSheet(SHEETS.RENT);
    const tenants = getSheet(SHEETS.TENANTS).getDataRange().getValues().slice(1);

    const tenant = tenants.find(r => r[TENANT_COLUMNS.TENANT_ID] === data.tenantId);
    if (!tenant) return jsonResponse("error", "Tenant Not Found");

    const normalizedMonth = normalizeMonthValue(data.month);
    if (!normalizedMonth) return jsonResponse("error", "Month is required");

    let previous;
    if (data.previous !== "" && data.previous !== null && data.previous !== undefined) {
      previous = Number(data.previous);
    } else {
      previous = getPreviousMeterReadingFromTenant(data.tenantId);
    }

    const current = Number(data.current);
    if (isNaN(current)) return jsonResponse("error", "Invalid Current Reading");
    if (isNaN(previous)) return jsonResponse("error", "Invalid Previous Reading");
    if (current < previous) return jsonResponse("error", "Current reading cannot be less than previous");

    const units = current - previous;
    const rate = Number(tenant[TENANT_COLUMNS.EB_RATE]) || 0;
    const rentAmount = Number(tenant[TENANT_COLUMNS.RENT_AMOUNT]) || 0;
    const ebAmount = units * rate;
    const totalAmount = rentAmount + ebAmount;
    const billId = generateBillId(normalizedMonth);

    rentSheet.appendRow([
      new Date(),
      billId,
      data.tenantId,
      tenant[TENANT_COLUMNS.NAME],
      normalizedMonth,
      rentAmount,
      previous,
      current,
      units,
      ebAmount,
      totalAmount,
      "",
      "Unpaid"
    ]);

    updatePreviousMeterReadingInTenant(data.tenantId, current);

    updateMonthlySummary(normalizedMonth);
    return jsonResponse("success", "Meter Recorded Successfully", { billId });
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getTenantDetails(tenantId) {
  const sheet = getSheet(SHEETS.TENANTS);
  const data = sheet.getDataRange().getValues().slice(1);

  const tenant = data.find(r => r[TENANT_COLUMNS.TENANT_ID] === tenantId && r[TENANT_COLUMNS.STATUS] === "Active");
  if (!tenant) return null;

  return {
    name: tenant[TENANT_COLUMNS.NAME],
    rent: tenant[TENANT_COLUMNS.RENT_AMOUNT],
    rate: tenant[TENANT_COLUMNS.EB_RATE],
    previousReading: Number(tenant[TENANT_COLUMNS.PREVIOUS_METER_READING]) || 0
  };
}

function getActiveTenantsDropdown() {
  const sheet = getSheet(SHEETS.TENANTS);
  const data = sheet.getDataRange().getValues().slice(1);

  return data.filter(r => r[TENANT_COLUMNS.STATUS] === "Active" && String(r[TENANT_COLUMNS.NAME]).trim() !== "").map(r => ({
    id: r[TENANT_COLUMNS.TENANT_ID],
    name: r[TENANT_COLUMNS.NAME]
  }));
}

/*************************************************
 PAYMENTS - FIXED WITH DATE FIELD
*************************************************/
function markPaid(data) {
  try {
    if (!data.billId || !data.month) {
      return jsonResponse("error", "Bill ID and Month are required");
    }

    // NEW: Validate payment date
    const paymentDate = data.paymentDate ? normalizeDateValue(data.paymentDate) : "";
    if (!paymentDate) {
      return jsonResponse("error", "Payment Date is required");
    }

    const rentSheet = getSheet(SHEETS.RENT);
    const rentData = rentSheet.getDataRange().getValues();

    let rentUpdated = false;

    for (let i = 1; i < rentData.length; i++) {
      const rentBillId = String(rentData[i][RENT_COLUMNS.BILL_ID]).trim();
      const rentStatus = String(rentData[i][RENT_COLUMNS.STATUS]).trim();

      if (rentBillId === String(data.billId).trim() && rentStatus === "Unpaid") {
        rentSheet.getRange(i + 1, RENT_COLUMNS.DATE + 1).setValue(paymentDate);  // NEW: Update payment date
        rentSheet.getRange(i + 1, RENT_COLUMNS.MOP + 1).setValue(data.paymentMode);
        rentSheet.getRange(i + 1, RENT_COLUMNS.STATUS + 1).setValue("Paid");
        rentUpdated = true;
        break;
      }
    }

    if (!rentUpdated) {
      return jsonResponse("error", "Bill not found or already paid");
    }

    updateMonthlySummary(data.month);
    return jsonResponse("success", "✅ Payment Successful! Bill marked as Paid on " + paymentDate);
  } catch (e) {
    return jsonResponse("error", "Error: " + e.message);
  }
}

/*************************************************
 MONTHLY SUMMARY
*************************************************/
function calculateSummaryForMonth(month) {
  const normalizedMonth = normalizeMonthValue(month);
  if (!normalizedMonth) return null;

  const rentRows = getSheet(SHEETS.RENT).getDataRange().getValues().slice(1);
  const foRows = getSheet(SHEETS.FO_INCOME).getDataRange().getValues().slice(1);
  const expenseRows = getSheet(SHEETS.EXPENSES).getDataRange().getValues().slice(1);

  const paidRentRows = rentRows.filter(r =>
    normalizeMonthValue(r[RENT_COLUMNS.MONTH]) === normalizedMonth &&
    String(r[RENT_COLUMNS.STATUS]).trim() === "Paid"
  );

  const totalRent = paidRentRows.reduce((s, r) => s + (Number(r[RENT_COLUMNS.RENT_AMOUNT]) || 0), 0);
  const totalEB = paidRentRows.reduce((s, r) => s + (Number(r[RENT_COLUMNS.EB_AMOUNT]) || 0), 0);
  const totalCollection = paidRentRows.reduce((s, r) => s + (Number(r[RENT_COLUMNS.TOTAL_AMOUNT]) || 0), 0);

  const foMonthRows = foRows.filter(r => getMonthFromDateValue(r[FO_COLUMNS.DATE]) === normalizedMonth);
  const totalGross = foMonthRows.reduce((s, r) => s + (Number(r[FO_COLUMNS.TOTAL_GROSS]) || 0), 0);
  const totalCharges = foMonthRows.reduce((s, r) => s + (Number(r[FO_COLUMNS.TOTAL_CHARGES]) || 0), 0);
  const totalNetPnl = foMonthRows.reduce((s, r) => s + (Number(r[FO_COLUMNS.TOTAL_NET_PNL]) || 0), 0);

  const expenseMonthRows = expenseRows.filter(r => getMonthFromDateValue(r[EXPENSE_COLUMNS.DATE]) === normalizedMonth);
  const totalExpenses = expenseMonthRows.reduce((s, r) => s + (Number(r[EXPENSE_COLUMNS.AMOUNT]) || 0), 0);

  return {
    month: normalizedMonth,
    totalRent,
    totalEB,
    totalCollection,
    totalGross,
    totalCharges,
    totalNetPnl,
    totalExpenses
  };
}

function updateMonthlySummary(month) {
  try {
    ensureSummaryHeader();

    const summary = getSheet(SHEETS.SUMMARY);
    const calculated = calculateSummaryForMonth(month);
    if (!calculated) return;

    const data = summary.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      const summaryMonth = normalizeMonthValue(data[i][0]);
      if (summaryMonth === calculated.month) {
        summary.getRange(i + 1, 1, 1, SUMMARY_HEADER.length).setValues([[
          calculated.month,
          calculated.totalRent,
          calculated.totalEB,
          calculated.totalCollection,
          calculated.totalGross,
          calculated.totalCharges,
          calculated.totalNetPnl,
          calculated.totalExpenses
        ]]);
        found = true;
        break;
      }
    }

    if (!found) {
      summary.appendRow([
        calculated.month,
        calculated.totalRent,
        calculated.totalEB,
        calculated.totalCollection,
        calculated.totalGross,
        calculated.totalCharges,
        calculated.totalNetPnl,
        calculated.totalExpenses
      ]);
    }
  } catch (e) {
    Logger.log("Error updating summary: " + e.message);
  }
}

/*************************************************
 FORMAT MONTH HELPER
*************************************************/
function formatMonthDisplay(monthRaw) {
  let month = String(monthRaw).trim();
  if (month.includes("T") || !isNaN(Date.parse(month))) {
    try {
      const date = new Date(month);
      month = date.toLocaleString("default", { month: "long", year: "numeric" });
    } catch (e) {
      // keep original
    }
  }
  return month;
}

/*************************************************
 GET UNPAID BILLS DROPDOWN
*************************************************/
function getUnpaidTenantsDropdown() {
  try {
    const sheet = getSheet(SHEETS.RENT);
    const data = sheet.getDataRange().getValues();
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const billId = String(row[RENT_COLUMNS.BILL_ID]).trim();
      const status = String(row[RENT_COLUMNS.STATUS]).trim();
      const tenantId = String(row[RENT_COLUMNS.TENANT_ID]).trim();
      const name = String(row[RENT_COLUMNS.NAME]).trim();
      const month = formatMonthDisplay(row[RENT_COLUMNS.MONTH]);

      if (status === "Unpaid" && billId && billId !== "undefined") {
        result.push({ billId, tenantId, name, month });
      }
    }

    return result;
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    return [];
  }
}

/*************************************************
 FETCH UNPAID BILL DETAILS
*************************************************/
function getUnpaidBillByBillId(billId) {
  try {
    const rentSheet = getSheet(SHEETS.RENT);
    const rentData = rentSheet.getDataRange().getValues().slice(1);

    for (const r of rentData) {
      const currentBillId = String(r[RENT_COLUMNS.BILL_ID]).trim();
      const currentStatus = String(r[RENT_COLUMNS.STATUS]).trim();

      if (currentBillId === String(billId).trim() && currentStatus === "Unpaid") {
        const month = formatMonthDisplay(r[RENT_COLUMNS.MONTH]);

        return jsonResponse("success", "Bill details fetched", {
          tenantId: r[RENT_COLUMNS.TENANT_ID],
          name: r[RENT_COLUMNS.NAME],
          month,
          rent: Number(r[RENT_COLUMNS.RENT_AMOUNT]),
          ebAmount: Number(r[RENT_COLUMNS.EB_AMOUNT]),
          total: Number(r[RENT_COLUMNS.TOTAL_AMOUNT]),
          billId: currentBillId,
          status: currentStatus
        });
      }
    }

    return jsonResponse("error", "Unpaid bill not found");
  } catch (e) {
    return jsonResponse("error", "Error: " + e.message);
  }
}

/*************************************************
 GET TENANT MOBILE NUMBER
*************************************************/
function getTenantMobile(tenantId) {
  try {
    const sheet = getSheet(SHEETS.TENANTS);
    const data = sheet.getDataRange().getValues().slice(1);
    const tenant = data.find(r => r[TENANT_COLUMNS.TENANT_ID] === tenantId);
    if (!tenant) return null;

    return { name: tenant[TENANT_COLUMNS.NAME], mobile: String(tenant[TENANT_COLUMNS.MOBILE]).trim() };
  } catch (e) {
    return null;
  }
}

/*************************************************
 GET METER READING BY BILL ID
*************************************************/
function getMeterReadingByBillId(billId) {
  try {
    const rentSheet = getSheet(SHEETS.RENT);
    const data = rentSheet.getDataRange().getValues().slice(1);

    for (const r of data) {
      const currentBillId = String(r[RENT_COLUMNS.BILL_ID]).trim();
      if (currentBillId === String(billId).trim()) {
        const tenantDetails = getTenantDetails(String(r[RENT_COLUMNS.TENANT_ID]).trim());
        return {
          previousReading: Number(r[RENT_COLUMNS.PREVIOUS_READING]),
          currentReading: Number(r[RENT_COLUMNS.CURRENT_READING]),
          units: Number(r[RENT_COLUMNS.UNITS]),
          rate: tenantDetails ? Number(tenantDetails.rate) : "N/A"
        };
      }
    }

    return null;
  } catch (e) {
    Logger.log("Error: " + e.message);
    return null;
  }
}

/*************************************************
 REBUILD MONTHLY SUMMARY (FULL RECALC)
*************************************************/
function rebuildMonthlySummary() {
  try {
    ensureSummaryHeader();

    const rentRows = getSheet(SHEETS.RENT).getDataRange().getValues().slice(1);
    const foRows = getSheet(SHEETS.FO_INCOME).getDataRange().getValues().slice(1);
    const expenseRows = getSheet(SHEETS.EXPENSES).getDataRange().getValues().slice(1);

    const months = new Set();

    rentRows.forEach(row => {
      const month = normalizeMonthValue(row[RENT_COLUMNS.MONTH]);
      if (month) months.add(month);
    });

    foRows.forEach(row => {
      const month = getMonthFromDateValue(row[FO_COLUMNS.DATE]);
      if (month) months.add(month);
    });

    expenseRows.forEach(row => {
      const month = getMonthFromDateValue(row[EXPENSE_COLUMNS.DATE]);
      if (month) months.add(month);
    });

    const sortedMonths = Array.from(months).sort();
    const rows = sortedMonths
      .map(month => calculateSummaryForMonth(month))
      .filter(Boolean)
      .map(summary => [
        summary.month,
        summary.totalRent,
        summary.totalEB,
        summary.totalCollection,
        summary.totalGross,
        summary.totalCharges,
        summary.totalNetPnl,
        summary.totalExpenses
      ]);

    const summarySheet = getSheet(SHEETS.SUMMARY);
    summarySheet.clearContents();
    summarySheet.getRange(1, 1, 1, SUMMARY_HEADER.length).setValues([SUMMARY_HEADER]);

    if (rows.length > 0) {
      summarySheet.getRange(2, 1, rows.length, SUMMARY_HEADER.length).setValues(rows);
    }

    return jsonResponse("success", "Monthly summary rebuilt", { months: rows.length });
  } catch (e) {
    return jsonResponse("error", "Error rebuilding summary: " + e.message);
  }
}

/*************************************************
 AUTO-SYNC SUMMARY ON SHEET STRUCTURE CHANGES
*************************************************/
function onChange(e) {
  try {
    if (!e || !e.changeType) return;

    const shouldRebuild = [
      "REMOVE_ROW",
      "INSERT_ROW",
      "EDIT",
      "REMOVE_COLUMN",
      "INSERT_COLUMN",
      "OTHER"
    ].includes(String(e.changeType));

    if (shouldRebuild) {
      rebuildMonthlySummary();
    }
  } catch (err) {
    Logger.log("onChange summary sync error: " + err.message);
  }
}

/*************************************************
 AUTO-SYNC SUMMARY ON CELL EDITS IN RELEVANT SHEETS
*************************************************/
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const watchedSheets = [SHEETS.RENT, SHEETS.FO_INCOME, SHEETS.EXPENSES];
    if (!watchedSheets.includes(sheetName)) return;

    if (e.range.getRow() > 1) {
      rebuildMonthlySummary();
    }
  } catch (err) {
    Logger.log("onEdit summary sync error: " + err.message);
  }
}

/*************************************************
 CLOSE FINANCIAL YEAR (ARCHIVE DATA)
*************************************************/
function closeFinancialYear() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    // Assuming FY ends in March, the year label might just be the current year.
    // E.g., "FY_2024" or a timestamp if multiple closes happen.
    const suffix = "_FY_" + now.getFullYear() + "_" + now.getTime();

    const sheetsToArchive = [SHEETS.RENT, SHEETS.SUMMARY, SHEETS.FO_INCOME, SHEETS.EXPENSES, SHEETS.SALARY, SHEETS.STAFF_ADVANCES];

    // Step 1: Duplicate existing sheets and rename them
    sheetsToArchive.forEach(sheetName => {
      const originalSheet = ss.getSheetByName(sheetName);
      if (originalSheet) {
        // Create an exact copy (includes data, formatting, etc)
        const copiedSheet = originalSheet.copyTo(ss);
        copiedSheet.setName(sheetName + suffix);
        // Optionally hide it so the tab bar doesn't get cluttered
        copiedSheet.hideSheet();
      }
    });

    // Step 2: Clear original sheets (leaving headers intact)
    sheetsToArchive.forEach(sheetName => {
      const originalSheet = ss.getSheetByName(sheetName);
      if (originalSheet) {
        const lastRow = originalSheet.getLastRow();
        const lastCol = originalSheet.getLastColumn();
        if (lastRow > 1) {
          // Clear all rows from row 2 down
          originalSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        }
      }
    });

    // Step 3: Rebuild the empty summary (optional, but ensures fresh state)
    rebuildMonthlySummary();

    return jsonResponse("success", "Financial Year Closed Successfully. Data archived to new sheets.");
  } catch (e) {
    Logger.log("Error in closeFinancialYear: " + e.message);
    return jsonResponse("error", "Error closing Financial Year: " + e.message);
  }
}
/*************************************************
 COMPREHENSIVE REPORTS
*************************************************/
function getComprehensiveReport() {
  try {
    const reportData = [];

    // 1. Rent Collection
    const rentSheet = getSheet(SHEETS.RENT);
    if (rentSheet) {
      const rentRows = rentSheet.getDataRange().getValues().slice(1);
      rentRows.forEach(row => {
        if (String(row[RENT_COLUMNS.STATUS]).trim() === 'Paid') {
          reportData.push({
            date: normalizeDateValue(row[RENT_COLUMNS.DATE]),
            type: 'Income',
            category: 'Rent',
            entity: row[RENT_COLUMNS.NAME],
            amount: Number(row[RENT_COLUMNS.TOTAL_AMOUNT]) || 0,
            mop: row[RENT_COLUMNS.MOP] || "",
            sop: "", // Rent typically goes to a designated account, but no SOP column exists in RENT_COLUMNS currently
            details: 'Bill ID: ' + (row[RENT_COLUMNS.BILL_ID] || "")
          });
        }
      });
    }

    // 2. F&O Income
    const foSheet = getSheet(SHEETS.FO_INCOME);
    if (foSheet) {
      const foRows = foSheet.getDataRange().getValues().slice(1);
      foRows.forEach(row => {
        reportData.push({
          date: normalizeDateValue(row[FO_COLUMNS.DATE]),
          type: 'Income',
          category: 'F&O Trading',
          entity: row[FO_COLUMNS.BROKER],
          amount: Number(row[FO_COLUMNS.TOTAL_NET_PNL]) || 0,
          mop: "Bank Transfer", // F&O is usually Bank Transfer
          sop: "",
          details: 'Gross: ' + (Number(row[FO_COLUMNS.TOTAL_GROSS]) || 0) + ', Charges: ' + (Number(row[FO_COLUMNS.TOTAL_CHARGES]) || 0)
        });
      });
    }

    // 3. Expenses
    const expenseSheet = getSheet(SHEETS.EXPENSES);
    if (expenseSheet) {
      const expenseRows = expenseSheet.getDataRange().getValues().slice(1);
      expenseRows.forEach(row => {
        reportData.push({
          date: normalizeDateValue(row[EXPENSE_COLUMNS.DATE]),
          type: 'Expense',
          category: 'Expense - ' + row[EXPENSE_COLUMNS.CATEGORY],
          entity: row[EXPENSE_COLUMNS.SUBCATEGORY],
          amount: Number(row[EXPENSE_COLUMNS.AMOUNT]) || 0,
          mop: row[EXPENSE_COLUMNS.MOP] || "",
          sop: row[EXPENSE_COLUMNS.SOP] || "",
          details: row[EXPENSE_COLUMNS.PURPOSE] || ""
        });
      });
    }

    // 4. Salary Payouts
    const salarySheet = getSheet(SHEETS.SALARY);
    if (salarySheet) {
      const salaryRows = salarySheet.getDataRange().getValues().slice(1);
      salaryRows.forEach(row => {
        reportData.push({
          date: normalizeDateValue(row[SALARY_COLUMNS.DATE]),
          type: 'Expense',
          category: 'Salary',
          entity: row[SALARY_COLUMNS.NAME],
          amount: Number(row[SALARY_COLUMNS.NET_PAID]) || 0,
          mop: row[SALARY_COLUMNS.MOP] || "",
          sop: row[SALARY_COLUMNS.SOP] || "",
          details: 'Month: ' + (row[SALARY_COLUMNS.MONTH] || "")
        });
      });
    }

    // 5. Staff Advances
    const advanceSheet = getSheet(SHEETS.STAFF_ADVANCES);
    if (advanceSheet) {
      const advanceRows = advanceSheet.getDataRange().getValues().slice(1);
      advanceRows.forEach(row => {
        if (String(row[ADVANCE_COLUMNS.TYPE]).trim() === 'Given') {
          reportData.push({
            date: normalizeDateValue(row[ADVANCE_COLUMNS.DATE]),
            type: 'Expense',
            category: 'Staff Advance',
            entity: row[ADVANCE_COLUMNS.NAME],
            amount: Number(row[ADVANCE_COLUMNS.AMOUNT]) || 0,
            mop: row[ADVANCE_COLUMNS.MOP] || "",
            sop: "", // SOP is not in Staff Advances yet, though maybe inferred
            details: row[ADVANCE_COLUMNS.DESCRIPTION] || ""
          });
        }
      });
    }

    // Sort by date descending (newest first)
    reportData.sort((a, b) => new Date(b.date) - new Date(a.date));

    return jsonResponse("success", "Comprehensive report generated", reportData);
  } catch (e) {
    Logger.log("Error in getComprehensiveReport: " + e.message);
    return jsonResponse("error", "Failed to generate report: " + e.message);
  }
}

/*************************************************
 STAFF MANAGEMENT
*************************************************/
function addStaff(data) {
  try {
    ensureStaffSheet();

    const sheet = getSheet(SHEETS.STAFF);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][STAFF_COLUMNS.STAFF_ID]).trim() === String(data.staffId).trim()) {
        const isReusable = !String(values[i][STAFF_COLUMNS.NAME] || "").trim() || String(values[i][STAFF_COLUMNS.STATUS] || "").trim() === "Inactive";
        if (!isReusable) {
          return jsonResponse("error", "Staff ID already exists");
        }

        sheet.getRange(i + 1, 1, 1, STAFF_HEADER.length).setValues([[
          data.staffId,
          data.name,
          data.company || "",
          Number(data.salary) || 0,
          data.bankName || "",
          data.accountName || "",
          data.accountNumber || "",
          Number(data.advanceBalance) || 0,
          data.status || "Active",
          data.joined || new Date(),
          data.left || ""
        ]]);

        return jsonResponse("success", "Staff Added Successfully");
      }
    }

    sheet.appendRow([
      data.staffId,
      data.name,
      data.company || "",
      Number(data.salary) || 0,
      data.bankName || "",
      data.accountName || "",
      data.accountNumber || "",
      Number(data.advanceBalance) || 0,
      data.status || "Active",
      data.joined || new Date(),
      data.left || ""
    ]);

    return jsonResponse("success", "Staff Added Successfully");
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function updateStaff(data) {
  try {
    ensureStaffSheet();

    const sheet = getSheet(SHEETS.STAFF);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][STAFF_COLUMNS.STAFF_ID]).trim() === String(data.staffId).trim()) {
        const rowNumber = i + 1;
        const leftDate = data.left ? normalizeDateValue(data.left) : "";

        sheet.getRange(rowNumber, 1, 1, STAFF_HEADER.length).setValues([[
          data.staffId,
          data.name,
          data.company || values[i][STAFF_COLUMNS.COMPANY] || "",
          Number(data.salary) || 0,
          data.bankName || "",
          data.accountName || "",
          data.accountNumber || "",
          Number(data.advanceBalance) || 0,
          data.status || values[i][STAFF_COLUMNS.STATUS] || "Active",
          data.joined || values[i][STAFF_COLUMNS.JOINED_DATE] || "",
          leftDate || values[i][STAFF_COLUMNS.LEFT_DATE] || ""
        ]]);

        return jsonResponse("success", "Staff Updated Successfully");
      }
    }

    return jsonResponse("error", "Staff Not Found");
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getStaffById(staffId) {
  try {
    const sheet = getSheet(SHEETS.STAFF);
    const data = sheet.getDataRange().getValues().slice(1);
    const staff = data.find(r => String(r[STAFF_COLUMNS.STAFF_ID]).trim() === String(staffId).trim());

    if (!staff) return null;

    return {
      staffId: staff[STAFF_COLUMNS.STAFF_ID],
      name: staff[STAFF_COLUMNS.NAME],
      company: staff[STAFF_COLUMNS.COMPANY],
      salary: Number(staff[STAFF_COLUMNS.SALARY]) || 0,
      bankName: staff[STAFF_COLUMNS.BANK_NAME],
      accountName: staff[STAFF_COLUMNS.ACCOUNT_NAME],
      accountNumber: staff[STAFF_COLUMNS.ACCOUNT_NUMBER],
      advanceBalance: Number(staff[STAFF_COLUMNS.ADVANCE_BALANCE]) || 0,
      status: staff[STAFF_COLUMNS.STATUS],
      joined: normalizeDateValue(staff[STAFF_COLUMNS.JOINED_DATE]),
      left: normalizeDateValue(staff[STAFF_COLUMNS.LEFT_DATE])
    };
  } catch (e) {
    return null;
  }
}

function getActiveStaffDropdown() {
  try {
    const sheet = getSheet(SHEETS.STAFF);
    const data = sheet.getDataRange().getValues().slice(1);
    return data
      .filter(r => r[STAFF_COLUMNS.STATUS] === "Active" && String(r[STAFF_COLUMNS.NAME]).trim() !== "")
      .map(r => ({
        id: r[STAFF_COLUMNS.STAFF_ID],
        name: r[STAFF_COLUMNS.NAME],
        company: String(r[STAFF_COLUMNS.COMPANY] || '').trim(),
        salary: Number(r[STAFF_COLUMNS.SALARY]) || 0,
        advanceBal: Number(r[STAFF_COLUMNS.ADVANCE_BALANCE]) || 0,
        bank: String(r[STAFF_COLUMNS.BANK_NAME] || '').trim(),
        accNo: String(r[STAFF_COLUMNS.ACCOUNT_NUMBER] || '').trim()
      }));
  } catch (e) {
    return [];
  }
}

function giveAdvance(data) {
  try {
    ensureStaffAdvancesSheet();
    ensureStaffSheet();

    const date = data.date ? normalizeDateValue(data.date) : new Date();
    const amount = Number(data.amount);

    if (!data.staffId || isNaN(amount) || amount <= 0) {
      return jsonResponse("error", "Valid Staff ID and Amount are required");
    }

    const staffSheet = getSheet(SHEETS.STAFF);
    const staffData = staffSheet.getDataRange().getValues();
    let staffRowIndex = -1;
    let staffName = "";
    let currentAdvanceBalance = 0;

    for (let i = 1; i < staffData.length; i++) {
      if (String(staffData[i][STAFF_COLUMNS.STAFF_ID]).trim() === String(data.staffId).trim()) {
        staffRowIndex = i + 1;
        staffName = staffData[i][STAFF_COLUMNS.NAME];
        currentAdvanceBalance = Number(staffData[i][STAFF_COLUMNS.ADVANCE_BALANCE]) || 0;
        break;
      }
    }

    if (staffRowIndex === -1) {
      return jsonResponse("error", "Staff member not found");
    }

    const advanceSheet = getSheet(SHEETS.STAFF_ADVANCES);
    const advanceRows = advanceSheet.getDataRange().getValues().slice(1);

    // Generate Transaction ID
    const yyyymm = normalizeDateValue(date).replace("-", "").slice(0, 6);
    const prefix = `ADV-${yyyymm}-`;
    let maxSeq = 0;

    advanceRows.forEach(row => {
      const transId = String(row[ADVANCE_COLUMNS.TRANSACTION_ID] || "").trim();
      if (!transId.startsWith(prefix)) return;

      const seq = Number(transId.split("-").pop());
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    });

    const nextSeq = String(maxSeq + 1).padStart(3, "0");
    const transactionId = `${prefix}${nextSeq}`;

    // 1. Append to ledger
    advanceSheet.appendRow([
      date,
      transactionId,
      data.staffId,
      staffName,
      "Given",
      amount,
      data.mop || "",
      data.description || ""
    ]);

    // 2. Update staff balance
    const newAdvanceBalance = currentAdvanceBalance + amount;
    staffSheet.getRange(staffRowIndex, STAFF_COLUMNS.ADVANCE_BALANCE + 1).setValue(newAdvanceBalance);

    return jsonResponse("success", "Advance recorded successfully", { transactionId, newAdvanceBalance });
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}

function getAdvanceHistory(staffId) {
  try {
    ensureStaffAdvancesSheet();
    const sheet = getSheet(SHEETS.STAFF_ADVANCES);
    const data = sheet.getDataRange().getValues().slice(1);

    const targetId = String(staffId).trim();

    const history = data
      .filter(r => String(r[ADVANCE_COLUMNS.STAFF_ID]).trim() === targetId)
      .map(r => ({
        date: normalizeDateValue(r[ADVANCE_COLUMNS.DATE]),
        transactionId: r[ADVANCE_COLUMNS.TRANSACTION_ID],
        type: r[ADVANCE_COLUMNS.TYPE],
        amount: Number(r[ADVANCE_COLUMNS.AMOUNT]) || 0,
        mop: r[ADVANCE_COLUMNS.MOP] || "",
        description: r[ADVANCE_COLUMNS.DESCRIPTION]
      }));

    // Sort descending by date
    history.sort((a, b) => new Date(b.date) - new Date(a.date));

    return jsonResponse("success", "History fetched", history);
  } catch (e) {
    return jsonResponse("error", "Error fetching history: " + e.message);
  }
}

/*************************************************
 PAYROLL MANAGEMENT
*************************************************/
function processSalaryPayment(data) {
  try {
    ensureSalarySheet();
    ensureStaffSheet();

    if (!data.staffId || !data.month) {
      return jsonResponse("error", "Staff ID and Month are required");
    }

    const paymentDate = data.paymentDate ? normalizeDateValue(data.paymentDate) : new Date();
    const normalizedMonth = normalizeMonthValue(data.month);

    const staffSheet = getSheet(SHEETS.STAFF);
    const staffData = staffSheet.getDataRange().getValues();
    let staffRowIndex = -1;
    let staffDetails = null;

    for (let i = 1; i < staffData.length; i++) {
      if (String(staffData[i][STAFF_COLUMNS.STAFF_ID]).trim() === String(data.staffId).trim()) {
        staffRowIndex = i + 1;
        staffDetails = staffData[i];
        break;
      }
    }

    if (staffRowIndex === -1) {
      return jsonResponse("error", "Staff member not found");
    }

    const baseSalary = Number(data.baseSalary) || Number(staffDetails[STAFF_COLUMNS.SALARY]) || 0;
    const daysWorked = data.daysWorked !== undefined && data.daysWorked !== "" ? Number(data.daysWorked) : "";
    const earnedSalary = data.earnedSalary !== undefined && data.earnedSalary !== "" ? Number(data.earnedSalary) : baseSalary;
    const bonus = Number(data.bonus) || 0;
    const advanceDeducted = Number(data.advanceDeducted) || 0;

    // In actual payroll calculation, Net Paid might include prorated base salary and bonus,
    // but we respect what the frontend computed if provided, or fallback to simple math.
    const netPaid = Number(data.netPaid) || (earnedSalary + bonus - advanceDeducted);
    const currentAdvanceBalance = Number(staffDetails[STAFF_COLUMNS.ADVANCE_BALANCE]) || 0;

    // Validate advance deduction
    if (advanceDeducted > currentAdvanceBalance) {
      return jsonResponse("error", "Advance deducted cannot exceed the current Advance Balance");
    }

    // Generate Transaction ID
    const yyyymm = normalizedMonth.replace("-", "");
    const salarySheet = getSheet(SHEETS.SALARY);
    const salaryRows = salarySheet.getDataRange().getValues().slice(1);

    const prefix = `SAL-${yyyymm}-`;
    let maxSeq = 0;

    salaryRows.forEach(row => {
      const transId = String(row[SALARY_COLUMNS.TRANSACTION_ID] || "").trim();
      if (!transId.startsWith(prefix)) return;

      const seq = Number(transId.split("-").pop());
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    });

    const nextSeq = String(maxSeq + 1).padStart(3, "0");
    const transactionId = `${prefix}${nextSeq}`;

    // 1. Record the payout
    salarySheet.appendRow([
      paymentDate,
      transactionId,
      data.staffId,
      staffDetails[STAFF_COLUMNS.NAME],
      normalizedMonth,
      baseSalary,
      daysWorked,
      earnedSalary,
      bonus,
      advanceDeducted,
      netPaid,
      data.mop || "Bank Transfer",
      data.sop || ""
    ]);

    // 2. Ledger & Balance updates for advance deductions
    let newAdvanceBalance = currentAdvanceBalance;
    if (advanceDeducted > 0) {
      newAdvanceBalance = currentAdvanceBalance - advanceDeducted;
      staffSheet.getRange(staffRowIndex, STAFF_COLUMNS.ADVANCE_BALANCE + 1).setValue(newAdvanceBalance);

      // Also log deduction in Staff_Advances sheet
      ensureStaffAdvancesSheet();
      const advanceSheet = getSheet(SHEETS.STAFF_ADVANCES);
      advanceSheet.appendRow([
        paymentDate,
        transactionId,
        data.staffId,
        staffDetails[STAFF_COLUMNS.NAME],
        "Deducted",
        advanceDeducted,
        data.mop || "Bank Transfer",
        `Salary Deduction (${normalizedMonth})`
      ]);
    }

    return jsonResponse("success", "Salary payment processed successfully", { transactionId });
  } catch (e) {
    return jsonResponse("error", e.message);
  }
}
