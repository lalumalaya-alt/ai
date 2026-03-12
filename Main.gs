/***************************************************
 * WEB APP ENTRY POINT
 ***************************************************/
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('MRI Hotel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

/***************************************************
 * SETUP DEMO DATA
 * Deletes ALL existing sheets, recreates them
 * with headers, and populates with generic demo data.
 * Run this once from the Script Editor to set up.
 ***************************************************/
function setupDemoData() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetNames = [LOGIN_SHEET_NAME, ROOMS_SHEET_NAME, BOOKINGS_SHEET_NAME, QUOTES_SHEET_NAME, FINANCE_SHEET_NAME, INVOICES_SHEET_NAME, SETTINGS_SHEET_NAME, BUDGETS_SHEET_NAME, CATEGORIES_SHEET_NAME, CUSTOMERS_SHEET_NAME, CHECKIN_SHEET_NAME, RESTAURANT_SHEET_NAME];

  // --- 1. Delete all existing sheets ---
  let tempSheet = ss.insertSheet("_TEMP_SETUP_");
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName() !== "_TEMP_SETUP_") {
      ss.deleteSheet(allSheets[i]);
    }
  }

  // --- 2. Create all sheets with headers ---

  // LOGIN
  const loginSheet = ss.insertSheet(LOGIN_SHEET_NAME);
  loginSheet.appendRow(["Username", "Password", "Role", "OTP", "OTPExpiry"]);

  // ROOMS
  const roomsSheet = ss.insertSheet(ROOMS_SHEET_NAME);
  roomsSheet.appendRow(["RoomNo", "RoomType", "RoomRate", "RoomStatus"]);

  // BOOKINGS
  const bookingsSheet = ss.insertSheet(BOOKINGS_SHEET_NAME);
  bookingsSheet.appendRow(["TicketID", "RoomNo", "GuestName", "Phone", "Email", "City", "MaritalStatus", "OccupancyType", "FamilyDetails", "CheckIn", "CheckOut", "Status", "RoomRate", "Discount", "Tax", "PaymentMethod", "TotalAmount", "PaymentStatus", "AmountPaid", "CheckInTime", "CheckOutTime", "FoodPlan", "AdvancePaid", "NumberOfRooms", "LinkedCheckInID"]);

  // QUOTES (26 columns)
  const quotesSheet = ss.insertSheet(QUOTES_SHEET_NAME);
  quotesSheet.appendRow(["QuoteID", "GuestName", "Phone", "Email", "CreatedDate", "ValidUntil", "Status", "Items", "SubTotal", "Tax", "Discount", "TotalAmount", "Notes", "CreatedBy", "Currency", "GSTEnabled", "GSTPercent", "GSTAmount", "GreenTaxEnabled", "GreenTaxPerNight", "GreenTaxPax", "GreenTaxNights", "GreenTaxAmount", "CustomerTIN", "ConvertedToInvoice", "PDFDriveLink"]);

  // FINANCE (12 columns)
  const financeSheet = ss.insertSheet(FINANCE_SHEET_NAME);
  financeSheet.appendRow(["ID", "Date", "Type", "Description", "ShopSource", "Amount", "Balance", "EnteredBy", "CreatedAt", "Category", "Currency", "LinkedInvoiceID"]);

  // INVOICES (26 columns)
  const invoicesSheet = ss.insertSheet(INVOICES_SHEET_NAME);
  invoicesSheet.appendRow(["InvoiceID", "GuestName", "Phone", "Email", "CustomerTIN", "Currency", "CreatedDate", "DueDate", "Status", "Items", "SubTotal", "GSTEnabled", "GSTPercent", "GSTAmount", "GreenTaxEnabled", "GreenTaxPerNight", "GreenTaxPax", "GreenTaxNights", "GreenTaxAmount", "Discount", "TotalAmount", "Notes", "SourceQuoteID", "PDFDriveLink", "CreatedBy", "UpdatedAt"]);

  // SETTINGS (14 columns, 1 data row)
  const settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  settingsSheet.appendRow(["HotelName", "HotelAddress", "HotelPhone", "HotelEmail", "HotelTIN", "LogoFileId", "LogoUrl", "DefaultCurrency", "GSTDefaultPercent", "GreenTaxDefaultRate", "NextInvoiceNum", "NextQuoteNum", "PDFDriveFolderId", "LogoDriveFolderId", "NextCheckInNum", "NextBillNum"]);
  settingsSheet.appendRow(["MRI Demo Hotel", "Demo Location, Maldives", "+960-0000000", "info@demo.com", "", "", "", "MVR", 5, 6, 5, 6, "", "", 4, 1]);

  // BUDGETS
  const budgetsSheet = ss.insertSheet(BUDGETS_SHEET_NAME);
  budgetsSheet.appendRow(["BudgetID", "Month", "Year", "BudgetAmount", "TotalSpent", "Remaining", "SetBy", "CreatedAt", "UpdatedAt"]);

  // CATEGORIES
  const categoriesSheet = ss.insertSheet(CATEGORIES_SHEET_NAME);
  categoriesSheet.appendRow(["CategoryID", "Name", "Type", "IsDefault", "CreatedBy", "CreatedAt"]);

  // CUSTOMERS
  const customersSheet = ss.insertSheet(CUSTOMERS_SHEET_NAME);
  customersSheet.appendRow(["CustomerID", "Name", "Phone", "Email", "City", "MaritalStatus", "Notes", "CreatedAt", "LinkedUsername"]);

  // CHECKIN (28 columns)
  const checkinSheet = ss.insertSheet(CHECKIN_SHEET_NAME);
  checkinSheet.appendRow(["CheckInID", "LinkedTicketID", "GuestName", "CompanyName", "GSTNumber", "IdentityProof", "Mobile", "Email", "Address", "PurposeOfVisit", "CheckInDate", "CheckInTime", "CheckOutDate", "CheckOutTime", "RoomNumbers", "RoomTypes", "NumberOfRooms", "Pax", "AdvancePaid", "ExtraPerson", "FoodPlan", "GSTType", "FixRoomRent", "FixRoomRentAmount", "BillTo", "DiscountPercent", "Status", "CreatedAt"]);

  // RESTAURANT (9 columns)
  const restaurantSheet = ss.insertSheet(RESTAURANT_SHEET_NAME);
  restaurantSheet.appendRow(["OrderID", "RoomNo", "CheckInID", "OrderDate", "Category", "Description", "Amount", "Status", "CreatedAt"]);

  // Delete temp sheet
  ss.deleteSheet(tempSheet);

  // --- 3. Populate demo data ---
  // ===== LOGIN (3 users) =====
  loginSheet.appendRow(["admin@demo.com", "admin123", "admin", "", ""]);
  loginSheet.appendRow(["user1@demo.com", "user123", "user", "", ""]);
  loginSheet.appendRow(["user2@demo.com", "user123", "user", "", ""]);
  loginSheet.appendRow(["client1@demo.com", "client123", "user", "", ""]);
  loginSheet.appendRow(["client2@demo.com", "client123", "user", "", ""]);

  // ===== ROOMS (10 rooms) =====
  const roomsData = [
    ["101", "Standard", 800,  "Available"],
    ["102", "Standard", 800,  "Booked"],
    ["103", "Deluxe",   1200, "Available"],
    ["104", "Deluxe",   1200, "Booked"],
    ["105", "Suite",    2500, "Available"],
    ["106", "Suite",    2500, "Booked"],
    ["107", "Family",   1800, "Booked"],
    ["108", "Single",   500,  "Reserved"],
    ["109", "Double",   1000, "Maintenance"],
    ["110", "Family",   1800, "Booked"]
  ];
  roomsSheet.getRange(2, 1, roomsData.length, 4).setValues(roomsData);

  // ===== BOOKINGS (9 bookings - varied dates/statuses for calendar testing, 25 columns) =====
  const bookingsData = [
    ["TKT-20260201-001", "104", "Demo Guest 1", "+960-1000001", "user1@demo.com",   "Demo City A", "Single",  "Single",   "",                  "2026-02-01T14:00:00Z", "2026-02-04T12:00:00Z", "Checked In",  1200, 0,   60,  "Cash",        3660,  "Unpaid",  0,    "14:00", "12:00", "Including Breakfast", 0, 1, "CHK-0001"],
    ["TKT-20260203-002", "107", "Demo Guest 2", "+960-1000002", "user2@demo.com",   "Demo City B", "Married", "Family",   "Spouse + 1 child",  "2026-02-03T14:00:00Z", "2026-02-06T12:00:00Z", "Booked",      1800, 100, 85,  "Card",        5385,  "Partial", 3000, "14:00", "12:00", "Including Breakfast and Dinner", 3000, 1, ""],
    ["TKT-20260110-003", "101", "Demo Guest 3", "+960-1000003", "guest3@demo.com",  "Demo City C", "Single",  "Single",   "",                  "2026-01-10T14:00:00Z", "2026-01-12T12:00:00Z", "Checked Out", 800,  0,   32,  "Cash",        1632,  "Paid",    1632, "14:00", "12:00", "None", 0, 1, ""],
    ["TKT-20260115-004", "103", "Demo Guest 4", "+960-1000004", "guest4@demo.com",  "Demo City A", "Married", "Couple",   "Spouse",            "2026-01-15T14:00:00Z", "2026-01-18T12:00:00Z", "Checked Out", 1200, 50,  71,  "Bank Transfer", 3621,  "Paid",    3621, "14:00", "12:00", "Including Breakfast", 1000, 1, ""],
    ["TKT-20260120-005", "105", "Demo Guest 5", "+960-1000005", "guest5@demo.com",  "Demo City D", "Single",  "Single",   "",                  "2026-01-20T14:00:00Z", "2026-01-23T12:00:00Z", "Checked Out", 2500, 200, 145, "Cash",        7445,  "Paid",    7445, "14:00", "12:00", "None", 2000, 1, ""],
    ["TKT-20260210-006", "108", "Demo Guest 6", "+960-1000006", "user1@demo.com",   "Demo City B", "Single",  "Single",   "",                  "2026-02-10T14:00:00Z", "2026-02-13T12:00:00Z", "Booked",      500,  0,   25,  "Card",        1525,  "Unpaid",  0,    "14:00", "12:00", "None", 0, 1, ""],
    ["TKT-20260215-007", "106", "Demo Guest 7", "+960-1000007", "user2@demo.com",   "Demo City E", "Married", "Couple",   "Spouse",            "2026-02-15T14:00:00Z", "2026-02-18T12:00:00Z", "Booked",      2500, 0,   125, "Card",        7625,  "Paid",    7625, "14:00", "12:00", "Including Breakfast", 5000, 1, ""],
    ["TKT-20260220-008", "110", "Demo Guest 8", "+960-1000008", "guest8@demo.com",  "Demo City A", "Single",  "Family",   "2 children",        "2026-02-20T14:00:00Z", "2026-02-25T12:00:00Z", "Checked In",  1800, 200, 80,  "Bank Transfer", 8880,  "Unpaid",  0,    "14:00", "12:00", "Including Breakfast and Dinner", 0, 1, "CHK-0002"],
    ["TKT-20260225-009", "102", "Demo Guest 9", "+960-1000009", "guest9@demo.com",  "Demo City C", "Single",  "Double",   "",                  "2026-02-25T14:00:00Z", "2026-02-28T12:00:00Z", "Checked In",  800,  0,   48,  "Cash",        2448,  "Unpaid",  0,    "14:00", "12:00", "None", 0, 1, "CHK-0003"]
  ];
  bookingsSheet.getRange(2, 1, bookingsData.length, 25).setValues(bookingsData);

  // ===== QUOTES (4 quotes, 26 columns) =====
  const quoteItems1 = JSON.stringify([
    { type: "room", description: "Deluxe Room", roomType: "Deluxe", quantity: 1, nights: 3, rate: 1200, amount: 3600 },
    { type: "service", description: "Airport Transfer", amount: 150 }
  ]);
  const quoteItems2 = JSON.stringify([
    { type: "room", description: "Suite Room", roomType: "Suite", quantity: 1, nights: 5, rate: 2500, amount: 12500 },
    { type: "activity", description: "Sunset Cruise", pax: 2, rate: 200, amount: 400 },
    { type: "service", description: "Airport Transfer", amount: 150 }
  ]);
  const quoteItems3 = JSON.stringify([
    { type: "room", description: "Standard Room", roomType: "Standard", quantity: 2, nights: 2, rate: 800, amount: 3200 }
  ]);
  const quoteItems4 = JSON.stringify([
    { type: "room", description: "Family Room", roomType: "Family", quantity: 1, nights: 4, rate: 1800, amount: 7200 },
    { type: "activity", description: "Snorkeling Trip", pax: 4, rate: 150, amount: 600 },
    { type: "service", description: "Spa Package", amount: 350 }
  ]);

  const quotesData = [
    ["QTN-0001", "Demo Client 1", "+960-2000001", "client1@demo.com", "2026-02-01T10:00:00Z", "2026-02-15T23:59:59Z", "Sent",     quoteItems1, 3750, 0, 0, 4350, "Includes breakfast", "admin@demo.com", "MVR", true, 16, 600, false, 6, 0, 0, 0, "", "", ""],
    ["QTN-0002", "Demo Client 2", "+960-2000002", "client2@demo.com", "2026-02-05T11:00:00Z", "2026-02-20T23:59:59Z", "Draft",    quoteItems2, 13050, 0, 500, 14558, "VIP demo guest", "admin@demo.com", "USD", true, 16, 2008, false, 6, 0, 0, 0, "", "", ""],
    ["QTN-0003", "Demo Client 3", "+960-2000003", "client3@demo.com", "2026-02-08T09:00:00Z", "2026-02-22T23:59:59Z", "Accepted", quoteItems3, 3200, 0, 0, 3736, "", "admin@demo.com", "MVR", true, 16, 512, true, 6, 2, 2, 24, "", "", ""],
    ["QTN-0004", "Demo Client 4", "+960-2000004", "client4@demo.com", "2026-02-10T15:00:00Z", "2026-01-25T23:59:59Z", "Expired",  quoteItems4, 8150, 0, 100, 9434, "Demo family vacation", "admin@demo.com", "MVR", true, 16, 1288, true, 6, 4, 4, 96, "", "", ""]
  ];
  quotesSheet.getRange(2, 1, quotesData.length, 26).setValues(quotesData);

  // Extra quote for Feb 21 reservation testing
  const quoteItems5 = JSON.stringify([
    { type: "room", description: "Single Room", roomType: "Single", quantity: 1, nights: 3, rate: 500, amount: 1500, reservedRoomNo: "108" },
    { type: "service", description: "Airport Transfer", amount: 150 }
  ]);
  quotesSheet.appendRow(["QTN-0005", "Demo Client 1", "+960-2000001", "client1@demo.com", "2026-02-21T10:00:00Z", "2026-03-07T23:59:59Z", "Accepted", quoteItems5, 1650, 0, 0, 1914, "Room reserve test - Feb 21", "admin@demo.com", "MVR", true, 16, 264, false, 6, 0, 0, 0, "", "", ""]);

  // ===== INVOICES (3 invoices) =====
  const invItems1 = JSON.stringify([
    { type: "room", roomType: "Deluxe", quantity: 2, nights: 3, rate: 1200, amount: 7200 },
    { type: "activity", description: "Snorkeling Trip", pax: 4, rate: 150, amount: 600 },
    { type: "service", description: "Airport Transfer", amount: 150 }
  ]);
  const invItems2 = JSON.stringify([
    { type: "room", roomType: "Suite", quantity: 1, nights: 5, rate: 2500, amount: 12500 },
    { type: "service", description: "Spa Package", amount: 500 }
  ]);
  const invItems3 = JSON.stringify([
    { type: "room", roomType: "Standard", quantity: 1, nights: 2, rate: 800, amount: 1600 },
    { type: "service", description: "Laundry", amount: 50 }
  ]);
  const invItems4 = JSON.stringify([
    { type: "room", roomType: "Family", quantity: 1, nights: 4, rate: 1800, amount: 7200 },
    { type: "service", description: "Airport Transfer", amount: 150 }
  ]);

  const invoicesData = [
    ["INV-0001", "Demo Guest 1", "+960-1000001", "user1@demo.com", "TIN-00001", "MVR", "2026-02-01T10:00:00Z", "2026-03-01T23:59:59Z", "Paid", invItems1, 7950, true, 16, 1272, true, 6, 4, 3, 72, 0, 9294, "Deluxe demo package", "", "", "admin@demo.com", "2026-02-01T10:00:00Z"],
    ["INV-0002", "Demo Guest 2", "+960-1000002", "user2@demo.com", "", "USD", "2026-02-05T11:00:00Z", "2026-02-15T23:59:59Z", "Sent", invItems2, 13000, true, 16, 2000, false, 6, 0, 0, 0, 500, 14500, "Demo suite package", "", "", "admin@demo.com", "2026-02-05T11:00:00Z"],
    ["INV-0003", "Demo Guest 3", "+960-1000003", "guest3@demo.com", "TIN-00003", "MVR", "2026-02-10T09:00:00Z", "2026-03-10T23:59:59Z", "Draft", invItems3, 1650, true, 16, 264, true, 6, 1, 2, 12, 0, 1926, "Demo standard booking", "", "", "admin@demo.com", "2026-02-10T09:00:00Z"],
    ["INV-0004", "Demo Guest 4", "+960-1000004", "guest4@demo.com", "", "MVR", "2026-01-20T10:00:00Z", "2026-02-01T23:59:59Z", "Sent", invItems4, 7350, true, 16, 1176, false, 6, 0, 0, 0, 0, 8526, "Demo overdue test", "", "", "admin@demo.com", "2026-01-20T10:00:00Z"]
  ];
  invoicesSheet.getRange(2, 1, invoicesData.length, 26).setValues(invoicesData);

  // ===== FINANCE (13 records, 12 columns) =====
  const financeData = [
    ["FIN-20260101-001", "2026-01-12T10:00:00Z", "Income",  "Room Checkout - Demo Guest 3",  "Room 101",            1632,  1632,   "admin@demo.com", "2026-01-12T12:05:00Z", "Room Checkout",    "MVR", ""],
    ["FIN-20260102-002", "2026-01-15T09:00:00Z", "Expense", "Electricity Bill - January",    "Demo Utility Co",     2800,  -1168,  "admin@demo.com", "2026-01-15T09:30:00Z", "Utilities",        "MVR", ""],
    ["FIN-20260103-003", "2026-01-18T11:00:00Z", "Income",  "Room Checkout - Demo Guest 4",  "Room 103",            3621,  2453,   "admin@demo.com", "2026-01-18T12:00:00Z", "Room Checkout",    "MVR", ""],
    ["FIN-20260104-004", "2026-01-20T14:00:00Z", "Expense", "Kitchen Supplies Restock",      "Demo Supplier A",     1500,  953,    "admin@demo.com", "2026-01-20T14:30:00Z", "Kitchen Supplies", "MVR", ""],
    ["FIN-20260105-005", "2026-01-23T10:00:00Z", "Income",  "Room Checkout - Demo Guest 5",  "Room 105",            7445,  8398,   "admin@demo.com", "2026-01-23T10:15:00Z", "Room Checkout",    "MVR", ""],
    ["FIN-20260106-006", "2026-01-25T08:00:00Z", "Expense", "Staff Salaries - January",      "Demo Payroll",        5000,  3398,   "admin@demo.com", "2026-01-25T08:00:00Z", "Staff Salaries",   "MVR", ""],
    ["FIN-20260107-007", "2026-01-28T16:00:00Z", "Income",  "Restaurant Sales - Week 4",     "Demo Restaurant",     3200,  6598,   "admin@demo.com", "2026-01-28T16:00:00Z", "Restaurant",       "MVR", ""],
    ["FIN-20260108-008", "2026-02-01T09:00:00Z", "Expense", "Water Bill - January",          "Demo Utility Co",     950,   5648,   "admin@demo.com", "2026-02-01T09:00:00Z", "Utilities",        "MVR", ""],
    ["FIN-20260113-013", "2026-02-01T10:00:00Z", "Income",  "Payment for INV-0001",          "Demo Invoice Payment", 9294, 14942,  "admin@demo.com", "2026-02-01T10:05:00Z", "Invoice Payment",  "MVR", "INV-0001"],
    ["FIN-20260109-009", "2026-02-03T11:00:00Z", "Income",  "Event Booking - Demo Conference","Demo Events Hall",    4500,  19442,  "admin@demo.com", "2026-02-03T11:00:00Z", "Events",           "MVR", ""],
    ["FIN-20260110-010", "2026-02-05T14:00:00Z", "Expense", "Laundry Service Supplies",      "Demo Supplier B",     800,   18642,  "admin@demo.com", "2026-02-05T14:30:00Z", "Laundry",          "MVR", ""],
    ["FIN-20260111-011", "2026-02-08T10:00:00Z", "Income",  "Spa Services - Week 1 Feb",     "Demo Spa",            2100,  20742,  "admin@demo.com", "2026-02-08T10:00:00Z", "Spa",              "MVR", ""],
    ["FIN-20260112-012", "2026-02-10T13:00:00Z", "Expense", "Maintenance - AC Repair",       "Demo Maintenance Co", 1350,  19392,  "admin@demo.com", "2026-02-10T13:00:00Z", "Maintenance",      "MVR", ""]
  ];
  financeSheet.getRange(2, 1, financeData.length, 12).setValues(financeData);

  // ===== CATEGORIES (default categories) =====
  const now = new Date().toISOString();
  const defaultCategories = [
    ["CAT-EXP-001", "Utilities",        "Expense", true, "system", now],
    ["CAT-EXP-002", "Kitchen Supplies",  "Expense", true, "system", now],
    ["CAT-EXP-003", "Staff Salaries",    "Expense", true, "system", now],
    ["CAT-EXP-004", "Maintenance",       "Expense", true, "system", now],
    ["CAT-EXP-005", "Laundry",           "Expense", true, "system", now],
    ["CAT-EXP-006", "Marketing",         "Expense", true, "system", now],
    ["CAT-EXP-007", "Miscellaneous",     "Expense", true, "system", now],
    ["CAT-INC-001", "Room Checkout",     "Income",  true, "system", now],
    ["CAT-INC-002", "Restaurant",        "Income",  true, "system", now],
    ["CAT-INC-003", "Events",            "Income",  true, "system", now],
    ["CAT-INC-004", "Spa",               "Income",  true, "system", now],
    ["CAT-INC-005", "Excursions",        "Income",  true, "system", now],
    ["CAT-INC-006", "Fishing Trips",     "Income",  true, "system", now],
    ["CAT-INC-007", "Other Income",      "Income",  true, "system", now],
    ["CAT-INC-008", "Invoice Payment",   "Income",  true, "system", now]
  ];
  categoriesSheet.getRange(2, 1, defaultCategories.length, 6).setValues(defaultCategories);

  // ===== BUDGETS (current month) =====
  const nowDate = new Date();
  budgetsSheet.appendRow([
    "BDG-" + nowDate.getFullYear() + "-" + String(nowDate.getMonth() + 1).padStart(2, '0'),
    nowDate.getMonth() + 1,
    nowDate.getFullYear(),
    50000,
    3100,
    46900,
    "admin@demo.com",
    now,
    now
  ]);

  // ===== CUSTOMERS (6 demo customers) =====
  const customersData = [
    ["CUST-000001", "Demo Guest 1",  "+960-1000001", "user1@demo.com",    "Demo City A", "Single",  "VIP guest",        now, "user1@demo.com"],
    ["CUST-000002", "Demo Guest 2",  "+960-1000002", "user2@demo.com",    "Demo City B", "Married", "Family traveller", now, "user2@demo.com"],
    ["CUST-000003", "Demo Guest 3",  "+960-1000003", "guest3@demo.com",   "Demo City C", "Single",  "Regular customer", now, "guest3@demo.com"],
    ["CUST-000004", "Demo Client 1", "+960-2000001", "client1@demo.com",  "Demo City D", "Married", "Corporate client", now, "client1@demo.com"],
    ["CUST-000005", "Demo Client 2", "+960-2000002", "client2@demo.com",  "Demo City E", "Single",  "Travel agency",    now, "client2@demo.com"],
    ["CUST-000006", "Walk-in Guest", "+960-3000001", "",                  "Demo City F", "Single",  "Walk-in",          now, ""]
  ];
  customersSheet.getRange(2, 1, customersData.length, 9).setValues(customersData);

  // ===== CHECKIN (3 active check-ins linked to demo bookings) =====
  const demoNow = new Date().toISOString();
  const checkinData = [
    ["CHK-0001", "TKT-20260201-001", "Demo Guest 1", "", "", "", "+960-1000001", "user1@demo.com",  "Demo City A", "Leisure",  "2026-02-01T14:00:00Z", "14:00", "2026-02-04T12:00:00Z", "12:00", "104", "Deluxe",   1, 2, 0, 0, "Including Breakfast",           "Excluding", "No", 0, "Individual", 0, "Active", demoNow],
    ["CHK-0002", "TKT-20260220-008", "Demo Guest 8", "", "", "", "+960-1000008", "guest8@demo.com", "Demo City A", "Leisure",  "2026-02-20T14:00:00Z", "14:00", "2026-02-25T12:00:00Z", "12:00", "110", "Family",   1, 4, 0, 2, "Including Breakfast and Dinner", "Excluding", "No", 0, "Individual", 0, "Active", demoNow],
    ["CHK-0003", "TKT-20260225-009", "Demo Guest 9", "", "", "", "+960-1000009", "guest9@demo.com", "Demo City C", "Business", "2026-02-25T14:00:00Z", "14:00", "2026-02-28T12:00:00Z", "12:00", "102", "Standard", 1, 1, 0, 0, "None",                          "Excluding", "No", 0, "Individual", 0, "Active", demoNow]
  ];
  checkinSheet.getRange(2, 1, checkinData.length, 28).setValues(checkinData);

  // ===== RESTAURANT (4 demo orders for active check-in rooms) =====
  const restaurantData = [
    ["ORD-0001", "104", "CHK-0001", "2026-02-01", "FoodBeverage", "Lunch for 2",         250, "Active", demoNow],
    ["ORD-0002", "104", "CHK-0001", "2026-02-02", "Laundry",      "2 shirts, 1 trouser", 150, "Active", demoNow],
    ["ORD-0003", "110", "CHK-0002", "2026-02-21", "FoodBeverage", "Dinner for 4",        600, "Active", demoNow],
    ["ORD-0004", "110", "CHK-0002", "2026-02-22", "ExtraBed",     "Extra mattress",      500, "Active", demoNow]
  ];
  restaurantSheet.getRange(2, 1, restaurantData.length, 9).setValues(restaurantData);

  // --- 4. Format header rows ---
  [loginSheet, roomsSheet, bookingsSheet, quotesSheet, financeSheet, invoicesSheet, settingsSheet, budgetsSheet, categoriesSheet, customersSheet, checkinSheet, restaurantSheet].forEach(function(sheet) {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#001f3f");
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  });

  SpreadsheetApp.getUi().alert("Demo data setup complete!\n\nLogin credentials:\n• admin@demo.com / admin123 (Admin)\n• user1@demo.com / user123 (User)\n• user2@demo.com / user123 (User)\n• client1@demo.com / client123 (Client)\n• client2@demo.com / client123 (Client)\n\nSheets created: " + sheetNames.join(", ") + "\n\nNew sheets added: CheckIn, Restaurant");
}
