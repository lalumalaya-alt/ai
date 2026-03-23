function setupHotelDatabase() {
  // Get the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Define our required tabs and their specific headers
  const databaseStructure = {
    'Rooms': ['Room_ID', 'Room_Type', 'Status', 'Price_Per_Night (₹)'],
    'Bookings': ['Booking_ID', 'Guest_Name', 'Phone', 'Room_ID', 'Booking_Date', 'Check_In_Date', 'Check_Out_Date', 'Status'],
    'Billing': ['Invoice_ID', 'Booking_ID', 'Room_Charges', 'Extra_Charges', 'Total_Amount (₹)', 'Payment_Status']
  };

  // Loop through each tab defined above
  for (const sheetName in databaseStructure) {
    let sheet = ss.getSheetByName(sheetName);

    // 1. Create the tab if it doesn't already exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // 2. Insert the headers into the first row
    const headers = databaseStructure[sheetName];
    // getRange(row, column, numRows, numColumns)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 3. Make it look nice (Bold text and freeze the top row)
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // Optional: Remove the default "Sheet1" if it's empty to keep things clean
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}
