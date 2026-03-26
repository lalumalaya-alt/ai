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

/**
 * Fetches available rooms for a given date range.
 * Currently, it simulates checking the database by returning mock available rooms.
 *
 * @param {string} checkIn The check-in datetime string.
 * @param {string} checkOut The check-out datetime string.
 * @return {string} A JSON string containing an array of available room objects.
 */
function getAvailableRooms(checkIn, checkOut) {
  // In a real application, you would connect to the active spreadsheet,
  // query the 'Bookings' and 'Rooms' sheets, and find rooms that are
  // not booked during the requested period.
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const roomsSheet = ss.getSheetByName('Rooms');
  // ... filter logic ...

  // Mock available rooms
  const mockAvailableRooms = [
    { roomId: '101', type: 'Standard Queen', price: 2500 },
    { roomId: '103', type: 'Deluxe King', price: 4000 },
    { roomId: '202', type: 'Ocean Suite', price: 8000 },
    { roomId: '203', type: 'Standard Twin', price: 2000 }
  ];

  return JSON.stringify(mockAvailableRooms);
}

/**
 * Handles the submission of an advanced booking form.
 *
 * @param {Object} formData The structured booking data from the frontend.
 * @return {string} A JSON string indicating the success or failure of the operation.
 */
function submitAdvancedBooking(formData) {
  try {
    // 1. Validate the form data
    if (!formData.guestName || !formData.phone || formData.selectedRooms.length === 0) {
      throw new Error('Missing required fields.');
    }

    // 2. Open the active spreadsheet and required sheets
    // const ss = SpreadsheetApp.getActiveSpreadsheet();
    // const bookingsSheet = ss.getSheetByName('Bookings');
    // const billingSheet = ss.getSheetByName('Billing');

    // 3. Generate IDs
    const bookingId = 'BK-' + Math.floor(1000 + Math.random() * 9000);
    const invoiceId = 'INV-' + Math.floor(1000 + Math.random() * 9000);

    // 4. Save to 'Bookings' sheet
    // Format: ['Booking_ID', 'Guest_Name', 'Phone', 'Room_ID', 'Booking_Date', 'Check_In_Date', 'Check_Out_Date', 'Status']
    // Example for single room mapping (for multi-room, you might insert multiple rows or join IDs)
    // bookingsSheet.appendRow([bookingId, formData.guestName, formData.phone, formData.selectedRooms.join(','), new Date(), formData.checkIn, formData.checkOut, 'Confirmed']);

    // 5. Save to 'Billing' sheet
    // Format: ['Invoice_ID', 'Booking_ID', 'Room_Charges', 'Extra_Charges', 'Total_Amount (₹)', 'Payment_Status']
    // billingSheet.appendRow([invoiceId, bookingId, formData.totalAmount, 0, formData.totalAmount, 'Pending']);

    // 6. Return success
    return JSON.stringify({
      success: true,
      message: 'Booking successfully created!',
      bookingId: bookingId
    });

  } catch (error) {
    // 7. Handle errors
    return JSON.stringify({
      success: false,
      message: error.message
    });
  }
}
