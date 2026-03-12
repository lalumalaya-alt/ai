/***************************************************
 * CHECK-IN FUNCTIONS
 ***************************************************/
function addCheckIn(checkInData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const ciSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!ciSheet) return { success: false, message: "CheckIn sheet not found. Run Setup Demo Data." };

    const checkInId = generateCheckInId();
    const now = new Date().toISOString();

    const roomNumbers = checkInData.roomNumbers || '';
    const roomNosArr = roomNumbers.split(',').map(r => r.trim()).filter(r => r);

    // Get room types for selected rooms
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const roomsData = roomsSheet.getDataRange().getValues();
    let roomTypes = [];
    for (let r = 0; r < roomNosArr.length; r++) {
      for (let i = 1; i < roomsData.length; i++) {
        if ((roomsData[i][ROOM_NO_COL] || '').toString() === roomNosArr[r]) {
          roomTypes.push((roomsData[i][ROOM_TYPE_COL] || '').toString());
          // Mark room as Booked
          roomsSheet.getRange(i + 1, ROOM_STATUS_COL + 1).setValue("Booked");
          break;
        }
      }
    }

    ciSheet.appendRow([
      checkInId,
      checkInData.linkedTicketId || '',
      checkInData.guestName || '',
      checkInData.companyName || '',
      checkInData.gstNumber || '',
      checkInData.identityProof || '',
      checkInData.mobile || '',
      checkInData.email || '',
      checkInData.address || '',
      checkInData.purposeOfVisit || '',
      checkInData.checkInDate || '',
      checkInData.checkInTime || '14:00',
      checkInData.checkOutDate || '',
      checkInData.checkOutTime || '12:00',
      roomNumbers,
      roomTypes.join(','),
      roomNosArr.length,
      parseInt(checkInData.pax) || 1,
      parseFloat(checkInData.advancePaid) || 0,
      parseInt(checkInData.extraPerson) || 0,
      checkInData.foodPlan || 'None',
      checkInData.gstType || 'Excluding',
      checkInData.fixRoomRent || 'No',
      parseFloat(checkInData.fixRoomRentAmount) || 0,
      checkInData.billTo || 'Individual',
      parseFloat(checkInData.discountPercent) || 0,
      'Active',
      now
    ]);

    // If linked to advance booking, update booking status
    if (checkInData.linkedTicketId) {
      const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
      const bData = bookingsSheet.getDataRange().getValues();
      for (let i = 1; i < bData.length; i++) {
        if ((bData[i][TICKET_ID_COL] || '').toString() === checkInData.linkedTicketId) {
          bookingsSheet.getRange(i + 1, BOOKING_STATUS_COL + 1).setValue("Checked In");
          bookingsSheet.getRange(i + 1, LINKED_CHECKIN_COL + 1).setValue(checkInId);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: `Check-in successful. ID: ${checkInId}`, checkInId };
  } catch (e) {
    Logger.log("Error in addCheckIn: " + e.toString());
    return { success: false, message: e.message };
  }
}

function getAllCheckIns() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    let checkIns = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      checkIns.push({
        rowIndex: i + 1,
        checkInId: (row[CI_ID_COL] || '').toString(),
        linkedTicketId: (row[CI_LINKED_TICKET_COL] || '').toString(),
        guestName: (row[CI_GUEST_NAME_COL] || '').toString(),
        companyName: (row[CI_COMPANY_COL] || '').toString(),
        gstNumber: (row[CI_GST_NUMBER_COL] || '').toString(),
        identityProof: (row[CI_IDENTITY_COL] || '').toString(),
        mobile: (row[CI_MOBILE_COL] || '').toString(),
        email: (row[CI_EMAIL_COL] || '').toString(),
        address: (row[CI_ADDRESS_COL] || '').toString(),
        purposeOfVisit: (row[CI_PURPOSE_COL] || '').toString(),
        checkInDate: (row[CI_CHECKIN_DATE_COL] || '').toString(),
        checkInTime: (row[CI_CHECKIN_TIME_COL] || '14:00').toString(),
        checkOutDate: (row[CI_CHECKOUT_DATE_COL] || '').toString(),
        checkOutTime: (row[CI_CHECKOUT_TIME_COL] || '12:00').toString(),
        roomNumbers: (row[CI_ROOM_NUMBERS_COL] || '').toString(),
        roomTypes: (row[CI_ROOM_TYPES_COL] || '').toString(),
        numberOfRooms: parseInt(row[CI_NUM_ROOMS_COL]) || 0,
        pax: parseInt(row[CI_PAX_COL]) || 1,
        advancePaid: parseFloat(row[CI_ADVANCE_PAID_COL]) || 0,
        extraPerson: parseInt(row[CI_EXTRA_PERSON_COL]) || 0,
        foodPlan: (row[CI_FOOD_PLAN_COL] || 'None').toString(),
        gstType: (row[CI_GST_TYPE_COL] || 'Excluding').toString(),
        fixRoomRent: (row[CI_FIX_RENT_COL] || 'No').toString(),
        fixRoomRentAmount: parseFloat(row[CI_FIX_RENT_AMT_COL]) || 0,
        billTo: (row[CI_BILL_TO_COL] || 'Individual').toString(),
        discountPercent: parseFloat(row[CI_DISCOUNT_COL]) || 0,
        status: (row[CI_STATUS_COL] || 'Active').toString(),
        createdAt: (row[CI_CREATED_AT_COL] || '').toString()
      });
    }
    return checkIns;
  } catch (e) {
    Logger.log("Error in getAllCheckIns: " + e.toString());
    return { error: e.message };
  }
}

function getCheckInByRoomNo(roomNo) {
  try {
    const checkIns = getAllCheckIns();
    if (!checkIns.error) {
      for (let i = 0; i < checkIns.length; i++) {
        if (checkIns[i].status === 'Active') {
          let rooms = checkIns[i].roomNumbers.split(',').map(r => r.trim());
          if (rooms.indexOf(roomNo.toString().trim()) !== -1) {
            return checkIns[i];
          }
        }
      }
    }
    // Fallback: search Bookings for a room with active booking but no check-in record
    const ss = SpreadsheetApp.openById(SS_ID);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    if (bookingsSheet && bookingsSheet.getLastRow() > 1) {
      const bData = bookingsSheet.getDataRange().getValues();
      for (let i = 1; i < bData.length; i++) {
        let bStatus = (bData[i][BOOKING_STATUS_COL] || '').toString();
        if (bStatus !== 'Booked' && bStatus !== 'Checked In') continue;
        let bRooms = (bData[i][BOOKING_ROOM_NO_COL] || '').toString().split(',').map(r => r.trim());
        if (bRooms.indexOf(roomNo.toString().trim()) !== -1) {
          return {
            checkInId: (bData[i][TICKET_ID_COL] || '').toString(),
            linkedTicketId: (bData[i][TICKET_ID_COL] || '').toString(),
            guestName: (bData[i][GUEST_NAME_COL] || '').toString(),
            companyName: '', gstNumber: '', identityProof: '',
            mobile: (bData[i][PHONE_COL] || '').toString(),
            email: (bData[i][EMAIL_COL] || '').toString(),
            address: (bData[i][CITY_COL] || '').toString(),
            purposeOfVisit: '',
            checkInDate: (bData[i][CHECK_IN_COL] || '').toString(),
            checkInTime: (bData[i][CHECKIN_TIME_COL] || '14:00').toString(),
            checkOutDate: (bData[i][CHECK_OUT_COL] || '').toString(),
            checkOutTime: (bData[i][CHECKOUT_TIME_COL] || '12:00').toString(),
            roomNumbers: (bData[i][BOOKING_ROOM_NO_COL] || '').toString(),
            roomTypes: '',
            numberOfRooms: parseInt(bData[i][NUM_ROOMS_COL]) || 1,
            pax: 1,
            advancePaid: parseFloat(bData[i][ADVANCE_PAID_COL]) || 0,
            extraPerson: 0,
            foodPlan: (bData[i][FOOD_PLAN_COL] || 'None').toString(),
            gstType: 'Excluding', fixRoomRent: 'No', fixRoomRentAmount: 0,
            billTo: 'Individual',
            discountPercent: parseFloat(bData[i][DISCOUNT_COL]) || 0,
            status: 'Active', createdAt: '', isFromBooking: true
          };
        }
      }
    }
    return null;
  } catch (e) {
    Logger.log("Error in getCheckInByRoomNo: " + e.toString());
    return null;
  }
}

function updateCheckIn(rowIndex, checkInData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    if (!sheet) return { success: false, message: "CheckIn sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    const existingId = sheet.getRange(rowIndex, CI_ID_COL + 1).getValue();
    const existingLinked = sheet.getRange(rowIndex, CI_LINKED_TICKET_COL + 1).getValue();
    const existingStatus = sheet.getRange(rowIndex, CI_STATUS_COL + 1).getValue();
    const existingCreatedAt = sheet.getRange(rowIndex, CI_CREATED_AT_COL + 1).getValue();

    const roomNumbers = checkInData.roomNumbers || '';
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const roomsData = roomsSheet.getDataRange().getValues();
    let roomNosArr = roomNumbers.split(',').map(r => r.trim()).filter(r => r);
    let roomTypes = [];
    for (let r = 0; r < roomNosArr.length; r++) {
      for (let i = 1; i < roomsData.length; i++) {
        if ((roomsData[i][ROOM_NO_COL] || '').toString() === roomNosArr[r]) {
          roomTypes.push((roomsData[i][ROOM_TYPE_COL] || '').toString());
          break;
        }
      }
    }

    const row = [
      existingId, existingLinked,
      checkInData.guestName || '', checkInData.companyName || '', checkInData.gstNumber || '',
      checkInData.identityProof || '', checkInData.mobile || '', checkInData.email || '',
      checkInData.address || '', checkInData.purposeOfVisit || '',
      checkInData.checkInDate || '', checkInData.checkInTime || '14:00',
      checkInData.checkOutDate || '', checkInData.checkOutTime || '12:00',
      roomNumbers, roomTypes.join(','), roomNosArr.length,
      parseInt(checkInData.pax) || 1, parseFloat(checkInData.advancePaid) || 0,
      parseInt(checkInData.extraPerson) || 0, checkInData.foodPlan || 'None',
      checkInData.gstType || 'Excluding', checkInData.fixRoomRent || 'No',
      parseFloat(checkInData.fixRoomRentAmount) || 0, checkInData.billTo || 'Individual',
      parseFloat(checkInData.discountPercent) || 0, existingStatus, existingCreatedAt
    ];

    sheet.getRange(rowIndex, 1, 1, 28).setValues([row]);
    SpreadsheetApp.flush();
    return { success: true, message: "Check-in updated successfully." };
  } catch (e) {
    Logger.log("Error in updateCheckIn: " + e.toString());
    return { success: false, message: e.message };
  }
}

/***************************************************
 * CHECKOUT FUNCTIONS
 ***************************************************/
function checkoutRoom(ticketId, paymentOverride) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const bookingsData = bookingsSheet.getDataRange().getValues();
    const roomsData = roomsSheet.getDataRange().getValues();

    let defaultCurrency = 'MVR';
    try {
      const setSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
      if (setSheet && setSheet.getLastRow() > 1) {
        defaultCurrency = (setSheet.getRange(2, SET_DEFAULT_CURRENCY_COL + 1).getValue() || 'MVR').toString();
      }
    } catch (ce) { Logger.log("Could not load settings currency: " + ce); }

    let bookingRowIndex = -1;
    let roomNoToCheckout = "";
    let guestName = "";
    let email = "";
    let phone = "";
    let city = "";
    let checkInDate, checkOutDate;
    let roomRate, discount, tax, paymentMethod;

    for (let i = 1; i < bookingsData.length; i++) {
      if (bookingsData[i][TICKET_ID_COL].toString() === ticketId.toString()) {
        let status = (bookingsData[i][BOOKING_STATUS_COL] || "").toString().toLowerCase();
        if (status === "completed" || status === "checked out") {
          return { success: false, message: `Ticket ID ${ticketId} has already been checked out.` };
        }
        bookingRowIndex = i;
        roomNoToCheckout = bookingsData[i][BOOKING_ROOM_NO_COL];
        guestName = bookingsData[i][GUEST_NAME_COL];
        phone = bookingsData[i][PHONE_COL];
        email = bookingsData[i][EMAIL_COL];
        city = bookingsData[i][CITY_COL];
        checkInDate = new Date(bookingsData[i][CHECK_IN_COL]);
        checkOutDate = new Date(bookingsData[i][CHECK_OUT_COL]);
        roomRate = parseFloat(bookingsData[i][ROOM_RATE_BOOK_COL]) || 0;
        discount = parseFloat(bookingsData[i][DISCOUNT_COL]) || 0;
        tax = parseFloat(bookingsData[i][TAX_COL]) || 0;
        paymentMethod = (bookingsData[i][PAYMENT_METHOD_COL] || "Cash").toString();
        break;
      }
    }
    if (bookingRowIndex === -1) {
      return { success: false, message: `Ticket ID ${ticketId} not found.` };
    }

    let actualCheckOut = new Date();
    checkOutDate = actualCheckOut;

    let nights = daysBetween(checkInDate, checkOutDate);
    if (nights < 1) nights = 1;
    let baseAmount = roomRate * nights;
    let finalAmount = (baseAmount - discount) + tax;

    let amountPaid = 0;
    let paymentStatus = "Unpaid";
    if (paymentOverride) {
      amountPaid = parseFloat(paymentOverride.amountPaid) || 0;
      if (paymentOverride.paymentMethod) paymentMethod = paymentOverride.paymentMethod;
    }
    if (amountPaid >= finalAmount) paymentStatus = "Paid";
    else if (amountPaid > 0) paymentStatus = "Partial";
    let balance = finalAmount - amountPaid;

    bookingsSheet.getRange(bookingRowIndex + 1, CHECK_OUT_COL + 1).setValue(checkOutDate.toISOString());
    bookingsSheet.getRange(bookingRowIndex + 1, BOOKING_STATUS_COL + 1).setValue("Checked Out");
    bookingsSheet.getRange(bookingRowIndex + 1, TOTAL_AMOUNT_COL + 1).setValue(finalAmount);
    bookingsSheet.getRange(bookingRowIndex + 1, PAYMENT_METHOD_COL + 1).setValue(paymentMethod);
    bookingsSheet.getRange(bookingRowIndex + 1, PAYMENT_STATUS_COL + 1).setValue(paymentStatus);
    bookingsSheet.getRange(bookingRowIndex + 1, AMOUNT_PAID_COL + 1).setValue(amountPaid);

    let roomRowIndex = -1;
    for (let j = 1; j < roomsData.length; j++) {
      let rowRoomNo = (roomsData[j][ROOM_NO_COL] || "").toString();
      if (rowRoomNo === roomNoToCheckout.toString()) {
        roomRowIndex = j;
        break;
      }
    }
    if (roomRowIndex !== -1) {
      roomsSheet.getRange(roomRowIndex + 1, ROOM_STATUS_COL + 1).setValue("Available");
    }
    SpreadsheetApp.flush();

    let invoiceHtml = generateInvoiceHtml({
      ticketId, occupantName: guestName, email, roomNo: roomNoToCheckout,
      checkIn: checkInDate.toISOString(), checkOut: checkOutDate.toISOString(),
      nights, roomRate, discount, tax, finalAmount, currency: defaultCurrency
    });

    try {
      let subject = `Invoice for your stay: Ticket ${ticketId}`;
      let bodyText = `Hello ${guestName},\n\nThank you for staying with us. Total: ${defaultCurrency} ${finalAmount.toFixed(2)}\n\nSafe travels!`;
      MailApp.sendEmail({
        to: email, subject, body: bodyText, htmlBody: invoiceHtml
      });
    } catch (emailErr) {
      Logger.log(`Email failed for checkout ${ticketId}: ${emailErr.message}`);
    }

    return {
      success: true,
      message: `Room ${roomNoToCheckout} (Ticket: ${ticketId}) checked out successfully.`,
      invoiceHtml,
      invoiceData: {
        ticketId: ticketId, guestName: guestName, phone: phone, email: email, city: city,
        roomNo: roomNoToCheckout.toString(), checkIn: checkInDate.toISOString(), checkOut: checkOutDate.toISOString(),
        nights: nights, roomRate: roomRate, baseAmount: baseAmount, discount: discount, tax: tax,
        paymentMethod: paymentMethod, finalAmount: finalAmount, paymentStatus: paymentStatus,
        amountPaid: amountPaid, balance: balance
      }
    };
  } catch (e) {
    Logger.log(`Error in checkoutRoom: ${e.toString()}`);
    return { success: false, message: `An error occurred during checkout: ${e.message}` };
  }
}

function processCheckoutPayment(ticketId, amountPaid, paymentMethod) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let roomRate = 0, discount = 0, tax = 0, checkInDate, checkOutDate;

    for (let i = 1; i < data.length; i++) {
      if (data[i][TICKET_ID_COL].toString() === ticketId.toString()) {
        let status = (data[i][BOOKING_STATUS_COL] || "").toString().toLowerCase();
        if (status === "checked out") {
          return { success: false, message: "Cannot update payment for a checked-out booking." };
        }
        rowIndex = i;
        checkInDate = new Date(data[i][CHECK_IN_COL]);
        checkOutDate = new Date(data[i][CHECK_OUT_COL]);
        roomRate = parseFloat(data[i][ROOM_RATE_BOOK_COL]) || 0;
        discount = parseFloat(data[i][DISCOUNT_COL]) || 0;
        tax = parseFloat(data[i][TAX_COL]) || 0;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `Ticket ID ${ticketId} not found.` };
    }

    let nights = daysBetween(checkInDate, checkOutDate);
    if (nights < 1) nights = 1;
    const finalAmount = (roomRate * nights) - discount + tax;
    const paid = parseFloat(amountPaid) || 0;

    let paymentStatus = "Unpaid";
    if (paid >= finalAmount) paymentStatus = "Paid";
    else if (paid > 0) paymentStatus = "Partial";

    sheet.getRange(rowIndex + 1, PAYMENT_STATUS_COL + 1).setValue(paymentStatus);
    sheet.getRange(rowIndex + 1, AMOUNT_PAID_COL + 1).setValue(paid);
    if (paymentMethod) {
      sheet.getRange(rowIndex + 1, PAYMENT_METHOD_COL + 1).setValue(paymentMethod);
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      message: `Payment recorded: MVR ${paid.toFixed(2)} (${paymentStatus})`,
      paymentStatus,
      amountPaid: paid,
      balance: finalAmount - paid
    };
  } catch (e) {
    Logger.log(`Error in processCheckoutPayment: ${e.toString()}`);
    return { success: false, message: `An error occurred: ${e.message}` };
  }
}

function processFullCheckout(checkInId, checkoutData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const ciSheet = ss.getSheetByName(CHECKIN_SHEET_NAME);
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    const restSheet = ss.getSheetByName(RESTAURANT_SHEET_NAME);

    const ciData = ciSheet ? ciSheet.getDataRange().getValues() : [[]];
    let ciRowIndex = -1;
    let ci = null;
    for (let i = 1; i < ciData.length; i++) {
      if ((ciData[i][CI_ID_COL] || '').toString() === checkInId) {
        if ((ciData[i][CI_STATUS_COL] || '').toString() === 'Checked Out') {
          return { success: false, message: "This check-in has already been checked out." };
        }
        ciRowIndex = i;
        ci = ciData[i];
        break;
      }
    }

    if (ciRowIndex === -1 && bookingsSheet && bookingsSheet.getLastRow() > 1) {
      const bData = bookingsSheet.getDataRange().getValues();
      for (let i = 1; i < bData.length; i++) {
        if ((bData[i][TICKET_ID_COL] || '').toString() === checkInId) {
          let bStatus = (bData[i][BOOKING_STATUS_COL] || '').toString();
          if (bStatus === 'Checked Out') return { success: false, message: "This booking has already been checked out." };
          ci = new Array(28).fill('');
          ci[CI_ID_COL]            = checkInId;
          ci[CI_LINKED_TICKET_COL] = checkInId;
          ci[CI_GUEST_NAME_COL]    = (bData[i][GUEST_NAME_COL] || '').toString();
          ci[CI_MOBILE_COL]        = (bData[i][PHONE_COL] || '').toString();
          ci[CI_EMAIL_COL]         = (bData[i][EMAIL_COL] || '').toString();
          ci[CI_ADDRESS_COL]       = (bData[i][CITY_COL] || '').toString();
          ci[CI_CHECKIN_DATE_COL]  = (bData[i][CHECK_IN_COL] || '').toString();
          ci[CI_CHECKIN_TIME_COL]  = (bData[i][CHECKIN_TIME_COL] || '14:00').toString();
          ci[CI_CHECKOUT_DATE_COL] = (bData[i][CHECK_OUT_COL] || '').toString();
          ci[CI_CHECKOUT_TIME_COL] = (bData[i][CHECKOUT_TIME_COL] || '12:00').toString();
          ci[CI_ROOM_NUMBERS_COL]  = (bData[i][BOOKING_ROOM_NO_COL] || '').toString();
          ci[CI_NUM_ROOMS_COL]     = parseInt(bData[i][NUM_ROOMS_COL]) || 1;
          ci[CI_PAX_COL]           = 1;
          ci[CI_ADVANCE_PAID_COL]  = parseFloat(bData[i][ADVANCE_PAID_COL]) || 0;
          ci[CI_FOOD_PLAN_COL]     = (bData[i][FOOD_PLAN_COL] || 'None').toString();
          ci[CI_GST_TYPE_COL]      = 'Excluding';
          ci[CI_FIX_RENT_COL]      = 'No';
          ci[CI_FIX_RENT_AMT_COL]  = 0;
          ci[CI_BILL_TO_COL]       = 'Individual';
          ci[CI_DISCOUNT_COL]      = parseFloat(bData[i][DISCOUNT_COL]) || 0;
          ci[CI_STATUS_COL]        = 'Active';
          ciRowIndex = -2;
          break;
        }
      }
    }
    if (!ci) return { success: false, message: "No check-in or booking record found." };

    let gstPercent = 5;
    let hotelName = 'MRI Hotel', hotelAddress = '', hotelPhone = '', hotelEmail = '', hotelTIN = '', hotelLogo = '', defaultCurrency = 'MVR';
    try {
      const setSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
      if (setSheet && setSheet.getLastRow() > 1) {
        const setRow = setSheet.getRange(2, 1, 1, setSheet.getLastColumn()).getValues()[0];
        hotelName = (setRow[SET_HOTEL_NAME_COL] || 'MRI Hotel').toString();
        hotelAddress = (setRow[SET_HOTEL_ADDRESS_COL] || '').toString();
        hotelPhone = (setRow[SET_HOTEL_PHONE_COL] || '').toString();
        hotelEmail = (setRow[SET_HOTEL_EMAIL_COL] || '').toString();
        hotelTIN = (setRow[SET_HOTEL_TIN_COL] || '').toString();
        hotelLogo = (setRow[SET_LOGO_URL_COL] || '').toString();
        defaultCurrency = (setRow[SET_DEFAULT_CURRENCY_COL] || 'MVR').toString();
        gstPercent = parseFloat(setRow[SET_GST_DEFAULT_COL]) || 5;
      }
    } catch (se) { Logger.log("Settings read error: " + se); }

    const guestName = (ci[CI_GUEST_NAME_COL] || '').toString();
    const companyName = (ci[CI_COMPANY_COL] || '').toString();
    const gstNumber = (ci[CI_GST_NUMBER_COL] || '').toString();
    const mobile = (ci[CI_MOBILE_COL] || '').toString();
    const email = (ci[CI_EMAIL_COL] || '').toString();
    const address = (ci[CI_ADDRESS_COL] || '').toString();
    const roomNumbers = (ci[CI_ROOM_NUMBERS_COL] || '').toString();
    const roomTypes = (ci[CI_ROOM_TYPES_COL] || '').toString();
    const pax = parseInt(ci[CI_PAX_COL]) || 1;
    const extraPerson = parseInt(ci[CI_EXTRA_PERSON_COL]) || 0;
    const advancePaid = parseFloat(ci[CI_ADVANCE_PAID_COL]) || 0;
    const foodPlan = (ci[CI_FOOD_PLAN_COL] || 'None').toString();
    const gstType = (ci[CI_GST_TYPE_COL] || 'Excluding').toString();
    const fixRoomRent = (ci[CI_FIX_RENT_COL] || 'No').toString();
    const fixRoomRentAmount = parseFloat(ci[CI_FIX_RENT_AMT_COL]) || 0;
    const billTo = (ci[CI_BILL_TO_COL] || 'Individual').toString();
    const discountPercent = parseFloat(ci[CI_DISCOUNT_COL]) || 0;

    const checkInDate = new Date(ci[CI_CHECKIN_DATE_COL]);
    const checkInTime = (ci[CI_CHECKIN_TIME_COL] || '14:00').toString();
    const actualCheckOutDate = checkoutData.checkOutDate ? new Date(checkoutData.checkOutDate) : new Date();
    const checkOutTime = checkoutData.checkOutTime || (ci[CI_CHECKOUT_TIME_COL] || '12:00').toString();

    let nights = daysBetween(checkInDate, actualCheckOutDate);
    if (nights < 1) nights = 1;

    let roomNosArr = roomNumbers.split(',').map(r => r.trim()).filter(r => r);
    let dailyRoomRate = 0;
    if (fixRoomRent === 'Yes' && fixRoomRentAmount > 0) {
      dailyRoomRate = fixRoomRentAmount;
    } else {
      const roomsData = roomsSheet.getDataRange().getValues();
      for (let r = 0; r < roomNosArr.length; r++) {
        for (let j = 1; j < roomsData.length; j++) {
          if ((roomsData[j][ROOM_NO_COL] || '').toString() === roomNosArr[r]) {
            dailyRoomRate += parseFloat(roomsData[j][ROOM_RATE_COL]) || 0;
            break;
          }
        }
      }
    }
    let totalRoomRent = dailyRoomRate * nights;

    let foodOrders = [];
    if (restSheet && restSheet.getLastRow() > 1) {
      const restData = restSheet.getDataRange().getValues();
      for (let i = 1; i < restData.length; i++) {
        if ((restData[i][REST_CHECKIN_ID_COL] || '').toString() === checkInId && (restData[i][REST_STATUS_COL] || '').toString() === 'Active') {
          foodOrders.push({
            orderDate: (restData[i][REST_ORDER_DATE_COL] || '').toString(),
            category: (restData[i][REST_CATEGORY_COL] || '').toString(),
            description: (restData[i][REST_DESC_COL] || '').toString(),
            amount: parseFloat(restData[i][REST_AMOUNT_COL]) || 0
          });
        }
      }
    }

    let totalFooding = 0;
    let totalExtraBed = 0;
    let totalOtherServices = 0;
    let categoryTotals = {};
    foodOrders.forEach(o => {
      categoryTotals[o.category] = (categoryTotals[o.category] || 0) + o.amount;
      if (o.category === 'FoodBeverage') totalFooding += o.amount;
      else if (o.category === 'ExtraBed') totalExtraBed += o.amount;
      else totalOtherServices += o.amount;
    });

    let subtotal = totalRoomRent + totalFooding + totalExtraBed + totalOtherServices;
    let discountAmount = subtotal * (discountPercent / 100);
    let afterDiscount = subtotal - discountAmount;

    let sgstPercent = gstPercent / 2;
    let cgstPercent = gstPercent / 2;
    let sgstAmount = 0, cgstAmount = 0;
    if (gstType === 'Excluding') {
      sgstAmount = afterDiscount * (sgstPercent / 100);
      cgstAmount = afterDiscount * (cgstPercent / 100);
    }

    let billAmount = afterDiscount + sgstAmount + cgstAmount;
    let netAmount = billAmount - advancePaid;
    if (netAmount < 0) netAmount = 0;

    let paymentMode = checkoutData.paymentMode || 'Cash';
    let amountPaid = parseFloat(checkoutData.amountPaid) || 0;
    let balance = netAmount - amountPaid;

    let billNumber = generateBillNumber();

    let dayByDay = [];
    let grandRunning = 0;
    for (let d = 0; d < nights; d++) {
      let dayDate = new Date(checkInDate);
      dayDate.setDate(dayDate.getDate() + d);
      let dateStr = dayDate.toISOString().split('T')[0];

      let dayRoom = dailyRoomRate;
      let dayCats = { ExtraBed: 0, FoodBeverage: 0, MiniBar: 0, EarlyClean: 0, Xerox: 0, Laundry: 0, Fax: 0, SPBUC: 0, Travels: 0, Misc: 0 };

      foodOrders.forEach(o => {
        let oDate = o.orderDate.split('T')[0];
        if (oDate === dateStr && dayCats.hasOwnProperty(o.category)) {
          dayCats[o.category] += o.amount;
        }
      });

      let dayTotal = dayRoom;
      Object.values(dayCats).forEach(v => dayTotal += v);
      grandRunning += dayTotal;

      dayByDay.push({
        date: dateStr, rooms: dayRoom, extraBed: dayCats.ExtraBed, foodBev: dayCats.FoodBeverage, miniBar: dayCats.MiniBar, earlyClean: dayCats.EarlyClean, xerox: dayCats.Xerox, laundry: dayCats.Laundry, fax: dayCats.Fax, spbuc: dayCats.SPBUC, travels: dayCats.Travels, misc: dayCats.Misc, dayTotal: dayTotal, grandTotal: grandRunning
      });
    }

    if (ciRowIndex >= 0 && ciSheet) {
      ciSheet.getRange(ciRowIndex + 1, CI_STATUS_COL + 1).setValue("Checked Out");
      ciSheet.getRange(ciRowIndex + 1, CI_CHECKOUT_DATE_COL + 1).setValue(actualCheckOutDate.toISOString());
      ciSheet.getRange(ciRowIndex + 1, CI_CHECKOUT_TIME_COL + 1).setValue(checkOutTime);
    }

    const roomsData = roomsSheet.getDataRange().getValues();
    for (let j = 1; j < roomsData.length; j++) {
      let rn = (roomsData[j][ROOM_NO_COL] || '').toString();
      if (roomNosArr.indexOf(rn) !== -1) {
        roomsSheet.getRange(j + 1, ROOM_STATUS_COL + 1).setValue("Available");
      }
    }

    if (ci[CI_LINKED_TICKET_COL]) {
      const bData = bookingsSheet.getDataRange().getValues();
      for (let i = 1; i < bData.length; i++) {
        if ((bData[i][TICKET_ID_COL] || '').toString() === ci[CI_LINKED_TICKET_COL].toString()) {
          bookingsSheet.getRange(i + 1, BOOKING_STATUS_COL + 1).setValue("Checked Out");
          break;
        }
      }
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Checkout completed successfully.",
      invoiceData: {
        billNumber, checkInId, hotelName, hotelAddress, hotelPhone, hotelEmail, hotelTIN, hotelLogo, currency: defaultCurrency, guestName, companyName, gstNumber, mobile, email, address, checkInDate: checkInDate.toISOString(), checkInTime, checkOutDate: actualCheckOutDate.toISOString(), checkOutTime, roomNumbers, roomTypes, numberOfRooms: roomNosArr.length, pax, extraPerson, foodPlan, billTo, nights, dailyRoomRate, totalRoomRent, totalFooding, totalExtraBed, totalOtherServices, categoryTotals, subtotal, discountPercent, discountAmount, gstType, sgstPercent, cgstPercent, sgstAmount, cgstAmount, billAmount, advancePaid, netAmount, paymentMode, amountPaid, balance, dayByDay
      }
    };
  } catch (e) {
    Logger.log("Error in processFullCheckout: " + e.toString());
    return { success: false, message: e.message };
  }
}
