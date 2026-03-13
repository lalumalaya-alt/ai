/***************************************************
 * ROOM MANAGEMENT
 ***************************************************/
function getAllRooms() {
  try {
    const roomsSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    const data = roomsSheet.getDataRange().getValues();
    let rooms = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      rooms.push({
        rowIndex: i + 1,
        roomNo: row[ROOM_NO_COL],
        roomType: row[ROOM_TYPE_COL],
        roomRate: row[ROOM_RATE_COL],
        roomStatus: row[ROOM_STATUS_COL]
      });
    }
    return rooms;
  } catch (err) {
    return { error: err.message };
  }
}

function addRoom(roomNo, roomType, roomRate, roomStatus) {
  try {
    if (!roomNo) {
      return { success: false, message: "Room No is required." };
    }

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let existingRoomNo = (data[i][ROOM_NO_COL] || "").toString().trim().toLowerCase();
      if (existingRoomNo === roomNo.toString().toLowerCase()) {
        return { success: false, message: "A room with this number already exists." };
      }
    }

    sheet.appendRow([
      roomNo.trim(),
      (roomType || "").trim(),
      parseFloat(roomRate) || 0,
      (roomStatus || "Available").trim()
    ]);

    return { success: true, message: "Room added successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateRoom(rowIndex, roomNo, roomType, roomRate, roomStatus) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    if (rowIndex <= 1) {
      return { success: false, message: "Invalid row index." };
    }

    const currentValue = sheet.getRange(rowIndex, ROOM_NO_COL + 1).getValue();
    if (roomNo && roomNo.toString().trim().toLowerCase() !== (currentValue || "").toString().trim().toLowerCase()) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (i + 1 === rowIndex) continue;
        let existingNo = (data[i][ROOM_NO_COL] || "").toString().toLowerCase();
        if (existingNo === roomNo.toString().toLowerCase()) {
          return { success: false, message: "Another room with this number already exists." };
        }
      }
    }

    if (roomNo) {
      sheet.getRange(rowIndex, ROOM_NO_COL + 1).setValue(roomNo);
    }
    if (roomType !== undefined) {
      sheet.getRange(rowIndex, ROOM_TYPE_COL + 1).setValue(roomType);
    }
    if (roomRate !== undefined) {
      sheet.getRange(rowIndex, ROOM_RATE_COL + 1).setValue(parseFloat(roomRate) || 0);
    }
    if (roomStatus !== undefined) {
      sheet.getRange(rowIndex, ROOM_STATUS_COL + 1).setValue(roomStatus);
    }

    return { success: true, message: "Room updated successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteRoom(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    if (rowIndex <= 1) {
      return { success: false, message: "Cannot delete header row." };
    }
    sheet.deleteRow(rowIndex);
    return { success: true, message: "Room deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getRoomsForKanban() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    const rooms = roomsSheet.getDataRange().getValues();
    const bookings = bookingsSheet.getDataRange().getValues();

    const guestMap = {};
    for (let i = 1; i < bookings.length; i++) {
      const status = (bookings[i][BOOKING_STATUS_COL] || '').toString().toLowerCase();
      if (status === 'booked') {
        const roomNo = (bookings[i][BOOKING_ROOM_NO_COL] || '').toString();
        guestMap[roomNo] = {
          guestName: (bookings[i][GUEST_NAME_COL] || '').toString(),
          checkIn: bookings[i][CHECK_IN_COL] ? new Date(bookings[i][CHECK_IN_COL]).toISOString() : '',
          checkOut: bookings[i][CHECK_OUT_COL] ? new Date(bookings[i][CHECK_OUT_COL]).toISOString() : ''
        };
      }
    }

    let result = [];
    for (let i = 1; i < rooms.length; i++) {
      const roomNo = (rooms[i][ROOM_NO_COL] || '').toString();
      const guest = guestMap[roomNo] || {};
      result.push({
        roomNo: roomNo,
        roomType: (rooms[i][ROOM_TYPE_COL] || '').toString(),
        roomRate: parseFloat(rooms[i][ROOM_RATE_COL]) || 0,
        roomStatus: (rooms[i][ROOM_STATUS_COL] || '').toString(),
        guestName: guest.guestName || '',
        checkIn: guest.checkIn || '',
        checkOut: guest.checkOut || ''
      });
    }
    return result;
  } catch (err) {
    return { error: err.message };
  }
}

function checkReservedRooms() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const quotesSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
    if (!roomsSheet || !quotesSheet) return { success: false, message: "Sheet not found." };

    const roomsData = roomsSheet.getDataRange().getValues();
    const quotesData = quotesSheet.getDataRange().getValues();
    const now = new Date();
    let releasedCount = 0;

    for (let i = 1; i < roomsData.length; i++) {
      if ((roomsData[i][ROOM_STATUS_COL] || '').toString() === 'Reserved') {
        const roomNo = (roomsData[i][ROOM_NO_COL] || '').toString();
        let shouldRelease = true;

        for (let q = 1; q < quotesData.length; q++) {
          const qStatus = (quotesData[q][QUOTE_STATUS_COL] || '').toString();
          const qCreated = quotesData[q][QUOTE_CREATED_COL] ? new Date(quotesData[q][QUOTE_CREATED_COL]) : null;
          const qConverted = (quotesData[q][QUOTE_CONVERTED_COL] || '').toString();

          if (qConverted) continue;
          if (qStatus === 'Expired' || qStatus === 'Converted') continue;

          try {
            const items = JSON.parse((quotesData[q][QUOTE_ITEMS_COL] || '[]').toString());
            const hasRoom = items.some(it => it.type === 'room' && it.reservedRoomNo === roomNo);
            if (hasRoom && qCreated) {
              const hoursSince = (now - qCreated) / (1000 * 60 * 60);
              if (hoursSince < 24) {
                shouldRelease = false;
                break;
              }
            }
          } catch (e) { /* ignore parse errors */ }
        }

        if (shouldRelease) {
          roomsSheet.getRange(i + 1, ROOM_STATUS_COL + 1).setValue("Available");
          releasedCount++;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, message: releasedCount + " room(s) released from reservation.", releasedCount: releasedCount };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getAvailableRoomNumbers() {
  try {
    const roomsSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    const data = roomsSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    data.shift();
    let availableRooms = [];
    data.forEach(row => {
      let status = (row[ROOM_STATUS_COL] || "").toString().toLowerCase();
      if (status === "available") {
        availableRooms.push((row[ROOM_NO_COL] || "").toString());
      }
    });
    return availableRooms;
  } catch (e) {
    Logger.log(`Error in getAvailableRoomNumbers: ${e.toString()}`);
    return [];
  }
}
/***************************************************
 * BOOKINGS MANAGEMENT
 ***************************************************/
function bookRoom(bookingDetails) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    const roomsData = roomsSheet.getDataRange().getValues();

    let roomNosArr = [];
    if (bookingDetails.roomNos) {
      roomNosArr = bookingDetails.roomNos.split(',').map(r => r.trim()).filter(r => r);
    } else if (bookingDetails.roomNo) {
      roomNosArr = [bookingDetails.roomNo.toString().trim()];
    }
    if (roomNosArr.length === 0) {
      return { success: false, message: "No rooms selected." };
    }

    let totalRoomRate = 0;
    let roomRowIndices = [];
    for (let r = 0; r < roomNosArr.length; r++) {
      let found = false;
      for (let i = 1; i < roomsData.length; i++) {
        let rowRoomNo = (roomsData[i][ROOM_NO_COL] || "").toString();
        if (rowRoomNo === roomNosArr[r]) {
          let status = (roomsData[i][ROOM_STATUS_COL] || "").toString().toLowerCase();
          if (status !== 'available' && status !== 'reserved') {
            return { success: false, message: `Room ${roomNosArr[r]} is not available.` };
          }
          totalRoomRate += parseFloat(roomsData[i][ROOM_RATE_COL]) || 0;
          roomRowIndices.push(i);
          found = true;
          break;
        }
      }
      if (!found) {
        return { success: false, message: `Room ${roomNosArr[r]} not found.` };
      }
    }

    const ticketId = generateTicketId();
    const checkInDate = new Date(bookingDetails.checkIn);
    const checkOutDate = new Date(bookingDetails.checkOut);
    const checkInTime = bookingDetails.checkInTime || "14:00";
    const checkOutTime = bookingDetails.checkOutTime || "12:00";
    const foodPlan = bookingDetails.foodPlan || "None";
    const advancePaid = parseFloat(bookingDetails.advancePaid || "0") || 0;

    let discount = parseFloat(bookingDetails.discount || "0") || 0;
    let tax = parseFloat(bookingDetails.tax || "0") || 0;
    let paymentMethod = bookingDetails.paymentMethod || "Cash";

    let nights = daysBetween(checkInDate, checkOutDate);
    if (nights < 1) nights = 1;
    let baseAmount = totalRoomRate * nights;
    let finalAmount = baseAmount - discount + tax;

    let roomNosStr = roomNosArr.join(',');
    let paymentStatus = advancePaid >= finalAmount ? "Paid" : advancePaid > 0 ? "Partial" : "Unpaid";

    bookingsSheet.appendRow([
      ticketId,
      roomNosStr,
      bookingDetails.guestName,
      bookingDetails.phone,
      bookingDetails.email,
      bookingDetails.city || '',
      bookingDetails.maritalStatus || '',
      bookingDetails.occupancyType || '',
      bookingDetails.familyDetails || '',
      checkInDate.toISOString(),
      checkOutDate.toISOString(),
      "Booked",
      totalRoomRate,
      discount,
      tax,
      paymentMethod,
      finalAmount,
      paymentStatus,
      advancePaid,
      checkInTime,
      checkOutTime,
      foodPlan,
      advancePaid,
      roomNosArr.length,
      ""
    ]);

    for (let ri = 0; ri < roomRowIndices.length; ri++) {
      roomsSheet.getRange(roomRowIndices[ri] + 1, ROOM_STATUS_COL + 1).setValue("Booked");
    }

    let autoGeneratedPass = "guest" + new Date().getTime().toString().slice(-3);
    if (bookingDetails.email) {
      createUserIfNotExists(bookingDetails.email, autoGeneratedPass);
    }

    SpreadsheetApp.flush();

    try {
      if (bookingDetails.email) {
        let subject = `Room Booking Confirmation - Ticket ${ticketId}`;
        let body = `Hello ${bookingDetails.guestName},\n\nThank you for booking Room(s) #${roomNosStr}.\nCheck-in: ${checkInDate.toISOString()} at ${checkInTime}\nCheck-out: ${checkOutDate.toISOString()} at ${checkOutTime}\nFood Plan: ${foodPlan}\nAdvance Paid: ${advancePaid}\n\nTicket ID: ${ticketId}\n\nWe look forward to your stay!\n- MRI Hotel`;
        MailApp.sendEmail({ to: bookingDetails.email, subject, body });
      }
    } catch (emailErr) {
      Logger.log(`Email failed for booking ${ticketId}: ${emailErr.message}`);
    }

    return {
      success: true,
      message: `Room(s) ${roomNosStr} booked successfully. Ticket ID: ${ticketId}`,
      ticketId
    };
  } catch (e) {
    Logger.log(`Error in bookRoom: ${e.toString()}`);
    return { success: false, message: `An error occurred: ${e.message}` };
  }
}

function updateBooking(rowIndex, bookingData) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Bookings sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    const existingTicket = sheet.getRange(rowIndex, TICKET_ID_COL + 1).getValue();
    const existingRoomNo = sheet.getRange(rowIndex, BOOKING_ROOM_NO_COL + 1).getValue();
    const existingRate = parseFloat(sheet.getRange(rowIndex, ROOM_RATE_BOOK_COL + 1).getValue()) || 0;
    const existingStatus = (sheet.getRange(rowIndex, BOOKING_STATUS_COL + 1).getValue() || '').toString();
    const existingPaymentStatus = (sheet.getRange(rowIndex, PAYMENT_STATUS_COL + 1).getValue() || 'Unpaid').toString();
    const existingAmountPaid = parseFloat(sheet.getRange(rowIndex, AMOUNT_PAID_COL + 1).getValue()) || 0;

    if (existingStatus.toLowerCase() === 'checked out') {
      return { success: false, message: "Cannot edit a checked-out booking." };
    }

    const checkInDate = new Date(bookingData.checkIn);
    const checkOutDate = new Date(bookingData.checkOut);
    if (isNaN(checkInDate.getTime()) || isNaN(checkOutDate.getTime())) {
      return { success: false, message: "Invalid dates provided." };
    }
    if (checkOutDate <= checkInDate) {
      return { success: false, message: "Check-out must be after check-in." };
    }

    let nights = daysBetween(checkInDate, checkOutDate);
    if (nights < 1) nights = 1;
    const discount = parseFloat(bookingData.discount || 0) || 0;
    const tax = parseFloat(bookingData.tax || 0) || 0;
    const baseAmount = existingRate * nights;
    const finalAmount = baseAmount - discount + tax;

    const existingCheckInTime = (sheet.getRange(rowIndex, CHECKIN_TIME_COL + 1).getValue() || '14:00').toString();
    const existingCheckOutTime = (sheet.getRange(rowIndex, CHECKOUT_TIME_COL + 1).getValue() || '12:00').toString();
    const existingFoodPlan = (sheet.getRange(rowIndex, FOOD_PLAN_COL + 1).getValue() || 'None').toString();
    const existingAdvancePaid = parseFloat(sheet.getRange(rowIndex, ADVANCE_PAID_COL + 1).getValue()) || 0;
    const existingNumRooms = parseFloat(sheet.getRange(rowIndex, NUM_ROOMS_COL + 1).getValue()) || 1;
    const existingLinkedCheckIn = (sheet.getRange(rowIndex, LINKED_CHECKIN_COL + 1).getValue() || '').toString();

    const row = [
      existingTicket,
      existingRoomNo,
      (bookingData.guestName || '').trim(),
      (bookingData.phone || '').trim(),
      (bookingData.email || '').trim(),
      (bookingData.city || '').trim(),
      bookingData.maritalStatus || 'Single',
      bookingData.occupancyType || 'Single',
      (bookingData.familyDetails || '').trim(),
      checkInDate.toISOString(),
      checkOutDate.toISOString(),
      existingStatus,
      existingRate,
      discount,
      tax,
      bookingData.paymentMethod || 'Cash',
      finalAmount,
      existingPaymentStatus,
      existingAmountPaid,
      bookingData.checkInTime || existingCheckInTime,
      bookingData.checkOutTime || existingCheckOutTime,
      bookingData.foodPlan || existingFoodPlan,
      bookingData.advancePaid !== undefined ? parseFloat(bookingData.advancePaid) || 0 : existingAdvancePaid,
      existingNumRooms,
      existingLinkedCheckIn
    ];

    sheet.getRange(rowIndex, 1, 1, 25).setValues([row]);
    SpreadsheetApp.flush();

    return { success: true, message: "Booking updated successfully." };
  } catch (err) {
    Logger.log("Error in updateBooking: " + err.toString());
    return { success: false, message: err.message };
  }
}

function deleteBooking(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Bookings sheet not found." };
    if (rowIndex <= 1 || rowIndex > sheet.getLastRow()) return { success: false, message: "Invalid row index." };

    const status = (sheet.getRange(rowIndex, BOOKING_STATUS_COL + 1).getValue() || '').toString().toLowerCase();
    if (status === 'checked out') {
      return { success: false, message: "Cannot delete a checked-out booking." };
    }

    if (status === 'booked') {
      const roomNoStr = (sheet.getRange(rowIndex, BOOKING_ROOM_NO_COL + 1).getValue() || '').toString();
      if (roomNoStr) {
        const roomNosArr = roomNoStr.split(',').map(r => r.trim()).filter(r => r);
        const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
        if (roomsSheet) {
          const roomsData = roomsSheet.getDataRange().getValues();
          for (let j = 1; j < roomsData.length; j++) {
            let rn = (roomsData[j][ROOM_NO_COL] || '').toString();
            if (roomNosArr.indexOf(rn) !== -1) {
              roomsSheet.getRange(j + 1, ROOM_STATUS_COL + 1).setValue("Available");
            }
          }
        }
      }
    }

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true, message: "Booking deleted successfully." };
  } catch (err) {
    Logger.log("Error in deleteBooking: " + err.toString());
    return { success: false, message: err.message };
  }
}

function getAllBookings() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let bookings = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      bookings.push({
        rowIndex: i + 1,
        ticketId: (row[TICKET_ID_COL] || "").toString(),
        roomNo: (row[BOOKING_ROOM_NO_COL] || "").toString(),
        guestName: (row[GUEST_NAME_COL] || "").toString(),
        phone: (row[PHONE_COL] || "").toString(),
        email: (row[EMAIL_COL] || "").toString(),
        city: (row[CITY_COL] || "").toString(),
        maritalStatus: (row[MARITAL_STATUS_COL] || "").toString(),
        occupancyType: (row[OCCUPANCY_TYPE_COL] || "").toString(),
        familyDetails: (row[FAMILY_DETAILS_COL] || "").toString(),
        checkIn: row[CHECK_IN_COL] ? new Date(row[CHECK_IN_COL]).toISOString() : "",
        checkOut: row[CHECK_OUT_COL] ? new Date(row[CHECK_OUT_COL]).toISOString() : "",
        status: (row[BOOKING_STATUS_COL] || "").toString(),
        roomRate: parseFloat(row[ROOM_RATE_BOOK_COL]) || 0,
        discount: parseFloat(row[DISCOUNT_COL]) || 0,
        tax: parseFloat(row[TAX_COL]) || 0,
        paymentMethod: (row[PAYMENT_METHOD_COL] || "").toString(),
        totalAmount: parseFloat(row[TOTAL_AMOUNT_COL]) || 0,
        paymentStatus: (row[PAYMENT_STATUS_COL] || "Unpaid").toString(),
        amountPaid: parseFloat(row[AMOUNT_PAID_COL]) || 0,
        checkInTime: (row[CHECKIN_TIME_COL] || "14:00").toString(),
        checkOutTime: (row[CHECKOUT_TIME_COL] || "12:00").toString(),
        foodPlan: (row[FOOD_PLAN_COL] || "None").toString(),
        advancePaid: parseFloat(row[ADVANCE_PAID_COL]) || 0,
        numberOfRooms: parseInt(row[NUM_ROOMS_COL]) || 1,
        linkedCheckInId: (row[LINKED_CHECKIN_COL] || "").toString()
      });
    }
    return bookings;
  } catch (err) {
    return { error: err.message };
  }
}

function getCurrentBookings() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BOOKINGS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    let allBookings = data.map(row => {
      let booking = {};
      headers.forEach((header, idx) => {
        booking[header.trim().replace(/\s+/g, '')] = row[idx];
      });
      return booking;
    });

    let active = allBookings.filter(b => {
      return b.Status && b.Status.toString().toLowerCase() === 'booked';
    });

    active.forEach(b => {
      if (b.CheckIn) b.CheckIn = new Date(b.CheckIn).toISOString();
      if (b.CheckOut) b.CheckOut = new Date(b.CheckOut).toISOString();
    });

    return active;
  } catch (e) {
    Logger.log(`Error in getCurrentBookings: ${e.toString()}`);
    return { error: e.message };
  }
}

function getBookingByTicketId(ticketId) {
  try {
    const bookings = getAllBookings();
    if (bookings.error) return null;
    for (let i = 0; i < bookings.length; i++) {
      if (bookings[i].ticketId === ticketId) return bookings[i];
    }
    return null;
  } catch (e) {
    Logger.log("Error in getBookingByTicketId: " + e.toString());
    return null;
  }
}

function searchBookingsByGuestName(query) {
  try {
    const bookings = getAllBookings();
    if (bookings.error) return [];
    let q = (query || '').toLowerCase().trim();
    if (!q) return [];
    return bookings.filter(b => {
      return b.status.toLowerCase() === 'booked' && b.guestName.toLowerCase().indexOf(q) !== -1;
    });
  } catch (e) {
    Logger.log("Error in searchBookingsByGuestName: " + e.toString());
    return [];
  }
}
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
/***************************************************
 * RESTAURANT FUNCTIONS
 ***************************************************/
function addFoodOrder(orderData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(RESTAURANT_SHEET_NAME);
    if (!sheet) return { success: false, message: "Restaurant sheet not found. Run Setup Demo Data." };

    const orderId = generateOrderId();
    const now = new Date().toISOString();

    sheet.appendRow([
      orderId,
      orderData.roomNo || '',
      orderData.checkInId || '',
      orderData.orderDate || now.split('T')[0],
      orderData.category || 'FoodBeverage',
      orderData.description || '',
      parseFloat(orderData.amount) || 0,
      'Active',
      now
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "Order added successfully.", orderId };
  } catch (e) {
    Logger.log("Error in addFoodOrder: " + e.toString());
    return { success: false, message: e.message };
  }
}

function getAllFoodOrders() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(RESTAURANT_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getDataRange().getValues();
    let orders = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      orders.push({
        rowIndex: i + 1,
        orderId: (row[REST_ORDER_ID_COL] || '').toString(),
        roomNo: (row[REST_ROOM_NO_COL] || '').toString(),
        checkInId: (row[REST_CHECKIN_ID_COL] || '').toString(),
        orderDate: (row[REST_ORDER_DATE_COL] || '').toString(),
        category: (row[REST_CATEGORY_COL] || '').toString(),
        description: (row[REST_DESC_COL] || '').toString(),
        amount: parseFloat(row[REST_AMOUNT_COL]) || 0,
        status: (row[REST_STATUS_COL] || 'Active').toString(),
        createdAt: (row[REST_CREATED_AT_COL] || '').toString()
      });
    }
    return orders;
  } catch (e) {
    Logger.log("Error in getAllFoodOrders: " + e.toString());
    return { error: e.message };
  }
}

function getFoodOrdersByCheckIn(checkInId) {
  try {
    const all = getAllFoodOrders();
    if (all.error) return [];
    return all.filter(o => o.checkInId === checkInId && o.status === 'Active');
  } catch (e) {
    return [];
  }
}

function getFoodOrdersByRoom(roomNo) {
  try {
    const all = getAllFoodOrders();
    if (all.error) return [];
    return all.filter(o => o.roomNo === roomNo.toString() && o.status === 'Active');
  } catch (e) {
    return [];
  }
}

function updateFoodOrder(rowIndex, orderData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(RESTAURANT_SHEET_NAME);
    if (!sheet) return { success: false, message: "Restaurant sheet not found." };

    const existingId = sheet.getRange(rowIndex, REST_ORDER_ID_COL + 1).getValue();
    const existingCreated = sheet.getRange(rowIndex, REST_CREATED_AT_COL + 1).getValue();

    const row = [
      existingId,
      orderData.roomNo || '',
      orderData.checkInId || '',
      orderData.orderDate || '',
      orderData.category || 'FoodBeverage',
      orderData.description || '',
      parseFloat(orderData.amount) || 0,
      orderData.status || 'Active',
      existingCreated
    ];
    sheet.getRange(rowIndex, 1, 1, 9).setValues([row]);
    SpreadsheetApp.flush();
    return { success: true, message: "Order updated successfully." };
  } catch (e) {
    Logger.log("Error in updateFoodOrder: " + e.toString());
    return { success: false, message: e.message };
  }
}

function deleteFoodOrder(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(RESTAURANT_SHEET_NAME);
    if (!sheet) return { success: false, message: "Restaurant sheet not found." };
    if (rowIndex <= 1 || rowIndex > sheet.getLastRow()) return { success: false, message: "Invalid row." };
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    return { success: true, message: "Order deleted successfully." };
  } catch (e) {
    Logger.log("Error in deleteFoodOrder: " + e.toString());
    return { success: false, message: e.message };
  }
}

function getActiveCheckInRooms() {
  try {
    const checkIns = getAllCheckIns();
    if (checkIns.error) return [];
    let rooms = [];
    checkIns.forEach(ci => {
      if (ci.status === 'Active') {
        let roomNos = ci.roomNumbers.split(',').map(r => r.trim()).filter(r => r);
        roomNos.forEach(rn => {
          rooms.push({ roomNo: rn, checkInId: ci.checkInId, guestName: ci.guestName });
        });
      }
    });
    return rooms;
  } catch (e) {
    return [];
  }
}
