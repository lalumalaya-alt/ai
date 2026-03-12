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
