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
