/***************************************************
 * DASHBOARD
 ***************************************************/
function getDashboardData() {
  try {
    const roomsSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
    const roomsData = roomsSheet.getDataRange().getValues();
    roomsData.shift();

    let totalRooms = roomsData.length;
    let bookedCount = 0;
    let availableRoomsList = [];
    let bookedRoomsList = [];
    let allRoomsDetails = [];

    roomsData.forEach(row => {
      let roomNo = (row[ROOM_NO_COL] || "").toString();
      let type   = (row[ROOM_TYPE_COL] || "").toString();
      let status = (row[ROOM_STATUS_COL] || "").toString();
      allRoomsDetails.push({ roomNo, type, status });
      if (status.toLowerCase() === "booked") {
        bookedCount++;
        bookedRoomsList.push(roomNo);
      } else {
        availableRoomsList.push(roomNo);
      }
    });

    let maintenanceCount = roomsData.filter(r => (r[ROOM_STATUS_COL] || "").toString().toLowerCase() === "maintenance").length;
    let reservedCount = roomsData.filter(r => (r[ROOM_STATUS_COL] || "").toString().toLowerCase() === "reserved").length;
    let availableCount = totalRooms - bookedCount - maintenanceCount - reservedCount;

    let roomTypeMap = {};
    roomsData.forEach(row => {
      let t = (row[ROOM_TYPE_COL] || "Other").toString();
      roomTypeMap[t] = (roomTypeMap[t] || 0) + 1;
    });

    let financeSummary = { totalIncome: 0, totalExpenses: 0, netBalance: 0 };
    let expenseCategories = {};
    let incomeCategories = {};
    try {
      const finSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
      if (finSheet) {
        const finData = finSheet.getDataRange().getValues();
        for (let i = 1; i < finData.length; i++) {
          let type = (finData[i][FIN_TYPE_COL] || "").toString();
          let amount = parseFloat(finData[i][FIN_AMOUNT_COL]) || 0;
          let category = (finData[i][FIN_CATEGORY_COL] || "Uncategorized").toString();
          if (type === "Income") {
            financeSummary.totalIncome += amount;
            incomeCategories[category] = (incomeCategories[category] || 0) + amount;
          } else if (type === "Expense") {
            financeSummary.totalExpenses += amount;
            expenseCategories[category] = (expenseCategories[category] || 0) + amount;
          }
        }
        financeSummary.netBalance = financeSummary.totalIncome - financeSummary.totalExpenses;
      }
    } catch (finErr) {
      Logger.log("Could not load finance data: " + finErr);
    }

    let quoteStats = { totalQuotes: 0, draftQuotes: 0, sentQuotes: 0, acceptedQuotes: 0, expiredQuotes: 0 };
    try {
      const quoteSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
      if (quoteSheet) {
        const quoteData = quoteSheet.getDataRange().getValues();
        quoteStats.totalQuotes = Math.max(0, quoteData.length - 1);
        for (let i = 1; i < quoteData.length; i++) {
          let status = (quoteData[i][QUOTE_STATUS_COL] || "").toString();
          if (status === "Draft") quoteStats.draftQuotes++;
          else if (status === "Sent") quoteStats.sentQuotes++;
          else if (status === "Accepted") quoteStats.acceptedQuotes++;
          else if (status === "Expired") quoteStats.expiredQuotes++;
        }
      }
    } catch (quoteErr) {
      Logger.log("Could not load quote data: " + quoteErr);
    }

    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    let monthlyBookings = {};
    let monthlyRevenue = {};
    let monthlyIncome = {};
    let monthlyExpense = {};
    const now = new Date();
    for (let m = 5; m >= 0; m--) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const key = monthNames[d.getMonth()] + ' ' + d.getFullYear();
      monthlyBookings[key] = 0;
      monthlyRevenue[key] = 0;
      monthlyIncome[key] = 0;
      monthlyExpense[key] = 0;
    }

    let bookingRevenue = { totalRevenue: 0, checkedOutCount: 0, activeBookingCount: 0, totalBookings: 0 };
    let recentBookings = [];
    try {
      const bookSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BOOKINGS_SHEET_NAME);
      const bookData = bookSheet.getDataRange().getValues();

      for (let i = 1; i < bookData.length; i++) {
        let bStatus = (bookData[i][BOOKING_STATUS_COL] || "").toString().toLowerCase();
        let bAmount = parseFloat(bookData[i][TOTAL_AMOUNT_COL]) || 0;
        let ciDate = bookData[i][CHECK_IN_COL] ? new Date(bookData[i][CHECK_IN_COL]) : null;
        if (bStatus === "checked out" || bStatus === "completed") {
          bookingRevenue.totalRevenue += bAmount;
          bookingRevenue.checkedOutCount++;
        } else if (bStatus === "booked") {
          bookingRevenue.activeBookingCount++;
        }
        bookingRevenue.totalBookings++;
        if (ciDate) {
          const mKey = monthNames[ciDate.getMonth()] + ' ' + ciDate.getFullYear();
          if (monthlyBookings.hasOwnProperty(mKey)) {
            monthlyBookings[mKey]++;
            monthlyRevenue[mKey] += bAmount;
          }
        }
        recentBookings.push({
          ticketId: (bookData[i][TICKET_ID_COL] || '').toString(),
          roomNo: (bookData[i][BOOKING_ROOM_NO_COL] || '').toString(),
          guestName: (bookData[i][GUEST_NAME_COL] || '').toString(),
          checkIn: ciDate ? ciDate.toISOString() : '',
          status: (bookData[i][BOOKING_STATUS_COL] || '').toString(),
          totalAmount: bAmount
        });
      }
      recentBookings.reverse();
      recentBookings = recentBookings.slice(0, 8);

      try {
        const finSheet2 = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
        if (finSheet2) {
          const finData2 = finSheet2.getDataRange().getValues();
          for (let i = 1; i < finData2.length; i++) {
            let fDate = finData2[i][FIN_DATE_COL] ? new Date(finData2[i][FIN_DATE_COL]) : null;
            let fType = (finData2[i][FIN_TYPE_COL] || "").toString();
            let fAmt = parseFloat(finData2[i][FIN_AMOUNT_COL]) || 0;
            if (fDate) {
              const mKey = monthNames[fDate.getMonth()] + ' ' + fDate.getFullYear();
              if (monthlyIncome.hasOwnProperty(mKey)) {
                if (fType === "Income") monthlyIncome[mKey] += fAmt;
                else if (fType === "Expense") monthlyExpense[mKey] += fAmt;
              }
            }
          }
        }
      } catch (e2) { Logger.log("Monthly finance error: " + e2); }

    } catch (bookErr) {
      Logger.log("Could not load booking revenue data: " + bookErr);
    }

    let invoiceStats = { totalInvoices: 0, draftInvoices: 0, sentInvoices: 0, paidInvoices: 0, overdueInvoices: 0, cancelledInvoices: 0, invoiceRevenue: 0 };
    try {
      const invSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
      if (invSheet && invSheet.getLastRow() > 1) {
        const invData = invSheet.getDataRange().getValues();
        invoiceStats.totalInvoices = Math.max(0, invData.length - 1);
        for (let i = 1; i < invData.length; i++) {
          let status = (invData[i][INV_STATUS_COL] || '').toString();
          let total = parseFloat(invData[i][INV_TOTAL_COL]) || 0;
          if (status === 'Draft') invoiceStats.draftInvoices++;
          else if (status === 'Sent') invoiceStats.sentInvoices++;
          else if (status === 'Paid') { invoiceStats.paidInvoices++; invoiceStats.invoiceRevenue += total; }
          else if (status === 'Overdue') invoiceStats.overdueInvoices++;
          else if (status === 'Cancelled') invoiceStats.cancelledInvoices++;
        }
      }
    } catch (invErr) { Logger.log("Could not load invoice data: " + invErr); }

    let currentBudget = null;
    try {
      currentBudget = getBudgetForMonth(now.getMonth() + 1, now.getFullYear());
    } catch (bdgErr) { Logger.log("Could not load budget: " + bdgErr); }

    let settingsDefaultCurrency = 'MVR';
    try {
      const setSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SETTINGS_SHEET_NAME);
      if (setSheet && setSheet.getLastRow() > 1) {
        settingsDefaultCurrency = (setSheet.getRange(2, SET_DEFAULT_CURRENCY_COL + 1).getValue() || 'MVR').toString();
      }
    } catch (setErr) { Logger.log("Could not load settings currency: " + setErr); }

    return {
      totalRooms,
      bookedRooms: bookedCount,
      availableRooms: availableCount,
      maintenanceRooms: maintenanceCount,
      reservedRooms: reservedCount,
      availableRoomNumbers: availableRoomsList,
      bookedRoomNumbers: bookedRoomsList,
      allRoomsDetails,
      roomTypeBreakdown: roomTypeMap,
      financeSummary,
      expenseCategories,
      incomeCategories,
      quoteStats,
      bookingRevenue,
      recentBookings,
      invoiceStats,
      currentBudget,
      defaultCurrency: settingsDefaultCurrency,
      monthlyBookings: monthlyBookings || {},
      monthlyRevenue: monthlyRevenue || {},
      monthlyIncome: monthlyIncome || {},
      monthlyExpense: monthlyExpense || {}
    };
  } catch (e) {
    Logger.log(`Error in getDashboardData: ${e.toString()}`);
    return { error: e.message };
  }
}

function getMonthlyReport(month, year, reportType) {
  try {
    month = parseInt(month);
    year = parseInt(year);
    if (!month || !year) return { success: false, message: "Month and year are required." };

    const finSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!finSheet || finSheet.getLastRow() <= 1) {
      return { success: true, data: { records: [], categoryTotals: {}, totalIncome: 0, totalExpenses: 0, net: 0, budget: null } };
    }

    const finData = finSheet.getDataRange().getValues();
    let records = [];
    let categoryTotals = {};
    let totalIncome = 0;
    let totalExpenses = 0;

    for (let i = 1; i < finData.length; i++) {
      const dateStr = (finData[i][FIN_DATE_COL] || '').toString();
      if (!dateStr) continue;
      const d = new Date(dateStr);
      if ((d.getMonth() + 1) !== month || d.getFullYear() !== year) continue;

      const type = (finData[i][FIN_TYPE_COL] || '').toString();
      const amount = parseFloat(finData[i][FIN_AMOUNT_COL]) || 0;
      const category = (finData[i][FIN_CATEGORY_COL] || 'Uncategorized').toString();

      if (reportType === 'income' && type !== 'Income') continue;
      if (reportType === 'expense' && type !== 'Expense') continue;

      if (type === 'Income') totalIncome += amount;
      if (type === 'Expense') totalExpenses += amount;

      const catKey = type + ':' + category;
      if (!categoryTotals[catKey]) categoryTotals[catKey] = { category: category, type: type, total: 0 };
      categoryTotals[catKey].total += amount;

      records.push({
        id: (finData[i][FIN_ID_COL] || '').toString(),
        date: dateStr,
        type: type,
        description: (finData[i][FIN_DESC_COL] || '').toString(),
        shopSource: (finData[i][FIN_SHOP_COL] || '').toString(),
        amount: amount,
        category: category,
        currency: (finData[i][FIN_CURRENCY_COL] || 'MVR').toString(),
        enteredBy: (finData[i][FIN_ENTERED_BY_COL] || '').toString()
      });
    }

    const budget = getBudgetForMonth(month, year);

    return {
      success: true,
      data: {
        records: records,
        categoryTotals: Object.values(categoryTotals),
        totalIncome: Math.round(totalIncome * 100) / 100,
        totalExpenses: Math.round(totalExpenses * 100) / 100,
        net: Math.round((totalIncome - totalExpenses) * 100) / 100,
        budget: budget
      }
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
