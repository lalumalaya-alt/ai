/***************************************************
 * FINANCE MANAGEMENT
 ***************************************************/
function recalculateBalances() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) return { success: false, message: "Finance sheet not found." };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, message: "No records to recalculate." };

    let runningBalance = 0;
    let balanceArray = [];

    for (let i = 1; i < data.length; i++) {
      let type = (data[i][FIN_TYPE_COL] || "").toString();
      let amount = parseFloat(data[i][FIN_AMOUNT_COL]) || 0;
      if (type === "Income") {
        runningBalance += amount;
      } else if (type === "Expense") {
        runningBalance -= amount;
      }
      balanceArray.push([runningBalance]);
    }

    sheet.getRange(2, FIN_BALANCE_COL + 1, balanceArray.length, 1).setValues(balanceArray);
    SpreadsheetApp.flush();
    return { success: true, message: "Balances recalculated." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getAllFinanceRecords() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let records = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      records.push({
        rowIndex: i + 1,
        id: (row[FIN_ID_COL] || "").toString(),
        date: (row[FIN_DATE_COL] || "").toString(),
        type: (row[FIN_TYPE_COL] || "").toString(),
        description: (row[FIN_DESC_COL] || "").toString(),
        shopSource: (row[FIN_SHOP_COL] || "").toString(),
        amount: parseFloat(row[FIN_AMOUNT_COL]) || 0,
        balance: parseFloat(row[FIN_BALANCE_COL]) || 0,
        enteredBy: (row[FIN_ENTERED_BY_COL] || "").toString(),
        createdAt: (row[FIN_CREATED_AT_COL] || "").toString(),
        category: (row[FIN_CATEGORY_COL] || '').toString(),
        currency: (row[FIN_CURRENCY_COL] || 'MVR').toString(),
        linkedInvoiceId: (row[FIN_LINKED_INV_COL] || '').toString()
      });
    }
    return records;
  } catch (err) {
    return { error: err.message };
  }
}

function addFinanceRecord(recordData) {
  try {
    if (!recordData.date || !recordData.type || !recordData.description) {
      return { success: false, message: "Date, type, and description are required." };
    }
    if (recordData.type !== "Income" && recordData.type !== "Expense") {
      return { success: false, message: "Type must be 'Income' or 'Expense'." };
    }
    let amount = parseFloat(recordData.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: "Amount must be a positive number." };
    }

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: "Finance sheet not found. Please create it first." };
    }

    const id = generateFinanceId();
    const createdAt = new Date().toISOString();

    const data = sheet.getDataRange().getValues();
    let previousBalance = 0;
    if (data.length > 1) {
      previousBalance = parseFloat(data[data.length - 1][FIN_BALANCE_COL]) || 0;
    }
    let newBalance = recordData.type === "Income" ? previousBalance + amount : previousBalance - amount;

    sheet.appendRow([
      id,
      recordData.date,
      recordData.type,
      recordData.description.trim(),
      (recordData.shopSource || "").trim(),
      amount,
      newBalance,
      (recordData.enteredBy || "").trim(),
      createdAt,
      (recordData.category || '').trim(),
      (recordData.currency || 'MVR').trim(),
      (recordData.linkedInvoiceId || '').trim()
    ]);

    if (recordData.type === 'Expense') {
      const d = new Date(recordData.date);
      getBudgetForMonth(d.getMonth() + 1, d.getFullYear());
    }

    return { success: true, message: "Finance record added successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateFinanceRecord(rowIndex, recordData) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) return { success: false, message: "Finance sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    if (recordData.date !== undefined) sheet.getRange(rowIndex, FIN_DATE_COL + 1).setValue(recordData.date);
    if (recordData.type !== undefined) {
      if (recordData.type !== "Income" && recordData.type !== "Expense") {
        return { success: false, message: "Type must be 'Income' or 'Expense'." };
      }
      sheet.getRange(rowIndex, FIN_TYPE_COL + 1).setValue(recordData.type);
    }
    if (recordData.description !== undefined) sheet.getRange(rowIndex, FIN_DESC_COL + 1).setValue(recordData.description);
    if (recordData.shopSource !== undefined) sheet.getRange(rowIndex, FIN_SHOP_COL + 1).setValue(recordData.shopSource);
    if (recordData.amount !== undefined) {
      let amount = parseFloat(recordData.amount);
      if (isNaN(amount) || amount <= 0) {
        return { success: false, message: "Amount must be a positive number." };
      }
      sheet.getRange(rowIndex, FIN_AMOUNT_COL + 1).setValue(amount);
    }
    if (recordData.category !== undefined) sheet.getRange(rowIndex, FIN_CATEGORY_COL + 1).setValue(recordData.category);
    if (recordData.currency !== undefined) sheet.getRange(rowIndex, FIN_CURRENCY_COL + 1).setValue(recordData.currency);

    recalculateBalances();

    return { success: true, message: "Finance record updated successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteFinanceRecord(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) return { success: false, message: "Finance sheet not found." };
    if (rowIndex <= 1) {
      return { success: false, message: "Cannot delete header row." };
    }
    sheet.deleteRow(rowIndex);
    recalculateBalances();
    return { success: true, message: "Finance record deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getFinanceSummary() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet) return { totalIncome: 0, totalExpenses: 0, netBalance: 0 };
    const data = sheet.getDataRange().getValues();
    let totalIncome = 0;
    let totalExpenses = 0;

    for (let i = 1; i < data.length; i++) {
      let type = (data[i][FIN_TYPE_COL] || "").toString();
      let amount = parseFloat(data[i][FIN_AMOUNT_COL]) || 0;
      if (type === "Income") totalIncome += amount;
      else if (type === "Expense") totalExpenses += amount;
    }

    return { totalIncome, totalExpenses, netBalance: totalIncome - totalExpenses };
  } catch (err) {
    return { error: err.message };
  }
}

/***************************************************
 * BUDGET MANAGEMENT
 ***************************************************/
function getAllBudgets() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BUDGETS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    let records = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      records.push({
        rowIndex: i + 1,
        budgetId: (row[BDG_ID_COL] || '').toString(),
        month: parseInt(row[BDG_MONTH_COL]) || 0,
        year: parseInt(row[BDG_YEAR_COL]) || 0,
        budgetAmount: parseFloat(row[BDG_AMOUNT_COL]) || 0,
        totalSpent: parseFloat(row[BDG_SPENT_COL]) || 0,
        remaining: parseFloat(row[BDG_REMAINING_COL]) || 0,
        setBy: (row[BDG_SET_BY_COL] || '').toString(),
        createdAt: (row[BDG_CREATED_AT_COL] || '').toString(),
        updatedAt: (row[BDG_UPDATED_AT_COL] || '').toString()
      });
    }
    return records;
  } catch (err) {
    return { error: err.message };
  }
}

function setBudget(month, year, budgetAmount, user) {
  try {
    month = parseInt(month);
    year = parseInt(year);
    budgetAmount = parseFloat(budgetAmount);
    if (!month || !year || isNaN(budgetAmount) || budgetAmount < 0) {
      return { success: false, message: "Valid month, year, and budget amount are required." };
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(BUDGETS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Budgets sheet not found." };

    const now = new Date().toISOString();
    const budgetId = 'BDG-' + year + '-' + String(month).padStart(2, '0');
    const spent = calculateMonthlyExpenses(month, year);
    const remaining = budgetAmount - spent;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][BDG_MONTH_COL]) === month && parseInt(data[i][BDG_YEAR_COL]) === year) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          budgetId, month, year, budgetAmount, spent, remaining, user || '', data[i][BDG_CREATED_AT_COL], now
        ]]);
        return { success: true, message: "Budget updated for " + month + "/" + year + "!" };
      }
    }

    sheet.appendRow([budgetId, month, year, budgetAmount, spent, remaining, user || '', now, now]);
    return { success: true, message: "Budget set for " + month + "/" + year + "!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getBudgetForMonth(month, year) {
  try {
    month = parseInt(month);
    year = parseInt(year);
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(BUDGETS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][BDG_MONTH_COL]) === month && parseInt(data[i][BDG_YEAR_COL]) === year) {
        const spent = calculateMonthlyExpenses(month, year);
        const budgetAmount = parseFloat(data[i][BDG_AMOUNT_COL]) || 0;
        const remaining = budgetAmount - spent;

        sheet.getRange(i + 1, BDG_SPENT_COL + 1).setValue(spent);
        sheet.getRange(i + 1, BDG_REMAINING_COL + 1).setValue(remaining);
        sheet.getRange(i + 1, BDG_UPDATED_AT_COL + 1).setValue(new Date().toISOString());

        return {
          budgetId: (data[i][BDG_ID_COL] || '').toString(),
          month: month, year: year,
          budgetAmount: budgetAmount,
          totalSpent: spent,
          remaining: remaining
        };
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

function calculateMonthlyExpenses(month, year) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return 0;
    const data = sheet.getDataRange().getValues();
    let total = 0;
    for (let i = 1; i < data.length; i++) {
      const type = (data[i][FIN_TYPE_COL] || '').toString();
      if (type !== 'Expense') continue;
      const dateStr = (data[i][FIN_DATE_COL] || '').toString();
      if (!dateStr) continue;
      const d = new Date(dateStr);
      if ((d.getMonth() + 1) === month && d.getFullYear() === year) {
        total += parseFloat(data[i][FIN_AMOUNT_COL]) || 0;
      }
    }
    return Math.round(total * 100) / 100;
  } catch (err) {
    return 0;
  }
}

/***************************************************
 * CATEGORIES MANAGEMENT
 ***************************************************/
function getAllCategories() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CATEGORIES_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    let records = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      records.push({
        rowIndex: i + 1,
        categoryId: (row[CAT_ID_COL] || '').toString(),
        name: (row[CAT_NAME_COL] || '').toString(),
        type: (row[CAT_TYPE_COL] || '').toString(),
        isDefault: row[CAT_IS_DEFAULT_COL] === true || row[CAT_IS_DEFAULT_COL] === 'true',
        createdBy: (row[CAT_CREATED_BY_COL] || '').toString(),
        createdAt: (row[CAT_CREATED_AT_COL] || '').toString()
      });
    }
    return records;
  } catch (err) {
    return { error: err.message };
  }
}

function addCategory(name, type, user) {
  try {
    if (!name || !type) return { success: false, message: "Category name and type are required." };
    if (type !== 'Income' && type !== 'Expense') return { success: false, message: "Type must be 'Income' or 'Expense'." };

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CATEGORIES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Categories sheet not found." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][CAT_NAME_COL] || '').toString().toLowerCase() === name.toLowerCase() &&
          (data[i][CAT_TYPE_COL] || '').toString() === type) {
        return { success: false, message: "Category '" + name + "' already exists for " + type + "." };
      }
    }

    const id = 'CAT-' + new Date().getTime();
    sheet.appendRow([id, name.trim(), type, false, user || '', new Date().toISOString()]);
    return { success: true, message: "Category '" + name + "' added successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteCategory(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CATEGORIES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Categories sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Cannot delete header row." };

    const isDefault = sheet.getRange(rowIndex, CAT_IS_DEFAULT_COL + 1).getValue();
    if (isDefault === true || isDefault === 'true') {
      return { success: false, message: "Cannot delete default categories." };
    }

    sheet.deleteRow(rowIndex);
    return { success: true, message: "Category deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
