/***************************************************
 * USERS MANAGEMENT
 ***************************************************/
function getAllUsers() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data = sheet.getDataRange().getValues();
    let users = [];

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      users.push({
        rowIndex: i + 1,
        username: (row[LOGIN_USERNAME_COL] || "").toString(),
        role: (row[LOGIN_ROLE_COL] || "").toString()
      });
    }

    return users;
  } catch (err) {
    return { error: err.message };
  }
}

function addUser(username, password, role) {
  try {
    if (!username || !password) {
      return { success: false, message: "Username and password are required." };
    }

    username = username.toString().trim().toLowerCase();

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    if (!sheet) return { success: false, message: "Login sheet not found." };

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase();
      if (storedUser === username) {
        return { success: false, message: "User already exists." };
      }
    }

    sheet.appendRow([username, password, role || "user", "", ""]);

    return { success: true, message: "User added successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateUser(rowIndex, newPassword, newRole) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    if (!sheet) return { success: false, message: "Login sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    if (newPassword) {
      if (newPassword.length < 4) return { success: false, message: "Password must be at least 4 characters." };
      sheet.getRange(rowIndex, LOGIN_PASSWORD_COL + 1).setValue(newPassword);
    }
    if (newRole) {
      sheet.getRange(rowIndex, LOGIN_ROLE_COL + 1).setValue(newRole);
    }

    return { success: true, message: "User updated successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteUser(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    if (!sheet) return { success: false, message: "Login sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Cannot delete header row." };

    sheet.deleteRow(rowIndex);

    return { success: true, message: "User deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/***************************************************
 * CUSTOMERS MANAGEMENT
 ***************************************************/
function getAllCustomers() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CUSTOMERS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data = sheet.getDataRange().getValues();
    let customers = [];

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      customers.push({
        rowIndex: i + 1,
        customerId: (row[CUST_ID_COL] || "").toString(),
        name: (row[CUST_NAME_COL] || "").toString(),
        phone: (row[CUST_PHONE_COL] || "").toString(),
        email: (row[CUST_EMAIL_COL] || "").toString(),
        city: (row[CUST_CITY_COL] || "").toString(),
        maritalStatus: (row[CUST_MARITAL_COL] || "").toString(),
        notes: (row[CUST_NOTES_COL] || "").toString(),
        createdAt: (row[CUST_CREATED_AT_COL] || "").toString(),
        linkedUsername: (row[CUST_LINKED_USER_COL] || "").toString()
      });
    }

    return customers;
  } catch (err) {
    return { error: err.message };
  }
}

function addCustomer(customerData) {
  try {
    if (!customerData.name) {
      return { success: false, message: "Customer name is required." };
    }

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CUSTOMERS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Customers sheet not found." };

    const customerId = "CUST-" + new Date().getTime().toString().slice(-6) + Math.floor(Math.random() * 900 + 100);
    const now = new Date().toISOString();

    let linkedUsername = "";
    if (customerData.email) {
      const email = customerData.email.toString().trim().toLowerCase();
      const loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
      if (loginSheet) {
        const loginData = loginSheet.getDataRange().getValues();
        let exists = false;
        for (let i = 1; i < loginData.length; i++) {
          if ((loginData[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase() === email) {
            exists = true;
            break;
          }
        }

        if (!exists) {
          const defaultPassword = "guest" + Math.floor(Math.random() * 9000 + 1000);
          loginSheet.appendRow([email, defaultPassword, "user", "", ""]);
        }
        linkedUsername = email;
      }
    }

    sheet.appendRow([
      customerId,
      customerData.name || "",
      customerData.phone || "",
      customerData.email || "",
      customerData.city || "",
      customerData.maritalStatus || "",
      customerData.notes || "",
      now,
      linkedUsername
    ]);

    return { success: true, message: "Customer added successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateCustomer(rowIndex, customerData) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CUSTOMERS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Customers sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    if (customerData.name !== undefined) sheet.getRange(rowIndex, CUST_NAME_COL + 1).setValue(customerData.name);
    if (customerData.phone !== undefined) sheet.getRange(rowIndex, CUST_PHONE_COL + 1).setValue(customerData.phone);
    if (customerData.email !== undefined) sheet.getRange(rowIndex, CUST_EMAIL_COL + 1).setValue(customerData.email);
    if (customerData.city !== undefined) sheet.getRange(rowIndex, CUST_CITY_COL + 1).setValue(customerData.city);
    if (customerData.maritalStatus !== undefined) sheet.getRange(rowIndex, CUST_MARITAL_COL + 1).setValue(customerData.maritalStatus);
    if (customerData.notes !== undefined) sheet.getRange(rowIndex, CUST_NOTES_COL + 1).setValue(customerData.notes);

    return { success: true, message: "Customer updated successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteCustomer(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(CUSTOMERS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Customers sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Cannot delete header row." };

    sheet.deleteRow(rowIndex);

    return { success: true, message: "Customer deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
