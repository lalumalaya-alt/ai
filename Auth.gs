/***************************************************
 * LOGIN LOGIC
 ***************************************************/
function checkLogin(username, password) {
  try {
    const loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    const data = loginSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim();
      let storedPass = (data[i][LOGIN_PASSWORD_COL] || "").toString().trim();
      let storedRole = (data[i][LOGIN_ROLE_COL] || "").toString().trim().toLowerCase();

      if (storedUser === username && storedPass === password) {
        return {
          success: true,
          message: "Login successful",
          username: storedUser,
          role: storedRole === 'admin' ? 'admin' : 'user'
        };
      }
    }
    return { success: false, message: "Invalid credentials", role: null };
  } catch (e) {
    return { success: false, message: e.toString(), role: null };
  }
}

function createUserIfNotExists(email, generatedPassword) {
  const loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
  const data = loginSheet.getDataRange().getValues();

  let userExists = false;
  for (let i = 1; i < data.length; i++) {
    let storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim();
    if (storedUser === email) {
      userExists = true;
      break;
    }
  }

  if (!userExists) {
    loginSheet.appendRow([email, generatedPassword, "user", "", ""]);
  }
}

function changePassword(username, oldPassword, newPassword) {
  try {
    if (!username || !oldPassword || !newPassword) {
      return { success: false, message: "All fields are required." };
    }
    if (newPassword.length < 4) {
      return { success: false, message: "New password must be at least 4 characters." };
    }

    const loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    const data = loginSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim();
      let storedPass = (data[i][LOGIN_PASSWORD_COL] || "").toString().trim();

      if (storedUser === username && storedPass === oldPassword) {
        loginSheet.getRange(i + 1, LOGIN_PASSWORD_COL + 1).setValue(newPassword);
        return { success: true, message: "Password changed successfully!" };
      }
    }
    return { success: false, message: "Current password is incorrect." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/***************************************************
 * CREATE ACCOUNT (Self-Registration)
 ***************************************************/
function createAccount(email, password) {
  try {
    if (!email || !password) {
      return { success: false, message: "Email and password are required." };
    }

    email = email.toString().trim().toLowerCase();

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return { success: false, message: "Please enter a valid email address." };
    }

    if (password.length < 4) {
      return { success: false, message: "Password must be at least 4 characters." };
    }

    var loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    var data = loginSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase();
      if (storedUser === email) {
        return { success: false, message: "An account with this email already exists." };
      }
    }

    loginSheet.appendRow([email, password, "user", "", ""]);

    return { success: true, message: "Account created successfully! You can now login." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/***************************************************
 * FORGOT PASSWORD — OTP FLOW
 ***************************************************/
function sendForgotPasswordOTP(email) {
  try {
    if (!email) {
      return { success: false, message: "Email is required." };
    }

    email = email.toString().trim().toLowerCase();

    var loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    var data = loginSheet.getDataRange().getValues();
    var userRowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase();
      if (storedUser === email) {
        userRowIndex = i + 1;
        break;
      }
    }

    if (userRowIndex === -1) {
      return { success: false, message: "No account found with this email." };
    }

    var otp = Math.floor(1000 + Math.random() * 9000).toString();
    var expiry = new Date(Date.now() + 10 * 60 * 1000).toISOString();

    loginSheet.getRange(userRowIndex, LOGIN_OTP_COL + 1).setValue(otp);
    loginSheet.getRange(userRowIndex, LOGIN_OTP_EXPIRY_COL + 1).setValue(expiry);
    SpreadsheetApp.flush();

    var hotelName = 'MRI Hotel';
    try {
      var setSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(SETTINGS_SHEET_NAME);
      if (setSheet && setSheet.getLastRow() > 1) {
        hotelName = (setSheet.getRange(2, SET_HOTEL_NAME_COL + 1).getValue() || 'MRI Hotel').toString();
      }
    } catch (se) { Logger.log("Could not load hotel name: " + se); }

    MailApp.sendEmail({
      to: email,
      subject: hotelName + ' - Password Reset OTP',
      body: 'Hello,\n\nYour OTP for password reset is: ' + otp + '\n\nThis code is valid for 10 minutes.\n\nIf you did not request this, please ignore this email.\n\n- ' + hotelName
    });

    return { success: true, message: "OTP sent to your email. Check your inbox." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function verifyOTP(email, otp) {
  try {
    if (!email || !otp) {
      return { success: false, message: "Email and OTP are required." };
    }

    email = email.toString().trim().toLowerCase();
    otp = otp.toString().trim();

    var loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    var data = loginSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase();
      if (storedUser === email) {
        var storedOtp = (data[i][LOGIN_OTP_COL] || "").toString().trim();
        var storedExpiry = (data[i][LOGIN_OTP_EXPIRY_COL] || "").toString().trim();

        if (storedOtp !== otp) {
          return { success: false, message: "Invalid OTP. Please try again." };
        }

        if (!storedExpiry || new Date(storedExpiry) < new Date()) {
          loginSheet.getRange(i + 1, LOGIN_OTP_COL + 1).setValue("");
          loginSheet.getRange(i + 1, LOGIN_OTP_EXPIRY_COL + 1).setValue("");
          return { success: false, message: "OTP has expired. Please request a new one." };
        }

        return { success: true, message: "OTP verified successfully." };
      }
    }

    return { success: false, message: "No account found with this email." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function resetPassword(email, otp, newPassword) {
  try {
    if (!email || !otp || !newPassword) {
      return { success: false, message: "All fields are required." };
    }
    if (newPassword.length < 4) {
      return { success: false, message: "New password must be at least 4 characters." };
    }

    email = email.toString().trim().toLowerCase();
    otp = otp.toString().trim();

    var loginSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LOGIN_SHEET_NAME);
    var data = loginSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var storedUser = (data[i][LOGIN_USERNAME_COL] || "").toString().trim().toLowerCase();
      if (storedUser === email) {
        var storedOtp = (data[i][LOGIN_OTP_COL] || "").toString().trim();
        var storedExpiry = (data[i][LOGIN_OTP_EXPIRY_COL] || "").toString().trim();

        if (storedOtp !== otp) {
          return { success: false, message: "Invalid OTP." };
        }
        if (!storedExpiry || new Date(storedExpiry) < new Date()) {
          loginSheet.getRange(i + 1, LOGIN_OTP_COL + 1).setValue("");
          loginSheet.getRange(i + 1, LOGIN_OTP_EXPIRY_COL + 1).setValue("");
          return { success: false, message: "OTP has expired. Please request a new one." };
        }

        loginSheet.getRange(i + 1, LOGIN_PASSWORD_COL + 1).setValue(newPassword);
        loginSheet.getRange(i + 1, LOGIN_OTP_COL + 1).setValue("");
        loginSheet.getRange(i + 1, LOGIN_OTP_EXPIRY_COL + 1).setValue("");

        return { success: true, message: "Password reset successfully! You can now login." };
      }
    }

    return { success: false, message: "No account found with this email." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
