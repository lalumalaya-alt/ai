/***************************************************
 * SETTINGS & STORAGE MANAGEMENT
 ***************************************************/
function getSettings() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: {
        hotelName: 'MRI Hotel', hotelAddress: '', hotelPhone: '', hotelEmail: '', hotelTIN: '',
        logoFileId: '', logoUrl: '', defaultCurrency: 'MVR', gstDefaultPercent: 16,
        greenTaxDefaultRate: 6, nextInvoiceNum: 1, nextQuoteNum: 1, pdfFolderId: '', logoFolderId: ''
      }};
    }
    const row = sheet.getRange(2, 1, 1, 14).getValues()[0];
    return { success: true, data: {
      hotelName: (row[SET_HOTEL_NAME_COL] || 'MRI Hotel').toString(),
      hotelAddress: (row[SET_HOTEL_ADDRESS_COL] || '').toString(),
      hotelPhone: (row[SET_HOTEL_PHONE_COL] || '').toString(),
      hotelEmail: (row[SET_HOTEL_EMAIL_COL] || '').toString(),
      hotelTIN: (row[SET_HOTEL_TIN_COL] || '').toString(),
      logoFileId: (row[SET_LOGO_FILE_ID_COL] || '').toString(),
      logoUrl: (row[SET_LOGO_URL_COL] || '').toString(),
      defaultCurrency: (row[SET_DEFAULT_CURRENCY_COL] || 'MVR').toString(),
      gstDefaultPercent: parseFloat(row[SET_GST_DEFAULT_COL]) || 16,
      greenTaxDefaultRate: parseFloat(row[SET_GREENTAX_DEFAULT_COL]) || 6,
      nextInvoiceNum: parseInt(row[SET_NEXT_INVOICE_COL]) || 1,
      nextQuoteNum: parseInt(row[SET_NEXT_QUOTE_COL]) || 1,
      pdfFolderId: (row[SET_PDF_FOLDER_ID_COL] || '').toString(),
      logoFolderId: (row[SET_LOGO_FOLDER_ID_COL] || '').toString()
    }};
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateSettings(settingsData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) return { success: false, message: "Settings sheet not found." };

    let currentInvoiceNum = 1;
    let currentQuoteNum = 1;
    let currentPdfFolderId = settingsData.pdfFolderId || '';
    let currentLogoFolderId = settingsData.logoFolderId || '';
    if (sheet.getLastRow() >= 2) {
      const existing = sheet.getRange(2, 1, 1, 14).getValues()[0];
      currentInvoiceNum = parseInt(existing[SET_NEXT_INVOICE_COL]) || 1;
      currentQuoteNum = parseInt(existing[SET_NEXT_QUOTE_COL]) || 1;
      currentPdfFolderId = (existing[SET_PDF_FOLDER_ID_COL] || '').toString() || currentPdfFolderId;
      currentLogoFolderId = (existing[SET_LOGO_FOLDER_ID_COL] || '').toString() || currentLogoFolderId;
    }

    const row = [
      settingsData.hotelName || 'MRI Hotel',
      settingsData.hotelAddress || '',
      settingsData.hotelPhone || '',
      settingsData.hotelEmail || '',
      settingsData.hotelTIN || '',
      settingsData.logoFileId || '',
      settingsData.logoUrl || '',
      settingsData.defaultCurrency || 'MVR',
      parseFloat(settingsData.gstDefaultPercent) || 16,
      parseFloat(settingsData.greenTaxDefaultRate) || 6,
      currentInvoiceNum,
      currentQuoteNum,
      currentPdfFolderId,
      currentLogoFolderId
    ];

    if (sheet.getLastRow() < 2) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(2, 1, 1, 14).setValues([row]);
    }

    return { success: true, message: "Settings updated successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function uploadLogo(base64Data, fileName, mimeType) {
  try {
    const folder = getOrCreateDriveFolder("Hotel Invoice App Logos");
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const logoUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400";

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, SET_LOGO_FILE_ID_COL + 1).setValue(fileId);
      sheet.getRange(2, SET_LOGO_URL_COL + 1).setValue(logoUrl);
      sheet.getRange(2, SET_LOGO_FOLDER_ID_COL + 1).setValue(folder.getId());
    }

    return { success: true, message: "Logo uploaded successfully!", fileId: fileId, logoUrl: logoUrl };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function savePdfToDrive(base64PdfData, fileName, recordId, type) {
  try {
    const folder = getOrCreateDriveFolder("Hotel Invoice PDFs");
    const decoded = Utilities.base64Decode(base64PdfData);
    const blob = Utilities.newBlob(decoded, 'application/pdf', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();

    const ss = SpreadsheetApp.openById(SS_ID);
    if (type === 'invoice') {
      const sheet = ss.getSheetByName(INVOICES_SHEET_NAME);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if ((data[i][INV_ID_COL] || '').toString() === recordId) {
            sheet.getRange(i + 1, INV_PDF_LINK_COL + 1).setValue(fileUrl);
            break;
          }
        }
      }
    } else if (type === 'quote') {
      const sheet = ss.getSheetByName(QUOTES_SHEET_NAME);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if ((data[i][QUOTE_ID_COL] || '').toString() === recordId) {
            sheet.getRange(i + 1, QUOTE_PDF_LINK_COL + 1).setValue(fileUrl);
            break;
          }
        }
      }
    }

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      settingsSheet.getRange(2, SET_PDF_FOLDER_ID_COL + 1).setValue(folder.getId());
    }

    return { success: true, message: "PDF saved to Drive!", fileUrl: fileUrl };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
