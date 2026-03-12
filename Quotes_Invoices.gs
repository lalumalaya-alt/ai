/***************************************************
 * QUOTE MANAGEMENT
 ***************************************************/
function getAllQuotes() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let quotes = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      quotes.push({
        rowIndex: i + 1,
        quoteId: (row[QUOTE_ID_COL] || "").toString(),
        guestName: (row[QUOTE_GUEST_NAME_COL] || "").toString(),
        phone: (row[QUOTE_PHONE_COL] || "").toString(),
        email: (row[QUOTE_EMAIL_COL] || "").toString(),
        createdDate: (row[QUOTE_CREATED_COL] || "").toString(),
        validUntil: (row[QUOTE_VALID_COL] || "").toString(),
        status: (row[QUOTE_STATUS_COL] || "").toString(),
        items: (row[QUOTE_ITEMS_COL] || "[]").toString(),
        subTotal: parseFloat(row[QUOTE_SUBTOTAL_COL]) || 0,
        tax: parseFloat(row[QUOTE_TAX_COL]) || 0,
        discount: parseFloat(row[QUOTE_DISCOUNT_COL]) || 0,
        totalAmount: parseFloat(row[QUOTE_TOTAL_COL]) || 0,
        notes: (row[QUOTE_NOTES_COL] || "").toString(),
        createdBy: (row[QUOTE_CREATED_BY_COL] || "").toString(),
        currency: (row[QUOTE_CURRENCY_COL] || 'MVR').toString(),
        gstEnabled: row[QUOTE_GST_ENABLED_COL] === true || row[QUOTE_GST_ENABLED_COL] === 'true',
        gstPercent: parseFloat(row[QUOTE_GST_PERCENT_COL]) || 0,
        gstAmount: parseFloat(row[QUOTE_GST_AMOUNT_COL]) || 0,
        greenTaxEnabled: row[QUOTE_GREENTAX_ENABLED_COL] === true || row[QUOTE_GREENTAX_ENABLED_COL] === 'true',
        greenTaxPerNight: parseFloat(row[QUOTE_GREENTAX_RATE_COL]) || 0,
        greenTaxPax: parseFloat(row[QUOTE_GREENTAX_PAX_COL]) || 0,
        greenTaxNights: parseFloat(row[QUOTE_GREENTAX_NIGHTS_COL]) || 0,
        greenTaxAmount: parseFloat(row[QUOTE_GREENTAX_AMOUNT_COL]) || 0,
        customerTIN: (row[QUOTE_CUSTOMER_TIN_COL] || '').toString(),
        convertedToInvoice: (row[QUOTE_CONVERTED_COL] || '').toString(),
        pdfDriveLink: (row[QUOTE_PDF_LINK_COL] || '').toString()
      });
    }
    return quotes;
  } catch (err) {
    return { error: err.message };
  }
}

function getQuoteById(quoteId) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Quotes sheet not found." };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if ((data[i][QUOTE_ID_COL] || "").toString() === quoteId.toString()) {
        return {
          success: true,
          quote: {
            rowIndex: i + 1, quoteId: (data[i][QUOTE_ID_COL] || "").toString(), guestName: (data[i][QUOTE_GUEST_NAME_COL] || "").toString(), phone: (data[i][QUOTE_PHONE_COL] || "").toString(), email: (data[i][QUOTE_EMAIL_COL] || "").toString(), createdDate: (data[i][QUOTE_CREATED_COL] || "").toString(), validUntil: (data[i][QUOTE_VALID_COL] || "").toString(), status: (data[i][QUOTE_STATUS_COL] || "").toString(), items: (data[i][QUOTE_ITEMS_COL] || "[]").toString(), subTotal: parseFloat(data[i][QUOTE_SUBTOTAL_COL]) || 0, tax: parseFloat(data[i][QUOTE_TAX_COL]) || 0, discount: parseFloat(data[i][QUOTE_DISCOUNT_COL]) || 0, totalAmount: parseFloat(data[i][QUOTE_TOTAL_COL]) || 0, notes: (data[i][QUOTE_NOTES_COL] || "").toString(), createdBy: (data[i][QUOTE_CREATED_BY_COL] || "").toString(), currency: (data[i][QUOTE_CURRENCY_COL] || 'MVR').toString(), gstEnabled: data[i][QUOTE_GST_ENABLED_COL] === true || data[i][QUOTE_GST_ENABLED_COL] === 'true', gstPercent: parseFloat(data[i][QUOTE_GST_PERCENT_COL]) || 0, gstAmount: parseFloat(data[i][QUOTE_GST_AMOUNT_COL]) || 0, greenTaxEnabled: data[i][QUOTE_GREENTAX_ENABLED_COL] === true || data[i][QUOTE_GREENTAX_ENABLED_COL] === 'true', greenTaxPerNight: parseFloat(data[i][QUOTE_GREENTAX_RATE_COL]) || 0, greenTaxPax: parseFloat(data[i][QUOTE_GREENTAX_PAX_COL]) || 0, greenTaxNights: parseFloat(data[i][QUOTE_GREENTAX_NIGHTS_COL]) || 0, greenTaxAmount: parseFloat(data[i][QUOTE_GREENTAX_AMOUNT_COL]) || 0, customerTIN: (data[i][QUOTE_CUSTOMER_TIN_COL] || '').toString(), convertedToInvoice: (data[i][QUOTE_CONVERTED_COL] || '').toString(), pdfDriveLink: (data[i][QUOTE_PDF_LINK_COL] || '').toString()
          }
        };
      }
    }
    return { success: false, message: "Quote not found." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addQuote(quoteData) {
  try {
    if (!quoteData.guestName) return { success: false, message: "Guest name is required." };
    if (!quoteData.items || quoteData.items === '[]') return { success: false, message: "At least one item is required." };
    try { JSON.parse(quoteData.items); } catch (jsonErr) { return { success: false, message: "Invalid items format." }; }

    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Quotes sheet not found." };

    const quoteId = getNextSequentialId('quote');
    const createdDate = new Date().toISOString();
    const status = quoteData.status || "Draft";
    const subTotal = parseFloat(quoteData.subTotal) || 0;
    const discount = parseFloat(quoteData.discount) || 0;
    const currency = quoteData.currency || 'MVR';

    const gstEnabled = quoteData.gstEnabled === true;
    const gstPercent = parseFloat(quoteData.gstPercent) || 0;
    const gstAmount = gstEnabled ? (subTotal - discount) * (gstPercent / 100) : 0;
    const greenTaxEnabled = quoteData.greenTaxEnabled === true;
    const greenTaxRate = parseFloat(quoteData.greenTaxPerNight) || 0;
    const greenTaxPax = parseFloat(quoteData.greenTaxPax) || 0;
    const greenTaxNights = parseFloat(quoteData.greenTaxNights) || 0;
    const greenTaxAmount = greenTaxEnabled ? greenTaxRate * greenTaxPax * greenTaxNights : 0;
    const totalAmount = subTotal - discount + gstAmount + greenTaxAmount;

    sheet.appendRow([
      quoteId, quoteData.guestName.trim(), (quoteData.phone || "").trim(), (quoteData.email || "").trim(), createdDate, (quoteData.validUntil || "").trim(), status, quoteData.items, subTotal, 0, discount, Math.round(totalAmount * 100) / 100, (quoteData.notes || "").trim(), (quoteData.createdBy || "").trim(), currency, gstEnabled, gstPercent, Math.round(gstAmount * 100) / 100, greenTaxEnabled, greenTaxRate, greenTaxPax, greenTaxNights, Math.round(greenTaxAmount * 100) / 100, (quoteData.customerTIN || '').trim(), '', ''
    ]);

    try {
      const parsedItems = JSON.parse(quoteData.items);
      const roomsSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
      if (roomsSheet) {
        const roomsData = roomsSheet.getDataRange().getValues();
        parsedItems.forEach(item => {
          if (item.type === 'room' && item.reservedRoomNo) {
            for (let r = 1; r < roomsData.length; r++) {
              if ((roomsData[r][ROOM_NO_COL] || '').toString() === item.reservedRoomNo.toString()) {
                const curStatus = (roomsData[r][ROOM_STATUS_COL] || '').toString().toLowerCase();
                if (curStatus === 'available') { roomsSheet.getRange(r + 1, ROOM_STATUS_COL + 1).setValue("Reserved"); }
                break;
              }
            }
          }
        });
      }
    } catch (reserveErr) { Logger.log("Room reserve error: " + reserveErr); }

    return { success: true, message: "Quote created successfully!", quoteId: quoteId };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateQuote(rowIndex, quoteData) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Quotes sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    if (quoteData.guestName !== undefined) sheet.getRange(rowIndex, QUOTE_GUEST_NAME_COL + 1).setValue(quoteData.guestName);
    if (quoteData.phone !== undefined) sheet.getRange(rowIndex, QUOTE_PHONE_COL + 1).setValue(quoteData.phone);
    if (quoteData.email !== undefined) sheet.getRange(rowIndex, QUOTE_EMAIL_COL + 1).setValue(quoteData.email);
    if (quoteData.validUntil !== undefined) sheet.getRange(rowIndex, QUOTE_VALID_COL + 1).setValue(quoteData.validUntil);
    if (quoteData.status !== undefined) sheet.getRange(rowIndex, QUOTE_STATUS_COL + 1).setValue(quoteData.status);
    if (quoteData.notes !== undefined) sheet.getRange(rowIndex, QUOTE_NOTES_COL + 1).setValue(quoteData.notes);

    if (quoteData.items !== undefined) {
      try { JSON.parse(quoteData.items); } catch (e) { return { success: false, message: "Invalid items format." }; }
      sheet.getRange(rowIndex, QUOTE_ITEMS_COL + 1).setValue(quoteData.items);
    }

    if (quoteData.subTotal !== undefined) sheet.getRange(rowIndex, QUOTE_SUBTOTAL_COL + 1).setValue(parseFloat(quoteData.subTotal) || 0);
    if (quoteData.tax !== undefined) sheet.getRange(rowIndex, QUOTE_TAX_COL + 1).setValue(parseFloat(quoteData.tax) || 0);
    if (quoteData.discount !== undefined) sheet.getRange(rowIndex, QUOTE_DISCOUNT_COL + 1).setValue(parseFloat(quoteData.discount) || 0);
    if (quoteData.currency !== undefined) sheet.getRange(rowIndex, QUOTE_CURRENCY_COL + 1).setValue(quoteData.currency);
    if (quoteData.customerTIN !== undefined) sheet.getRange(rowIndex, QUOTE_CUSTOMER_TIN_COL + 1).setValue(quoteData.customerTIN);

    if (quoteData.gstEnabled !== undefined) sheet.getRange(rowIndex, QUOTE_GST_ENABLED_COL + 1).setValue(quoteData.gstEnabled === true);
    if (quoteData.gstPercent !== undefined) sheet.getRange(rowIndex, QUOTE_GST_PERCENT_COL + 1).setValue(parseFloat(quoteData.gstPercent) || 0);
    if (quoteData.greenTaxEnabled !== undefined) sheet.getRange(rowIndex, QUOTE_GREENTAX_ENABLED_COL + 1).setValue(quoteData.greenTaxEnabled === true);
    if (quoteData.greenTaxPerNight !== undefined) sheet.getRange(rowIndex, QUOTE_GREENTAX_RATE_COL + 1).setValue(parseFloat(quoteData.greenTaxPerNight) || 0);
    if (quoteData.greenTaxPax !== undefined) sheet.getRange(rowIndex, QUOTE_GREENTAX_PAX_COL + 1).setValue(parseFloat(quoteData.greenTaxPax) || 0);
    if (quoteData.greenTaxNights !== undefined) sheet.getRange(rowIndex, QUOTE_GREENTAX_NIGHTS_COL + 1).setValue(parseFloat(quoteData.greenTaxNights) || 0);

    const subTotal = parseFloat(sheet.getRange(rowIndex, QUOTE_SUBTOTAL_COL + 1).getValue()) || 0;
    const discount = parseFloat(sheet.getRange(rowIndex, QUOTE_DISCOUNT_COL + 1).getValue()) || 0;
    const gstEnabled = sheet.getRange(rowIndex, QUOTE_GST_ENABLED_COL + 1).getValue() === true;
    const gstPercent = parseFloat(sheet.getRange(rowIndex, QUOTE_GST_PERCENT_COL + 1).getValue()) || 0;
    const gstAmount = gstEnabled ? (subTotal - discount) * (gstPercent / 100) : 0;
    const greenTaxEnabled = sheet.getRange(rowIndex, QUOTE_GREENTAX_ENABLED_COL + 1).getValue() === true;
    const greenTaxRate = parseFloat(sheet.getRange(rowIndex, QUOTE_GREENTAX_RATE_COL + 1).getValue()) || 0;
    const greenTaxPax = parseFloat(sheet.getRange(rowIndex, QUOTE_GREENTAX_PAX_COL + 1).getValue()) || 0;
    const greenTaxNights = parseFloat(sheet.getRange(rowIndex, QUOTE_GREENTAX_NIGHTS_COL + 1).getValue()) || 0;
    const greenTaxAmount = greenTaxEnabled ? greenTaxRate * greenTaxPax * greenTaxNights : 0;
    const total = subTotal - discount + gstAmount + greenTaxAmount;

    sheet.getRange(rowIndex, QUOTE_GST_AMOUNT_COL + 1).setValue(Math.round(gstAmount * 100) / 100);
    sheet.getRange(rowIndex, QUOTE_GREENTAX_AMOUNT_COL + 1).setValue(Math.round(greenTaxAmount * 100) / 100);
    sheet.getRange(rowIndex, QUOTE_TOTAL_COL + 1).setValue(Math.round(total * 100) / 100);

    if (quoteData.status === 'Sent' || quoteData.status === 'Accepted') {
      try {
        const itemsStr = (quoteData.items || sheet.getRange(rowIndex, QUOTE_ITEMS_COL + 1).getValue() || '[]').toString();
        const parsedItems = JSON.parse(itemsStr);
        const roomsSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(ROOMS_SHEET_NAME);
        if (roomsSheet) {
          const roomsData = roomsSheet.getDataRange().getValues();
          parsedItems.forEach(item => {
            if (item.type === 'room' && item.reservedRoomNo) {
              for (let r = 1; r < roomsData.length; r++) {
                if ((roomsData[r][ROOM_NO_COL] || '').toString() === item.reservedRoomNo.toString()) {
                  const curStatus = (roomsData[r][ROOM_STATUS_COL] || '').toString().toLowerCase();
                  if (curStatus === 'available') { roomsSheet.getRange(r + 1, ROOM_STATUS_COL + 1).setValue("Reserved"); }
                  break;
                }
              }
            }
          });
        }
      } catch (reserveErr) { Logger.log("Room reserve error on update: " + reserveErr); }
    }

    return { success: true, message: "Quote updated successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteQuote(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(QUOTES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Quotes sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Cannot delete header row." };
    sheet.deleteRow(rowIndex);
    return { success: true, message: "Quote deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/***************************************************
 * INVOICE MANAGEMENT
 ***************************************************/
function getAllInvoices() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getDataRange().getValues();
    let records = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      records.push({
        rowIndex: i + 1,
        invoiceId: (row[INV_ID_COL] || '').toString(),
        guestName: (row[INV_GUEST_NAME_COL] || '').toString(),
        phone: (row[INV_PHONE_COL] || '').toString(),
        email: (row[INV_EMAIL_COL] || '').toString(),
        customerTIN: (row[INV_CUSTOMER_TIN_COL] || '').toString(),
        currency: (row[INV_CURRENCY_COL] || 'MVR').toString(),
        createdDate: (row[INV_CREATED_DATE_COL] || '').toString(),
        dueDate: (row[INV_DUE_DATE_COL] || '').toString(),
        status: (row[INV_STATUS_COL] || 'Draft').toString(),
        items: (row[INV_ITEMS_COL] || '[]').toString(),
        subTotal: parseFloat(row[INV_SUBTOTAL_COL]) || 0,
        gstEnabled: row[INV_GST_ENABLED_COL] === true || row[INV_GST_ENABLED_COL] === 'true',
        gstPercent: parseFloat(row[INV_GST_PERCENT_COL]) || 0,
        gstAmount: parseFloat(row[INV_GST_AMOUNT_COL]) || 0,
        greenTaxEnabled: row[INV_GREENTAX_ENABLED_COL] === true || row[INV_GREENTAX_ENABLED_COL] === 'true',
        greenTaxPerNight: parseFloat(row[INV_GREENTAX_RATE_COL]) || 0,
        greenTaxPax: parseFloat(row[INV_GREENTAX_PAX_COL]) || 0,
        greenTaxNights: parseFloat(row[INV_GREENTAX_NIGHTS_COL]) || 0,
        greenTaxAmount: parseFloat(row[INV_GREENTAX_AMOUNT_COL]) || 0,
        discount: parseFloat(row[INV_DISCOUNT_COL]) || 0,
        totalAmount: parseFloat(row[INV_TOTAL_COL]) || 0,
        notes: (row[INV_NOTES_COL] || '').toString(),
        sourceQuoteId: (row[INV_SOURCE_QUOTE_COL] || '').toString(),
        pdfDriveLink: (row[INV_PDF_LINK_COL] || '').toString(),
        createdBy: (row[INV_CREATED_BY_COL] || '').toString(),
        updatedAt: (row[INV_UPDATED_AT_COL] || '').toString()
      });
    }
    return records;
  } catch (err) {
    return { error: err.message };
  }
}

function getInvoiceById(invoiceId) {
  try {
    const invoices = getAllInvoices();
    if (invoices.error) return { success: false, message: invoices.error };
    const found = invoices.find(inv => inv.invoiceId === invoiceId);
    if (!found) return { success: false, message: "Invoice not found." };
    return { success: true, data: found };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function addInvoice(invoiceData) {
  try {
    if (!invoiceData.guestName) return { success: false, message: "Guest name is required." };
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Invoices sheet not found." };

    const id = getNextSequentialId('invoice');
    const now = new Date().toISOString();

    const subTotal = parseFloat(invoiceData.subTotal) || 0;
    const discount = parseFloat(invoiceData.discount) || 0;
    const gstEnabled = invoiceData.gstEnabled === true;
    const gstPercent = parseFloat(invoiceData.gstPercent) || 0;
    const gstAmount = gstEnabled ? (subTotal - discount) * (gstPercent / 100) : 0;
    const greenTaxEnabled = invoiceData.greenTaxEnabled === true;
    const greenTaxRate = parseFloat(invoiceData.greenTaxPerNight) || 0;
    const greenTaxPax = parseFloat(invoiceData.greenTaxPax) || 0;
    const greenTaxNights = parseFloat(invoiceData.greenTaxNights) || 0;
    const greenTaxAmount = greenTaxEnabled ? greenTaxRate * greenTaxPax * greenTaxNights : 0;
    const totalAmount = subTotal - discount + gstAmount + greenTaxAmount;

    sheet.appendRow([
      id, (invoiceData.guestName || '').trim(), (invoiceData.phone || '').trim(), (invoiceData.email || '').trim(), (invoiceData.customerTIN || '').trim(), invoiceData.currency || 'MVR', now, invoiceData.dueDate || '', invoiceData.status || 'Draft', typeof invoiceData.items === 'string' ? invoiceData.items : JSON.stringify(invoiceData.items || []), subTotal, gstEnabled, gstPercent, Math.round(gstAmount * 100) / 100, greenTaxEnabled, greenTaxRate, greenTaxPax, greenTaxNights, Math.round(greenTaxAmount * 100) / 100, discount, Math.round(totalAmount * 100) / 100, (invoiceData.notes || '').trim(), invoiceData.sourceQuoteId || '', '', (invoiceData.createdBy || '').trim(), now
    ]);

    return { success: true, message: "Invoice created successfully!", invoiceId: id };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateInvoice(rowIndex, invoiceData) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Invoices sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };

    const now = new Date().toISOString();
    const subTotal = parseFloat(invoiceData.subTotal) || 0;
    const discount = parseFloat(invoiceData.discount) || 0;
    const gstEnabled = invoiceData.gstEnabled === true;
    const gstPercent = parseFloat(invoiceData.gstPercent) || 0;
    const gstAmount = gstEnabled ? (subTotal - discount) * (gstPercent / 100) : 0;
    const greenTaxEnabled = invoiceData.greenTaxEnabled === true;
    const greenTaxRate = parseFloat(invoiceData.greenTaxPerNight) || 0;
    const greenTaxPax = parseFloat(invoiceData.greenTaxPax) || 0;
    const greenTaxNights = parseFloat(invoiceData.greenTaxNights) || 0;
    const greenTaxAmount = greenTaxEnabled ? greenTaxRate * greenTaxPax * greenTaxNights : 0;
    const totalAmount = subTotal - discount + gstAmount + greenTaxAmount;

    const oldStatus = (sheet.getRange(rowIndex, INV_STATUS_COL + 1).getValue() || '').toString();

    const existingId = sheet.getRange(rowIndex, INV_ID_COL + 1).getValue();
    const existingCreated = sheet.getRange(rowIndex, INV_CREATED_DATE_COL + 1).getValue();
    const existingSource = sheet.getRange(rowIndex, INV_SOURCE_QUOTE_COL + 1).getValue();
    const existingPdf = sheet.getRange(rowIndex, INV_PDF_LINK_COL + 1).getValue();
    const existingCreatedBy = sheet.getRange(rowIndex, INV_CREATED_BY_COL + 1).getValue();

    const row = [
      existingId, (invoiceData.guestName || '').trim(), (invoiceData.phone || '').trim(), (invoiceData.email || '').trim(), (invoiceData.customerTIN || '').trim(), invoiceData.currency || 'MVR', existingCreated, invoiceData.dueDate || '', invoiceData.status || 'Draft', typeof invoiceData.items === 'string' ? invoiceData.items : JSON.stringify(invoiceData.items || []), subTotal, gstEnabled, gstPercent, Math.round(gstAmount * 100) / 100, greenTaxEnabled, greenTaxRate, greenTaxPax, greenTaxNights, Math.round(greenTaxAmount * 100) / 100, discount, Math.round(totalAmount * 100) / 100, (invoiceData.notes || '').trim(), existingSource, existingPdf, existingCreatedBy, now
    ];

    sheet.getRange(rowIndex, 1, 1, 26).setValues([row]);

    const newStatus = (invoiceData.status || 'Draft').toString();
    let paymentRecorded = false;
    if (newStatus === 'Paid' && oldStatus !== 'Paid') {
      try {
        const finSheet = SpreadsheetApp.openById(SS_ID).getSheetByName(FINANCE_SHEET_NAME);
        if (finSheet) {
          const finData = finSheet.getDataRange().getValues();
          let alreadyRecorded = false;
          for (let f = 1; f < finData.length; f++) {
            if ((finData[f][FIN_LINKED_INV_COL] || '').toString() === existingId.toString()) {
              alreadyRecorded = true;
              break;
            }
          }
          if (!alreadyRecorded) {
            addFinanceRecord({
              date: new Date().toISOString().slice(0, 10), type: 'Income', description: 'Payment for ' + existingId, shopSource: 'Invoice Payment', amount: Math.round(totalAmount * 100) / 100, enteredBy: invoiceData.createdBy || existingCreatedBy || '', category: 'Invoice Payment', currency: invoiceData.currency || 'MVR', linkedInvoiceId: existingId.toString()
            });
            recalculateBalances();
            paymentRecorded = true;
          }
        }
      } catch (finErr) { Logger.log("Auto-payment error: " + finErr.message); }
    }

    return { success: true, message: "Invoice updated successfully!", paymentRecorded: paymentRecorded };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteInvoice(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Invoices sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Cannot delete header row." };
    sheet.deleteRow(rowIndex);
    return { success: true, message: "Invoice deleted successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function reopenInvoice(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet) return { success: false, message: "Invoices sheet not found." };
    if (rowIndex <= 1) return { success: false, message: "Invalid row index." };
    sheet.getRange(rowIndex, INV_STATUS_COL + 1).setValue('Draft');
    sheet.getRange(rowIndex, INV_UPDATED_AT_COL + 1).setValue(new Date().toISOString());
    return { success: true, message: "Invoice reopened as Draft." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function checkOverdueInvoices() {
  try {
    const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(INVOICES_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, overdueCount: 0 };

    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let overdueCount = 0;

    for (let i = 1; i < data.length; i++) {
      const status = (data[i][INV_STATUS_COL] || '').toString();
      const dueDateStr = (data[i][INV_DUE_DATE_COL] || '').toString();

      if (status !== 'Draft' && status !== 'Sent') continue;
      if (!dueDateStr) continue;

      const dueDate = new Date(dueDateStr);
      dueDate.setHours(0, 0, 0, 0);

      if (dueDate < today) {
        sheet.getRange(i + 1, INV_STATUS_COL + 1).setValue('Overdue');
        sheet.getRange(i + 1, INV_UPDATED_AT_COL + 1).setValue(new Date().toISOString());
        overdueCount++;
      }
    }

    if (overdueCount > 0) SpreadsheetApp.flush();

    return {
      success: true,
      overdueCount: overdueCount,
      message: overdueCount > 0 ? overdueCount + ' invoice(s) marked as overdue.' : 'No new overdue invoices.'
    };
  } catch (err) {
    return { success: false, message: err.message, overdueCount: 0 };
  }
}

/***************************************************
 * CONVERSIONS & EMAIL FUNCTIONS
 ***************************************************/
function convertQuoteToInvoice(quoteRowIndex, user) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const quotesSheet = ss.getSheetByName(QUOTES_SHEET_NAME);
    if (!quotesSheet) return { success: false, message: "Quotes sheet not found." };

    const quoteRow = quotesSheet.getRange(quoteRowIndex, 1, 1, 26).getValues()[0];
    const quoteId = (quoteRow[QUOTE_ID_COL] || '').toString();
    const converted = (quoteRow[QUOTE_CONVERTED_COL] || '').toString();
    if (converted && converted !== '' && converted !== 'false') {
      return { success: false, message: "This quote has already been converted to invoice " + converted + "." };
    }

    const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
    if (!invoicesSheet) return { success: false, message: "Invoices sheet not found." };

    const invId = getNextSequentialId('invoice');
    const now = new Date().toISOString();

    const items = (quoteRow[QUOTE_ITEMS_COL] || '[]').toString();
    const subTotal = parseFloat(quoteRow[QUOTE_SUBTOTAL_COL]) || 0;
    const discount = parseFloat(quoteRow[QUOTE_DISCOUNT_COL]) || 0;
    const currency = (quoteRow[QUOTE_CURRENCY_COL] || 'MVR').toString();
    const gstEnabled = quoteRow[QUOTE_GST_ENABLED_COL] === true || quoteRow[QUOTE_GST_ENABLED_COL] === 'true';
    const gstPercent = parseFloat(quoteRow[QUOTE_GST_PERCENT_COL]) || 0;
    const gstAmount = parseFloat(quoteRow[QUOTE_GST_AMOUNT_COL]) || 0;
    const greenTaxEnabled = quoteRow[QUOTE_GREENTAX_ENABLED_COL] === true || quoteRow[QUOTE_GREENTAX_ENABLED_COL] === 'true';
    const greenTaxRate = parseFloat(quoteRow[QUOTE_GREENTAX_RATE_COL]) || 0;
    const greenTaxPax = parseFloat(quoteRow[QUOTE_GREENTAX_PAX_COL]) || 0;
    const greenTaxNights = parseFloat(quoteRow[QUOTE_GREENTAX_NIGHTS_COL]) || 0;
    const greenTaxAmount = parseFloat(quoteRow[QUOTE_GREENTAX_AMOUNT_COL]) || 0;
    const totalAmount = parseFloat(quoteRow[QUOTE_TOTAL_COL]) || 0;

    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 30);

    invoicesSheet.appendRow([
      invId, (quoteRow[QUOTE_GUEST_NAME_COL] || '').toString(), (quoteRow[QUOTE_PHONE_COL] || '').toString(), (quoteRow[QUOTE_EMAIL_COL] || '').toString(), (quoteRow[QUOTE_CUSTOMER_TIN_COL] || '').toString(), currency, now, dueDate.toISOString(), 'Draft', items, subTotal, gstEnabled, gstPercent, gstAmount, greenTaxEnabled, greenTaxRate, greenTaxPax, greenTaxNights, greenTaxAmount, discount, totalAmount, (quoteRow[QUOTE_NOTES_COL] || '').toString(), quoteId, '', user || '', now
    ]);

    quotesSheet.getRange(quoteRowIndex, QUOTE_STATUS_COL + 1).setValue('Converted');
    quotesSheet.getRange(quoteRowIndex, QUOTE_CONVERTED_COL + 1).setValue(invId);

    try {
      const parsedItems = JSON.parse(items);
      const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
      const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
      const roomsData = roomsSheet ? roomsSheet.getDataRange().getValues() : [];

      const guestName = (quoteRow[QUOTE_GUEST_NAME_COL] || '').toString();
      const phone = (quoteRow[QUOTE_PHONE_COL] || '').toString();
      const email = (quoteRow[QUOTE_EMAIL_COL] || '').toString();

      parsedItems.forEach(item => {
        if (item.type === 'room' && item.reservedRoomNo) {
          const roomNo = item.reservedRoomNo.toString();

          for (let r = 1; r < roomsData.length; r++) {
            if ((roomsData[r][ROOM_NO_COL] || '').toString() === roomNo) {
              roomsSheet.getRange(r + 1, ROOM_STATUS_COL + 1).setValue("Booked");
              break;
            }
          }

          const ticketId = generateTicketId();
          const checkInDate = new Date();
          const checkOutDate = new Date();
          checkOutDate.setDate(checkOutDate.getDate() + (parseInt(item.nights) || 1));
          const roomRate = parseFloat(item.rate) || 0;
          const nights = parseInt(item.nights) || 1;
          const qty = parseInt(item.quantity) || 1;
          const baseAmount = roomRate * nights * qty;

          bookingsSheet.appendRow([
            ticketId, roomNo, guestName, phone, email, '', '', 'Single', '', checkInDate.toISOString(), checkOutDate.toISOString(), 'Booked', roomRate, 0, 0, 'Invoice', baseAmount, 'Unpaid', 0, '', '', 'None', 0, qty, ''
          ]);
        }
      });
      SpreadsheetApp.flush();
    } catch (bookErr) { Logger.log("Auto-booking from quote conversion error: " + bookErr); }

    return { success: true, message: "Quote converted to invoice " + invId + " successfully!", invoiceId: invId };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function generateDocumentEmailHtml(type, data, settings) {
  const hotelName = settings.hotelName || 'MRI Hotel';
  const hotelAddress = settings.hotelAddress || '';
  const hotelPhone = settings.hotelPhone || '';
  const hotelEmail = settings.hotelEmail || '';
  const cur = data.currency || 'MVR';
  const isInvoice = type === 'invoice';
  const docLabel = isInvoice ? 'INVOICE' : 'QUOTATION';
  const docId = isInvoice ? data.invoiceId : data.quoteId;

  let items = [];
  try { items = JSON.parse(typeof data.items === 'string' ? data.items : '[]'); } catch (e) { items = []; }
  const roomItems = items.filter(i => i.type === 'room');
  const actItems = items.filter(i => i.type === 'activity');
  const svcItems = items.filter(i => i.type === 'service');

  let itemRows = '';
  roomItems.forEach(it => {
    itemRows += '<tr><td>' + (it.roomType || 'Room') + '</td><td>' + (it.quantity || 1) + ' room(s) x ' + (it.nights || 0) + ' night(s) x ' + cur + ' ' + (parseFloat(it.rate) || 0).toFixed(2) + '</td><td class="right">' + cur + ' ' + (parseFloat(it.amount) || 0).toFixed(2) + '</td></tr>';
  });
  actItems.forEach(it => {
    itemRows += '<tr><td>' + (it.description || 'Activity') + '</td><td>' + (it.pax || 1) + ' pax x ' + cur + ' ' + (parseFloat(it.rate) || 0).toFixed(2) + '</td><td class="right">' + cur + ' ' + (parseFloat(it.amount) || 0).toFixed(2) + '</td></tr>';
  });
  svcItems.forEach(it => {
    itemRows += '<tr><td>' + (it.description || 'Service') + '</td><td>-</td><td class="right">' + cur + ' ' + (parseFloat(it.amount) || 0).toFixed(2) + '</td></tr>';
  });

  const subTotal = parseFloat(data.subTotal) || 0;
  const discount = parseFloat(data.discount) || 0;
  const gstAmount = parseFloat(data.gstAmount) || 0;
  const greenTaxAmount = parseFloat(data.greenTaxAmount) || 0;
  const totalAmount = parseFloat(data.totalAmount) || 0;

  let totalsRows = '<tr><td colspan="2"><strong>Subtotal</strong></td><td class="right"><strong>' + cur + ' ' + subTotal.toFixed(2) + '</strong></td></tr>';
  if (discount > 0) totalsRows += '<tr><td colspan="2">Discount</td><td class="right">- ' + cur + ' ' + discount.toFixed(2) + '</td></tr>';
  if (data.gstEnabled) totalsRows += '<tr><td colspan="2">GST (' + (data.gstPercent || 0) + '%)</td><td class="right">' + cur + ' ' + gstAmount.toFixed(2) + '</td></tr>';
  if (data.greenTaxEnabled) totalsRows += '<tr><td colspan="2">Green Tax</td><td class="right">' + cur + ' ' + greenTaxAmount.toFixed(2) + '</td></tr>';
  totalsRows += '<tr class="total"><td colspan="2"><strong>TOTAL</strong></td><td class="right"><strong>' + cur + ' ' + totalAmount.toFixed(2) + '</strong></td></tr>';

  let dateInfo = '';
  if (isInvoice) {
    dateInfo = '<p><strong>Date:</strong> ' + (data.createdDate || '') + '</p><p><strong>Due Date:</strong> ' + (data.dueDate || '') + '</p><p><strong>Status:</strong> ' + (data.status || '') + '</p>';
  } else {
    dateInfo = '<p><strong>Created:</strong> ' + (data.createdDate || '') + '</p><p><strong>Valid Until:</strong> ' + (data.validUntil || '') + '</p>';
  }

  return '<html><head><style>body{font-family:Arial,sans-serif;margin:20px;color:#333}' +
    '.doc-container{max-width:650px;margin:auto;border:1px solid #ddd;padding:30px;border-radius:4px}' +
    'h2{text-align:center;color:#001f3f;margin-bottom:5px}' +
    '.subtitle{text-align:center;color:#666;font-size:14px;margin-bottom:20px}' +
    'table{width:100%;border-collapse:collapse;margin:15px 0}' +
    'th,td{padding:10px;border:1px solid #ddd;text-align:left;font-size:14px}' +
    'th{background:#001f3f;color:white}' +
    '.right{text-align:right}' +
    '.total{font-weight:bold;background:#f0f0f0}' +
    '.hotel-info{text-align:center;color:#666;font-size:13px;margin-bottom:20px}' +
    '.footer{text-align:center;margin-top:25px;padding-top:15px;border-top:1px solid #ddd;color:#888;font-size:12px}' +
    '</style></head><body><div class="doc-container">' +
    '<h2>' + hotelName + '</h2>' +
    '<p class="subtitle">' + docLabel + '</p>' +
    '<div class="hotel-info">' + (hotelAddress ? hotelAddress + '<br>' : '') + (hotelPhone ? 'Phone: ' + hotelPhone + ' | ' : '') + (hotelEmail ? 'Email: ' + hotelEmail : '') + '</div>' +
    '<p><strong>' + docLabel + ' #:</strong> ' + docId + '</p>' +
    '<p><strong>Guest:</strong> ' + (data.guestName || '') + '</p>' +
    '<p><strong>Email:</strong> ' + (data.email || '') + '</p>' +
    (data.phone ? '<p><strong>Phone:</strong> ' + data.phone + '</p>' : '') +
    (data.customerTIN ? '<p><strong>Customer TIN:</strong> ' + data.customerTIN + '</p>' : '') +
    dateInfo +
    '<table><tr><th>Description</th><th>Details</th><th class="right">Amount (' + cur + ')</th></tr>' +
    itemRows + totalsRows + '</table>' +
    (data.notes ? '<p style="font-style:italic;color:#666">Notes: ' + data.notes + '</p>' : '') +
    '<div class="footer"><p>Thank you for choosing ' + hotelName + '!</p></div>' +
    '</div></body></html>';
}

function generateInvoiceHtml(invoiceData) {
  let { ticketId, occupantName, email, roomNo, checkIn, checkOut, nights, roomRate, discount, tax, finalAmount, currency } = invoiceData;
  const cur = currency || 'MVR';

  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
          .invoice-container { max-width: 600px; margin: auto; border: 1px solid #ccc; padding: 20px; }
          h2, h3 { text-align: center; color: #001f3f; }
          table { width: 100%; border-collapse: collapse; }
          th, td { padding: 8px; border: 1px solid #ddd; text-align: left; }
          th { background: #001f3f; color: white; }
          .right { text-align: right; }
          .total { font-weight: bold; background: #f0f0f0; }
        </style>
      </head>
      <body>
        <div class="invoice-container">
          <h2>MRI Hotel - Invoice</h2>
          <p><strong>Ticket ID:</strong> ${ticketId}</p>
          <p><strong>Guest Name:</strong> ${occupantName}</p>
          <p><strong>Email:</strong> ${email}</p>
          <p><strong>Room #:</strong> ${roomNo}</p>
          <p><strong>Check-in:</strong> ${checkIn}</p>
          <p><strong>Check-out:</strong> ${checkOut}</p>
          <p><strong>Nights Stayed:</strong> ${nights}</p>
          <hr>
          <table>
            <tr><th>Description</th><th class="right">Amount (${cur})</th></tr>
            <tr><td>Room Rate (${nights} night${nights > 1 ? 's' : ''} x ${cur} ${roomRate.toFixed(2)})</td><td class="right">${cur} ${(roomRate * nights).toFixed(2)}</td></tr>
            <tr><td>Discount</td><td class="right">- ${cur} ${discount.toFixed(2)}</td></tr>
            <tr><td>Tax</td><td class="right">${cur} ${tax.toFixed(2)}</td></tr>
            <tr class="total"><td>Total Amount</td><td class="right">${cur} ${finalAmount.toFixed(2)}</td></tr>
          </table>
          <hr>
          <p style="text-align:center;">Thank you for staying with us!</p>
        </div>
      </body>
    </html>
  `;
}

function emailInvoice(invoiceId) {
  try {
    const result = getInvoiceById(invoiceId);
    if (!result.success) return result;
    const inv = result.data;

    if (!inv.email) return { success: false, message: "No email address on this invoice." };

    const settingsResult = getSettings();
    const settings = settingsResult.success ? settingsResult.data : { hotelName: 'MRI Hotel' };

    const htmlBody = generateDocumentEmailHtml('invoice', inv, settings);
    const subject = settings.hotelName + ' - Invoice ' + inv.invoiceId;

    MailApp.sendEmail({
      to: inv.email,
      subject: subject,
      body: 'Dear ' + inv.guestName + ',\n\nPlease find your invoice ' + inv.invoiceId + ' for ' + (inv.currency || 'MVR') + ' ' + inv.totalAmount.toFixed(2) + '.\n\nThank you!\n' + settings.hotelName,
      htmlBody: htmlBody
    });

    return { success: true, message: "Invoice emailed to " + inv.email + " successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function emailQuote(quoteId) {
  try {
    const result = getQuoteById(quoteId);
    if (!result.success) return result;
    const q = result.quote;

    if (!q.email) return { success: false, message: "No email address on this quote." };

    const settingsResult = getSettings();
    const settings = settingsResult.success ? settingsResult.data : { hotelName: 'MRI Hotel' };

    const htmlBody = generateDocumentEmailHtml('quote', q, settings);
    const subject = settings.hotelName + ' - Quotation ' + q.quoteId;

    MailApp.sendEmail({
      to: q.email,
      subject: subject,
      body: 'Dear ' + q.guestName + ',\n\nPlease find your quotation ' + q.quoteId + ' for ' + (q.currency || 'MVR') + ' ' + q.totalAmount.toFixed(2) + '.\nValid until: ' + (q.validUntil || 'N/A') + '\n\nThank you!\n' + settings.hotelName,
      htmlBody: htmlBody
    });

    return { success: true, message: "Quote emailed to " + q.email + " successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}
