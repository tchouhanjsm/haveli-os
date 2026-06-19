/**
 * SuiteFlow India - Backend Controller (Google Apps Script)
 * Handles Database initialization, RBAC operations, WhatsApp parsing, 
 * Housekeeping, Shift Handover logging, and Automated checkout PDF compiling.
 */

// 1. WEB APP ENTRY POINT
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SuiteFlow Operations Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. CONFIGURATION & AUTO-INITIALIZATION OF DATABASE
function getActiveDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Bookings Ledger
  let bookingsSheet = ss.getSheetByName('Bookings');
  if (!bookingsSheet) {
    bookingsSheet = ss.insertSheet('Bookings');
    bookingsSheet.appendRow(['ID', 'Guest Name', 'Check-In', 'Check-Out', 'Company', 'Note', 'Status', 'Room Charge', 'Extra Charge', 'Total Amount', 'Created At']);
    bookingsSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#EEF2F6');
  }
  
  // Daily Expenses Ledger
  let expensesSheet = ss.getSheetByName('Expenses');
  if (!expensesSheet) {
    expensesSheet = ss.insertSheet('Expenses');
    expensesSheet.appendRow(['ID', 'Date', 'Category', 'Amount', 'Description']);
    expensesSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#EEF2F6');
  }

  // Room Grid Table (Housekeeping status storage)
  let roomsSheet = ss.getSheetByName('Rooms');
  if (!roomsSheet) {
    roomsSheet = ss.insertSheet('Rooms');
    roomsSheet.appendRow(['Room Number', 'Status']);
    // Seed standard hotel rooms
    const standardRooms = [
      ['101', 'ready'], ['102', 'ready'], ['103', 'ready'], ['104', 'ready'],
      ['201', 'ready'], ['202', 'ready'], ['203', 'ready'], ['204', 'ready']
    ];
    roomsSheet.getRange(2, 1, standardRooms.length, 2).setValues(standardRooms);
    roomsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#EEF2F6');
  }

  // Digital Shift Handover Bulletins
  let handoversSheet = ss.getSheetByName('Handovers');
  if (!handoversSheet) {
    handoversSheet = ss.insertSheet('Handovers');
    handoversSheet.appendRow(['ID', 'Author', 'Timestamp', 'Message']);
    handoversSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#EEF2F6');
  }

  // Dynamic Settings Store
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Key', 'Value', 'Instructions']);
    settingsSheet.appendRow(['INVOICE_TEMPLATE_DOC_ID', '', 'Paste Google Doc Invoice Template ID here']);
    settingsSheet.appendRow(['INVOICES_FOLDER_ID', '', 'Paste Google Drive Folder ID here for saved PDFs']);
    settingsSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#EEF2F6');
  }

  return ss;
}

// 3. SETTINGS MANIPULATION
function getSetting(key) {
  const ss = getActiveDatabase();
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return '';
}

function fetchSettings() {
  return {
    INVOICE_TEMPLATE_DOC_ID: getSetting('INVOICE_TEMPLATE_DOC_ID'),
    INVOICES_FOLDER_ID: getSetting('INVOICES_FOLDER_ID')
  };
}

function saveSettings(settings) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Settings');
    const data = sheet.getDataRange().getValues();
    
    for (let key in settings) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settings[key]);
        }
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 4. CRUD: BOOKINGS
function getBookings() {
  const ss = getActiveDatabase();
  const sheet = ss.getSheetByName('Bookings');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const bookings = [];
  
  for (let i = 1; i < data.length; i++) {
    let booking = {};
    for (let j = 0; j < headers.length; j++) {
      let value = data[i][j];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      booking[headers[j].replace(/\s+/g, '')] = value;
    }
    bookings.push(booking);
  }
  return bookings.reverse(); // Return newest first
}

function saveBooking(booking) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Bookings');
    const id = 'BK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);
    
    sheet.appendRow([
      id,
      booking.guestName,
      booking.checkIn,
      booking.checkOut,
      booking.company || 'N/A',
      booking.note || '',
      'Checked-In',
      0, // Base Room Cost
      0, // Extras (Laundry/Food)
      0, // Total Invoice Amount
      new Date()
    ]);
    
    return { success: true, id: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 5. CRUD: HOUSEKEEPING GRID
function getRooms() {
  const ss = getActiveDatabase();
  const sheet = ss.getSheetByName('Rooms');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rooms = [];
  for (let i = 1; i < data.length; i++) {
    rooms.push({
      number: data[i][0].toString(),
      status: data[i][1]
    });
  }
  return rooms;
}

function updateRoom(roomNumber, status) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Rooms');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === roomNumber.toString()) {
        sheet.getRange(i + 1, 2).setValue(status);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 6. CRUD: DAILY EXPENSES
function getExpenses() {
  const ss = getActiveDatabase();
  const sheet = ss.getSheetByName('Expenses');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const expenses = [];
  
  for (let i = 1; i < data.length; i++) {
    let exp = {};
    for (let j = 0; j < headers.length; j++) {
      let value = data[i][j];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      exp[headers[j]] = value;
    }
    expenses.push(exp);
  }
  return expenses.reverse();
}

function saveExpense(expense) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Expenses');
    const id = 'EXP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);
    
    sheet.appendRow([
      id,
      expense.date,
      expense.category,
      expense.amount,
      expense.description || ''
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 7. CRUD: SHIFT HANDOVER LOGS
function getHandovers() {
  const ss = getActiveDatabase();
  const sheet = ss.getSheetByName('Handovers');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const handovers = [];
  for (let i = 1; i < data.length; i++) {
    handovers.push({
      id: data[i][0],
      author: data[i][1],
      time: data[i][2],
      message: data[i][3]
    });
  }
  return handovers.reverse();
}

function saveHandover(message) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Handovers');
    const id = 'HO-' + Math.floor(100 + Math.random() * 900);
    
    // Explicit Indian Time Zone Formatter for Handover timestamps
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd MMM yyyy, hh:mm a");
    
    sheet.appendRow([
      id,
      "Duty Staff Member",
      timestamp,
      message
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 8. METRICS FOR MAIN OVERVIEW CARD
function getDashboardMetrics() {
  try {
    const bookings = getBookings();
    const expenses = getExpenses();
    
    const activeCheckins = bookings.filter(b => b.Status === 'Checked-In').length;
    
    const totalEarnings = bookings
      .filter(b => b.Status === 'Checked-Out-Invoiced')
      .reduce((sum, b) => sum + parseFloat(b.TotalAmount || 0), 0);
      
    const totalExpenses = expenses
      .reduce((sum, e) => sum + parseFloat(e.Amount || 0), 0);
      
    return {
      activeCheckins,
      totalEarnings: totalEarnings.toFixed(2),
      totalExpenses: totalExpenses.toFixed(2),
      netProfit: (totalEarnings - totalExpenses).toFixed(2)
    };
  } catch (e) {
    return { activeCheckins: 0, totalEarnings: '0.00', totalExpenses: '0.00', netProfit: '0.00' };
  }
}

// 9. COPIES GOOGLE DOC TEMPLATE -> POPULATES DATA -> GENERATES INVOICE PDF
function checkoutAndGenerateInvoice(checkoutData) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Bookings');
    const data = sheet.getDataRange().getValues();
    let rowNum = -1;
    let bRecord = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === checkoutData.id) {
        rowNum = i + 1;
        bRecord = {
          id: data[i][0],
          guestName: data[i][1],
          checkIn: Utilities.formatDate(new Date(data[i][2]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          checkOut: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          company: data[i][4] || 'N/A',
          note: data[i][5] || ''
        };
        break;
      }
    }

    if (rowNum === -1) {
      throw new Error("Target guest ID not found: " + checkoutData.id);
    }

    const roomRate = parseFloat(checkoutData.roomCharge);
    const extraRate = parseFloat(checkoutData.extraCharge);
    const total = roomRate + extraRate;

    // Save financials and status back to Database Sheet
    sheet.getRange(rowNum, 7).setValue('Checked-Out-Invoiced');
    sheet.getRange(rowNum, 8).setValue(roomRate);
    sheet.getRange(rowNum, 9).setValue(extraRate);
    sheet.getRange(rowNum, 10).setValue(total);

    // Retrieve Document generation settings
    const templateDocId = getSetting('INVOICE_TEMPLATE_DOC_ID');
    const targetFolderId = getSetting('INVOICES_FOLDER_ID');

    if (!templateDocId || !targetFolderId) {
      return { 
        success: true, 
        message: 'Checked out in database! Automated PDF invoice generation skipped because Template ID or Folder ID is missing from Setup tab.'
      };
    }

    // Connect to target G-Drive Directories
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const templateFile = DriveApp.getFileById(templateDocId);
    const newDocFile = templateFile.makeCopy(`Invoice_${bRecord.guestName}_${bRecord.id}`, targetFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();

    const invoiceNum = bRecord.id.replace('BK-', 'INV-');
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Replace designated brackets inside GDoc
    body.replaceText('{{INVOICE_NUMBER}}', invoiceNum);
    body.replaceText('{{DATE}}', today);
    body.replaceText('{{GUEST_NAME}}', bRecord.guestName);
    body.replaceText('{{COMPANY}}', bRecord.company);
    body.replaceText('{{CHECKIN}}', bRecord.checkIn);
    body.replaceText('{{CHECKOUT}}', bRecord.checkOut);
    body.replaceText('{{ROOM_CHARGE}}', 'Rs. ' + roomRate.toLocaleString('en-IN'));
    body.replaceText('{{EXTRA_CHARGE}}', 'Rs. ' + extraRate.toLocaleString('en-IN'));
    body.replaceText('{{TOTAL_AMOUNT}}', 'Rs. ' + total.toLocaleString('en-IN'));
    body.replaceText('{{NOTES}}', bRecord.note || 'None');

    newDoc.saveAndClose();

    // Compile GDoc dynamically to readable PDF
    const pdfBlob = newDocFile.getAs(MimeType.PDF);
    const pdfFile = targetFolder.createFile(pdfBlob);
    
    // Trash copy of the temporary Doc file (Leaves GDrive clean)
    newDocFile.setTrashed(true);

    return { 
      success: true, 
      message: 'Checkout complete. PDF generated & archived in G-Drive folder!', 
      pdfUrl: pdfFile.getUrl() 
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
