/**
 * Hotel Manager & Automation System - Backend
 * Google Apps Script (GAS) Web App Controller
 * Integrates native open-source Invoice-Generator API mapped to Garh Jaisal Heritage Layout
 */

// 1. WEB APP ENTRY POINTS
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Hotel Operations Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 2. CONFIGURATION & DATABASE SETUP (Initializes sheets automatically)
function getActiveDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Bookings Sheet if missing
  let bookingsSheet = ss.getSheetByName('Bookings');
  if (!bookingsSheet) {
    bookingsSheet = ss.insertSheet('Bookings');
    bookingsSheet.appendRow(['ID', 'Guest Name', 'Check-In', 'Check-Out', 'Company', 'Note', 'Status', 'Room Charge', 'Extra Charge', 'Total Amount', 'Created At']);
    bookingsSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#E2E8F0');
  }
  
  // Create Expenses Sheet if missing
  let expensesSheet = ss.getSheetByName('Expenses');
  if (!expensesSheet) {
    expensesSheet = ss.insertSheet('Expenses');
    expensesSheet.appendRow(['ID', 'Date', 'Category', 'Amount', 'Description']);
    expensesSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E2E8F0');
  }

  // Create Settings Sheet for template configurations
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Key', 'Value', 'Instructions']);
    settingsSheet.appendRow(['INVOICE_GENERATOR_API_URL', 'https://invoice-generator.com', 'API Endpoint (Change if self-hosting)']);
    settingsSheet.appendRow(['INVOICES_FOLDER_ID', '', 'Paste Google Drive Folder ID here for saved PDFs']);
    settingsSheet.appendRow(['HOTEL_BILLING_NAME', 'Hotel Garh Jaisal Haveli', 'Your Hotel Name for Bills']);
    settingsSheet.appendRow(['HOTEL_ADDRESS', 'Inside Fort Kotari Para,\nJaisalmer, Rajasthan, 345001', 'Multiline address info']);
    settingsSheet.appendRow(['HOTEL_GSTIN', '08AGHPC9718Q2ZP', 'Your Hotel GST Number']);
    settingsSheet.appendRow(['HOTEL_EMAIL', 'mukeshkila@yahoo.com', 'Contact Email']);
    settingsSheet.appendRow(['HOTEL_PHONE', '+91-7727-047-567', 'Contact Phone']);
    settingsSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E2E8F0');
  }

  // Create Housekeeping Rooms state if missing
  let roomsSheet = ss.getSheetByName('Rooms');
  if (!roomsSheet) {
    roomsSheet = ss.insertSheet('Rooms');
    roomsSheet.appendRow(['Room Number', 'Status']);
    const rooms = ['101', '102', '103', '104', '201', '202', '203', '204'];
    rooms.forEach(r => roomsSheet.appendRow([r, 'ready']));
    roomsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E2E8F0');
  }

  // Create Shift Handover log sheet if missing
  let handoverSheet = ss.getSheetByName('Handovers');
  if (!handoverSheet) {
    handoverSheet = ss.insertSheet('Handovers');
    handoverSheet.appendRow(['ID', 'Author', 'Timestamp', 'Message']);
    handoverSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#E2E8F0');
  }

  return ss;
}

// Helper to query settings
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

// 3. CRUD BOOKINGS
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
    booking.rowNumber = i + 1;
    bookings.push(booking);
  }
  
  return bookings.reverse();
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
      0,
      0,
      0,
      new Date()
    ]);
    
    return { success: true, message: 'Booking created successfully!', id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// 4. HOUSEKEEPING ROOM LOGIC
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

function updateRoom(number, status) {
  try {
    const ss = getActiveDatabase();
    const sheet = ss.getSheetByName('Rooms');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === number.toString()) {
        sheet.getRange(i + 1, 2).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: "Room not found." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 5. SHIFT HANDOVER BULLETIN LOGIC
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
    const author = 'Duty Staff';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');
    
    sheet.appendRow([id, author, timestamp, message]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 6. CHECKOUT AND INSTANT PDF GENERATION (Replicates Garh Jaisal Heritage Layout)
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
          checkIn: Utilities.formatDate(new Date(data[i][2]), Session.getScriptTimeZone(), 'MMM dd, yyyy'),
          checkOut: Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'MMM dd, yyyy'),
          company: data[i][4] || 'N/A',
          note: data[i][5] || ''
        };
        break;
      }
    }

    if (rowNum === -1) {
      throw new Error("Booking Record not found matching ID: " + checkoutData.id);
    }

    // Days Count Calculation
    const checkInDate = new Date(bRecord.checkIn);
    const checkOutDate = new Date(bRecord.checkOut);
    const diffTime = Math.abs(checkOutDate - checkInDate);
    const nightsCount = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) || 1;

    const baseRoomRate = parseFloat(checkoutData.roomCharge);
    const totalBeforeTax = baseRoomRate + parseFloat(checkoutData.extraCharge);
    
    // We auto-apportion a 12% standard GST calculation as shown in your image
    const gstRatePercent = 12; 
    const totalWithTax = totalBeforeTax * (1 + (gstRatePercent / 100));

    // Commit final calculated state back to Google Sheets database
    sheet.getRange(rowNum, 7).setValue('Checked-Out-Invoiced');
    sheet.getRange(rowNum, 8).setValue(checkoutData.roomCharge);
    sheet.getRange(rowNum, 9).setValue(checkoutData.extraCharge);
    sheet.getRange(rowNum, 10).setValue(totalWithTax);

    // Retrieve custom settings properties
    const apiUrl = getSetting('INVOICE_GENERATOR_API_URL') || 'https://invoice-generator.com';
    const targetFolderId = getSetting('INVOICES_FOLDER_ID');
    const hotelName = getSetting('HOTEL_BILLING_NAME') || 'Hotel Garh Jaisal Haveli';
    const hotelAddress = getSetting('HOTEL_ADDRESS') || 'Inside Fort Kotari Para,\nJaisalmer, Rajasthan, 345001';
    const hotelGstin = getSetting('HOTEL_GSTIN') || '08AGHPC9718Q2ZP';
    const hotelEmail = getSetting('HOTEL_EMAIL') || 'mukeshkila@yahoo.com';
    const hotelPhone = getSetting('HOTEL_PHONE') || '+91-7727-047-567';

    if (!targetFolderId) {
      return { 
        success: true, 
        message: 'Checkout complete in database! Invoice PDF skipped (Target Google Drive Folder ID missing in settings).'
      };
    }

    const invoiceNum = bRecord.id.replace('BK-', '');
    const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy');

    // 1. Construct customized JSON payload to match the image structure
    const invoicePayload = {
      logo: "https://i.imgur.com/vjWn08z.jpeg", // Web hosted Jaisalmer Fort/Hotel logo URL
      from: `${hotelName}\n${hotelAddress}`,
      to: bRecord.company !== 'N/A' ? bRecord.company : bRecord.guestName,
      ship_to_title: "Address",
      ship_to: bRecord.company !== 'N/A' ? `Guest Contact: ${bRecord.guestName}\nStay Duration Details` : "Direct Walk-In Guest Details",
      number: invoiceNum,
      date: todayFormatted,
      currency: "INR",
      payment_terms: "Immediate",
      custom_fields: [
        { "name": "GJ GSTIN", "value": hotelGstin },
        { "name": "Checkin Date", "value": bRecord.checkIn },
        { "name": "Checkout Date", "value": bRecord.checkOut }
      ],
      // Override column header titles precisely matching your uploaded template
      item_header: `${bRecord.guestName} Stay Details`,
      quantity_header: "Night",
      unit_cost_header: "Rate",
      amount_header: "Amount",
      balance_title: "Balance Due",
      tax_title: `GST (${gstRatePercent}%)`,
      fields: {
        tax: "%",
        discounts: false,
        shipping: false
      },
      items: [
        {
          name: "Heritage Room Stay Accommodation",
          quantity: nightsCount,
          unit_cost: baseRoomRate / nightsCount
        }
      ],
      tax: gstRatePercent,
      amount_paid: totalWithTax, // Full payment assumed at checkout
      terms_title: "Terms",
      terms: `- Every tourist will be responsible for abiding by local laws as dictated by local authorities with absolutely no obligations whatsoever on ${hotelName}.\n- No cancellation charges if the reservation is canceled more than 15 days prior to arrival. Bank charges are applicable in case of any refund.\n- The above cancellation policy is not valid during long weekends, Christmas, New Year, and during certain other times of the year wherein rooms will be confirmed subject to non-refundable advance.\n- Every guest must show their ID at the time of check-in.\n- For any enquiries or special requests, email us at ${hotelEmail} or call us at ${hotelPhone}`
    };

    // Include Extra amenities dynamically if logged at checkout
    if (parseFloat(checkoutData.extraCharge) > 0) {
      invoicePayload.items.push({
        name: "Laundry / F&B Amenities / Extra Bed Charges",
        quantity: 1,
        unit_cost: parseFloat(checkoutData.extraCharge)
      });
    }

    // 2. Transmit HTTP POST Request directly to endpoint API
    const fetchOptions = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(invoicePayload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, fetchOptions);

    if (response.getResponseCode() !== 200) {
      throw new Error(`API failed with response code ${response.getResponseCode()}: ${response.getContentText()}`);
    }

    // 3. Save raw returned PDF Binary Stream into target Google Drive Folder
    const pdfBlob = response.getBlob().setName(`Invoice_${invoiceNum}_${bRecord.guestName.replace(/\s+/g, '_')}.pdf`);
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const pdfFile = targetFolder.createFile(pdfBlob);

    return { 
      success: true, 
      message: 'Checkout successful! Bill generated via API & archived to Drive folder.', 
      pdfUrl: pdfFile.getUrl() 
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// 7. EXPENSES LEDGER LOGIC
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
    
    return { success: true, message: 'Expense tracked successfully!' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// 8. METRICS AGGREGATIONS
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

// 9. SETTINGS OPERATIONS
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
    return { success: true, message: 'Settings saved successfully!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function fetchSettings() {
  return {
    INVOICE_GENERATOR_API_URL: getSetting('INVOICE_GENERATOR_API_URL'),
    INVOICES_FOLDER_ID: getSetting('INVOICES_FOLDER_ID'),
    HOTEL_BILLING_NAME: getSetting('HOTEL_BILLING_NAME'),
    HOTEL_UPI_ID: getSetting('HOTEL_UPI_ID')
  };
}