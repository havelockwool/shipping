/**
 * Google Apps Script for Home Depot Invoice Processor
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1tYK9vcHh2Ry8xXL9Iuym6RZfwn4_eL0SpSWiDV28wl8/edit
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code in Code.gs
 * 4. Copy and paste this entire file into Code.gs
 * 5. Click Save (disk icon)
 * 6. Click Deploy > New deployment
 * 7. Click "Select type" > Web app
 * 8. Settings:
 *    - Description: "Invoice Data Import"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone within havelockwool.com"
 * 9. Click Deploy
 * 10. Copy the Web app URL (looks like: https://script.google.com/macros/s/...../exec)
 * 11. Paste that URL into your frontend code where it says APPS_SCRIPT_URL
 *
 * SECURITY NOTES:
 * - This script only accepts POST requests with specific data structure
 * - You can optionally add origin checking (see commented code below)
 * - No credentials are exposed in your frontend code
 * - The deployment URL is safe to be public (it only appends data to your sheet)
 */

/**
 * Handle POST requests from the web application
 */
function doPost(e) {
  try {
    // Verify user is authenticated with @havelockwool.com account
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail || !userEmail.endsWith('@havelockwool.com')) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Unauthorized: Must be signed in with @havelockwool.com account'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    Logger.log('Authenticated user: ' + userEmail);

    // Parse the incoming JSON data
    const data = JSON.parse(e.postData.contents);

    // Validate data structure
    if (!data.orders || !Array.isArray(data.orders)) {
      throw new Error('Invalid data structure: orders array required');
    }

    // Get the spreadsheet and IMPORT sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IMPORT');

    if (!sheet) {
      throw new Error('IMPORT sheet not found');
    }

    // Prepare headers
    const headers = [
      'Page', 'Date', 'Cust Order #', 'PO Number', 'Customer Name',
      'Ship To Name', 'Customer Address', 'Ship To Address', 'Phone',
      'Address Type', 'Model Number', 'Internet Num', 'Qty Shipped'
    ];

    // Prepare rows from the order data
    const rows = data.orders.map(order => [
      order.page || '',
      order.date || '',
      order.custNum || '',
      order.poNumber || '',
      order.customerName || '',
      order.shipToName || '',
      order.customerAddress || '',
      order.shipToAddress || '',
      order.phone || '',
      order.addressType || '',
      order.modelNumber || '',
      order.internetNumber || '',
      order.quantity || ''
    ]);

    // Check if we should add headers (if sheet is empty or first row is not headers)
    const lastRow = sheet.getLastRow();
    let shouldAddHeaders = false;

    if (lastRow === 0) {
      shouldAddHeaders = true;
    } else {
      // Check if first row contains headers
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      const hasHeaders = firstRow[0] === 'Page' && firstRow[2] === 'Cust Order #';
      if (!hasHeaders) {
        shouldAddHeaders = true;
      }
    }

    // Append data to sheet
    if (shouldAddHeaders) {
      sheet.appendRow(headers);
    }

    // Append all rows
    rows.forEach(row => {
      sheet.appendRow(row);
    });

    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: `Successfully imported ${rows.length} order(s)`,
      rowsAdded: rows.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Return error response
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle GET requests (optional - for testing)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'OK',
    message: 'Invoice Import API is running. Use POST to send data.',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}
