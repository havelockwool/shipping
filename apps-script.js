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
      return HtmlService.createHtmlOutput(`
        <html>
          <body>
            <h2>Access Denied</h2>
            <p>You must sign in with a @havelockwool.com account.</p>
            <p>Please close this window and try again.</p>
            <script>
              if (window.opener) {
                window.opener.postMessage({
                  success: false,
                  error: 'Must use @havelockwool.com account'
                }, '*');
              }
            </script>
          </body>
        </html>
      `);
    }

    Logger.log('Authenticated user: ' + userEmail);

    // Parse the incoming JSON data from form parameter
    const data = JSON.parse(e.parameter.data);

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

    // Create HTML table for easy copying
    let tableHtml = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 12px;">';

    // Add headers
    tableHtml += '<thead><tr style="background-color: #f0f0f0;">';
    headers.forEach(header => {
      tableHtml += `<th style="padding: 8px; text-align: left;">${header}</th>`;
    });
    tableHtml += '</tr></thead>';

    // Add data rows
    tableHtml += '<tbody>';
    rows.forEach(row => {
      tableHtml += '<tr>';
      row.forEach(cell => {
        tableHtml += `<td style="padding: 8px; border: 1px solid #ddd;">${cell}</td>`;
      });
      tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';

    // Return success HTML with post message and copyable table
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .status { margin-bottom: 20px; }
            .copy-btn {
              background-color: #4CAF50;
              color: white;
              padding: 10px 20px;
              border: none;
              cursor: pointer;
              font-size: 14px;
              margin: 10px 0;
            }
            .copy-btn:hover { background-color: #45a049; }
            .table-container {
              max-height: 400px;
              overflow: auto;
              border: 1px solid #ddd;
              margin: 10px 0;
            }
          </style>
        </head>
        <body>
          <div class="status">
            <h2>✓ Success!</h2>
            <p>Successfully imported ${rows.length} order(s) to IMPORT sheet</p>
            <p>Authenticated as: ${userEmail}</p>
          </div>

          <button class="copy-btn" onclick="copyTable()">📋 Copy Table to Clipboard</button>

          <div class="table-container" id="tableContainer">
            ${tableHtml}
          </div>

          <p style="color: #666; font-size: 12px;">
            Click "Copy Table" above, then paste (Ctrl+V) into Google Sheets
          </p>

          <script>
            function copyTable() {
              const table = document.querySelector('table');
              const range = document.createRange();
              range.selectNode(table);
              window.getSelection().removeAllRanges();
              window.getSelection().addRange(range);
              document.execCommand('copy');
              window.getSelection().removeAllRanges();

              const btn = document.querySelector('.copy-btn');
              btn.textContent = '✓ Copied!';
              btn.style.backgroundColor = '#2196F3';

              setTimeout(() => {
                btn.textContent = '📋 Copy Table to Clipboard';
                btn.style.backgroundColor = '#4CAF50';
              }, 2000);
            }

            if (window.opener) {
              window.opener.postMessage({
                success: true,
                rowsAdded: ${rows.length}
              }, '*');
            }
          </script>
        </body>
      </html>
    `);

  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h2>Error</h2>
          <p>${error.toString()}</p>
          <script>
            if (window.opener) {
              window.opener.postMessage({
                success: false,
                error: '${error.toString()}'
              }, '*');
            }
          </script>
        </body>
      </html>
    `);
  }
}

/**
 * Handle GET requests (for testing/status)
 */
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();

  return HtmlService.createHtmlOutput(`
    <html>
      <body>
        <h2>Invoice Import API</h2>
        <p>Status: Running</p>
        <p>Authenticated as: ${userEmail || 'Not signed in'}</p>
        <p>Use POST to import data.</p>
      </body>
    </html>
  `);
}
