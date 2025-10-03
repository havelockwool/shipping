/**
 * Show copyable table in popup
 */
function showCopyableTable() {
    if (invoiceData.length === 0) {
        showStatus('No data to copy. Please upload a PDF first.', 'error');
        return;
    }

    // Prepare headers
    const headers = ['Page', 'Date', 'Cust Order #', 'PO Number', 'Customer Name',
                     'Ship To Name', 'Customer Address', 'Ship To Address', 'Phone',
                     'Address Type', 'Model Number', 'Internet Num', 'Qty Shipped'];

    // Prepare rows
    const rows = invoiceData.map(order => [
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

    // Create HTML table
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

    // Create popup HTML
    const popupHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body { font-family: Arial, sans-serif; padding: 20px; }
                .header { margin-bottom: 20px; }
                .copy-btn {
                    background-color: #4CAF50;
                    color: white;
                    padding: 12px 24px;
                    border: none;
                    cursor: pointer;
                    font-size: 16px;
                    margin: 10px 0;
                    border-radius: 4px;
                }
                .copy-btn:hover { background-color: #45a049; }
                .table-container {
                    max-height: 500px;
                    overflow: auto;
                    border: 1px solid #ddd;
                    margin: 10px 0;
                }
                .instructions {
                    background-color: #e3f2fd;
                    padding: 12px;
                    border-radius: 4px;
                    margin: 10px 0;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h2>📋 Copy Data to Google Sheets</h2>
                <p>${rows.length} order(s) ready to copy</p>
            </div>

            <button class="copy-btn" onclick="copyTable()">📋 Copy Table to Clipboard</button>

            <div class="instructions">
                <strong>How to paste into Google Sheets:</strong>
                <ol style="margin: 8px 0; padding-left: 20px;">
                    <li>Click "Copy Table to Clipboard" button above</li>
                    <li>Open your Google Sheet</li>
                    <li>Click on cell A1 (or wherever you want the data)</li>
                    <li>Press Ctrl+V (or Cmd+V on Mac) to paste</li>
                </ol>
            </div>

            <div class="table-container">
                ${tableHtml}
            </div>

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
                    btn.textContent = '✓ Copied to Clipboard!';
                    btn.style.backgroundColor = '#2196F3';

                    setTimeout(() => {
                        btn.textContent = '📋 Copy Table to Clipboard';
                        btn.style.backgroundColor = '#4CAF50';
                    }, 2000);
                }
            </script>
        </body>
        </html>
    `;

    // Open popup
    const popup = window.open('', 'CopyableData', 'width=900,height=700,menubar=no,toolbar=no,location=no');

    if (!popup) {
        showStatus('❌ Popup blocked. Please allow popups for this site.', 'error');
        return;
    }

    popup.document.write(popupHtml);
    popup.document.close();

    showStatus('✓ Table opened in new window. Click "Copy Table" to copy data.', 'success');
}

/**
 * Initialize copy table button
 */
function initCopyTableButton() {
    const btn = document.getElementById('saveSheetBtn');
    if (!btn) {
        console.error('Copy table button not found');
        return;
    }

    btn.addEventListener('click', () => {
        console.log('Copy table clicked');
        showCopyableTable();
    });

    btn.style.opacity = '1';
    btn.title = 'Show copyable table';
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initCopyTableButton);
} else {
    initCopyTableButton();
}
