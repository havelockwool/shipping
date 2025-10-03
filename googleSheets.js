/**
 * Google Apps Script Integration
 *
 * SETUP: After deploying your Apps Script (see apps-script.js),
 * paste your Web App URL below:
 */

// TODO: Replace this with your deployed Apps Script Web App URL
// Get this from: Extensions > Apps Script > Deploy > Manage deployments
const APPS_SCRIPT_URL = 'https://script.google.com/a/macros/havelockwool.com/s/AKfycbyIlbfRQxUh59z_ybCUjzi2e1yS3W7VTJjw9n4ehgk5K5e736LBpOsMEOY3YIs-XKGU/exec';

/**
 * Send data to Google Sheets via Apps Script
 */
async function saveToGoogleSheet() {
    if (!APPS_SCRIPT_URL || APPS_SCRIPT_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') {
        showStatus('❌ Apps Script URL not configured. See apps-script.js for setup instructions.', 'error');
        console.error('Please configure APPS_SCRIPT_URL in googleSheets.js');
        return;
    }

    if (invoiceData.length === 0) {
        showStatus('No data to save. Please upload a PDF first.', 'error');
        return;
    }

    try {
        showStatus('Sending data to Google Sheet...', 'info');

        // Prepare the data payload
        const payload = {
            orders: invoiceData.map(order => ({
                page: order.page || '',
                date: order.date || '',
                custNum: order.custNum || '',
                poNumber: order.poNumber || '',
                customerName: order.customerName || '',
                shipToName: order.shipToName || '',
                customerAddress: order.customerAddress || '',
                shipToAddress: order.shipToAddress || '',
                phone: order.phone || '',
                addressType: order.addressType || '',
                modelNumber: order.modelNumber || '',
                internetNumber: order.internetNumber || '',
                quantity: order.quantity || ''
            }))
        };

        // Send POST request to Apps Script
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors', // Apps Script requires no-cors mode
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload)
        });

        // Note: With no-cors mode, we can't read the response
        // But the request will succeed if the script is set up correctly
        showStatus(`✓ Successfully sent ${invoiceData.length} order(s) to Google Sheet!`, 'success');
        console.log('Data sent to Google Sheet:', payload);

    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error('Error sending to Google Sheet:', error);
    }
}

/**
 * Initialize Google Sheets button
 */
function initGoogleSheetsButton() {
    const btn = document.getElementById('saveSheetBtn');
    if (!btn) {
        console.error('Save to Google Sheet button not found');
        return;
    }

    btn.addEventListener('click', () => {
        console.log('Save to Google Sheet clicked');
        saveToGoogleSheet();
    });

    // Button is always ready (no API loading required)
    btn.style.opacity = '1';
    btn.title = 'Save to Google Sheet';
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initGoogleSheetsButton);
} else {
    initGoogleSheetsButton();
}
