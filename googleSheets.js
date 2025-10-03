/**
 * Google Sheets Integration for Workspace
 *
 * Opens Apps Script in a popup for authentication with @havelockwool.com
 * The Apps Script is set to "Anyone within havelockwool.com"
 */

const APPS_SCRIPT_URL = 'https://script.google.com/a/macros/havelockwool.com/s/AKfycbyIlbfRQxUh59z_ybCUjzi2e1yS3W7VTJjw9n4ehgk5K5e736LBpOsMEOY3YIs-XKGU/exec';

/**
 * Send data to Google Sheets by opening popup
 * This works with "Anyone within havelockwool.com" setting
 */
async function saveToGoogleSheet() {
    if (invoiceData.length === 0) {
        showStatus('No data to save. Please upload a PDF first.', 'error');
        return;
    }

    try {
        showStatus('Opening authentication window...', 'info');

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

        // Encode data as URL parameter
        const encodedData = encodeURIComponent(JSON.stringify(payload));
        const url = `${APPS_SCRIPT_URL}?data=${encodedData}`;

        // Open in popup window
        const popup = window.open(
            url,
            'GoogleSheetsAuth',
            'width=600,height=700,menubar=no,toolbar=no,location=no'
        );

        if (!popup) {
            showStatus('❌ Popup blocked. Please allow popups for this site.', 'error');
            return;
        }

        showStatus('✓ Please sign in with @havelockwool.com in the popup window...', 'info');

        // Listen for popup messages
        window.addEventListener('message', function handleMessage(event) {
            // Security: verify origin
            if (event.origin !== 'https://script.google.com' && event.origin !== 'https://script.googleusercontent.com') {
                return;
            }

            const data = event.data;

            if (data.success) {
                showStatus(`✓ Successfully saved ${invoiceData.length} order(s)!`, 'success');
                popup.close();
            } else if (data.error) {
                showStatus(`Error: ${data.error}`, 'error');
            }

            window.removeEventListener('message', handleMessage);
        });

    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error('Error:', error);
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

    btn.style.opacity = '1';
    btn.title = 'Save to Google Sheet (Requires @havelockwool.com account)';
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initGoogleSheetsButton);
} else {
    initGoogleSheetsButton();
}
