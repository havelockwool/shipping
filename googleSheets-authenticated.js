/**
 * Google Apps Script Integration with Password Protection
 *
 * This version requires users to enter a password before sending data to Google Sheets.
 * Only share the password with havelockwool.com employees.
 */

const APPS_SCRIPT_URL = 'https://script.google.com/a/macros/havelockwool.com/s/AKfycbyIlbfRQxUh59z_ybCUjzi2e1yS3W7VTJjw9n4ehgk5K5e736LBpOsMEOY3YIs-XKGU/exec';

// Simple password hash (SHA-256)
// To change password: use an online SHA-256 generator and replace this hash
const PASSWORD_HASH = 'ef92b778bafe771e89245b89ecbc08a44a4e166c06659911881f383d4473e94f'; // "password123"

let isAuthenticated = false;

/**
 * Verify password
 */
async function verifyPassword(password) {
    const encoder = new TextEncoder();
    const data = encoder.encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex === PASSWORD_HASH;
}

/**
 * Show password prompt
 */
function promptForPassword() {
    return new Promise((resolve) => {
        const password = prompt('Enter password to save to Google Sheet:');
        if (!password) {
            resolve(false);
            return;
        }

        verifyPassword(password).then(isValid => {
            if (isValid) {
                isAuthenticated = true;
                showStatus('✓ Password correct', 'success');
                resolve(true);
            } else {
                showStatus('❌ Incorrect password', 'error');
                resolve(false);
            }
        });
    });
}

/**
 * Send data to Google Sheets via Apps Script
 */
async function saveToGoogleSheet() {
    if (invoiceData.length === 0) {
        showStatus('No data to save. Please upload a PDF first.', 'error');
        return;
    }

    // Check authentication
    if (!isAuthenticated) {
        const authenticated = await promptForPassword();
        if (!authenticated) return;
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
            mode: 'no-cors',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload)
        });

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

    btn.style.opacity = '1';
    btn.title = 'Save to Google Sheet (Password Required)';
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initGoogleSheetsButton);
} else {
    initGoogleSheetsButton();
}

/**
 * TO CHANGE PASSWORD:
 * 1. Go to https://emn178.github.io/online-tools/sha256.html
 * 2. Enter your new password
 * 3. Copy the hash output
 * 4. Replace PASSWORD_HASH value above with the new hash
 */
