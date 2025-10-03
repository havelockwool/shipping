/**
 * Google Sheets Integration with OAuth Authentication
 *
 * This requires users to sign in with their @havelockwool.com Google account.
 * Works with your Apps Script set to "Anyone within havelockwool.com"
 *
 * SETUP:
 * 1. Go to Google Cloud Console: https://console.cloud.google.com
 * 2. Enable Google+ API (for user info)
 * 3. Your OAuth Client ID is already set up
 * 4. Make sure authorized JavaScript origins includes: https://havelockwool.github.io
 */

const GOOGLE_CLIENT_ID = '504638897346-hqf8mdu7rlrp8g9adalmadi7q79s77e2.apps.googleusercontent.com';
const APPS_SCRIPT_URL = 'https://script.google.com/a/macros/havelockwool.com/s/AKfycbyIlbfRQxUh59z_ybCUjzi2e1yS3W7VTJjw9n4ehgk5K5e736LBpOsMEOY3YIs-XKGU/exec';

let tokenClient;
let accessToken = null;
let userEmail = null;

/**
 * Initialize Google Identity Services
 */
function initGoogleAuth() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: 'https://www.googleapis.com/auth/userinfo.email',
        callback: (response) => {
            if (response.error) {
                showStatus(`Auth error: ${response.error}`, 'error');
                return;
            }
            accessToken = response.access_token;
            getUserInfo();
        },
    });
}

/**
 * Get user email to verify domain
 */
async function getUserInfo() {
    try {
        const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        const data = await response.json();
        userEmail = data.email;

        // Verify user is from havelockwool.com
        if (!userEmail.endsWith('@havelockwool.com')) {
            showStatus('❌ Access denied. Must use @havelockwool.com account.', 'error');
            accessToken = null;
            userEmail = null;
            return;
        }

        showStatus(`✓ Signed in as ${userEmail}`, 'success');
        // Now send the data
        await sendToGoogleSheet();

    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error('Error getting user info:', error);
    }
}

/**
 * Trigger sign-in flow
 */
function signIn() {
    tokenClient.requestAccessToken({ prompt: 'select_account' });
}

/**
 * Send data to Google Sheets
 */
async function sendToGoogleSheet() {
    try {
        showStatus('Sending data to Google Sheet...', 'info');

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

        // Send with access token
        const response = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload)
        });

        const result = await response.json();

        if (result.success) {
            showStatus(`✓ Successfully saved ${invoiceData.length} order(s)!`, 'success');
        } else {
            showStatus(`Error: ${result.error}`, 'error');
        }

    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error('Error sending to Google Sheet:', error);
    }
}

/**
 * Main function when Save button is clicked
 */
async function saveToGoogleSheet() {
    if (invoiceData.length === 0) {
        showStatus('No data to save. Please upload a PDF first.', 'error');
        return;
    }

    // Check if already authenticated
    if (accessToken && userEmail) {
        await sendToGoogleSheet();
    } else {
        // Trigger sign-in
        signIn();
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
    btn.title = 'Save to Google Sheet (Sign in with @havelockwool.com)';
}

// Initialize when DOM and Google scripts are ready
function onGoogleScriptsLoaded() {
    initGoogleAuth();
    initGoogleSheetsButton();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        if (typeof google !== 'undefined') {
            onGoogleScriptsLoaded();
        }
    });
} else {
    if (typeof google !== 'undefined') {
        onGoogleScriptsLoaded();
    }
}

// Expose for script loading
window.gisLoaded = function() {
    console.log('✓ Google Identity Services loaded');
    if (document.readyState !== 'loading') {
        onGoogleScriptsLoaded();
    }
};
