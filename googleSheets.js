// Google Sheets API configuration
const GOOGLE_CLIENT_ID = '.apps.googleusercontent.com';
const GOOGLE_API_KEY = '_';
const GOOGLE_SHEET_ID = '1tYK9vcHh2Ry8xXL9Iuym6RZfwn4_eL0SpSWiDV28wl8'; // Your Google Sheet ID
const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

let tokenClient;
let gapiInited = false;
let gisInited = false;
let isAuthorized = false;

// Initialize Google API - call when page loads
window.gapiLoaded = function() {
    console.log('✓ Google API script loaded');
    gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
    try {
        await gapi.client.init({
            apiKey: GOOGLE_API_KEY,
            discoveryDocs: [DISCOVERY_DOC],
        });
        gapiInited = true;
        console.log('✓ Google API client initialized');
        maybeEnableButtons();
    } catch (error) {
        console.error('Error initializing Google API client:', error);
    }
}

window.gisLoaded = function() {
    console.log('✓ Google Identity Services loaded');
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: SCOPES,
        callback: '', // defined later
    });
    gisInited = true;
    maybeEnableButtons();
}

function maybeEnableButtons() {
    if (gapiInited && gisInited) {
        console.log('✅ Google Sheets API fully ready!');
        const btn = document.getElementById('saveSheetBtn');
        if (btn) {
            btn.style.opacity = '1';
            btn.title = 'Save to Google Sheet';
        }
    }
}

// Add timeout warning
setTimeout(() => {
    if (!gapiInited || !gisInited) {
        console.warn('⚠️ Google Sheets API taking longer than expected to load...');
        console.log('Status: gapiInited=' + gapiInited + ', gisInited=' + gisInited);

        // Check if scripts loaded
        if (typeof gapi === 'undefined') {
            console.error('❌ Google API script failed to load. Check network/firewall.');
        }
        if (typeof google === 'undefined') {
            console.error('❌ Google Identity Services script failed to load. Check network/firewall.');
        }

        // Show user-friendly message
        const btn = document.getElementById('saveSheetBtn');
        if (btn) {
            btn.title = 'Google Sheets API failed to load. Check console for details.';
            btn.style.opacity = '0.3';
        }
    }
}, 5000);

// Also check after 10 seconds and provide actionable feedback
setTimeout(() => {
    if (!gapiInited || !gisInited) {
        console.error('❌ Google Sheets API failed to initialize after 10 seconds.');
        console.log('Troubleshooting:');
        console.log('1. Check if you are accessing from an authorized domain in Google Cloud Console');
        console.log('2. Current page URL:', window.location.href);
        console.log('3. Disable ad blockers or browser extensions');
        console.log('4. Check browser console for CORS or network errors');

        if (!gapiInited) console.log('   - Google API (gapi) failed to initialize');
        if (!gisInited) console.log('   - Google Identity Services (gis) failed to initialize');
    }
}, 10000);

// Handle authorization
function handleAuthClick() {
    tokenClient.callback = async (resp) => {
        if (resp.error !== undefined) {
            showStatus(`Auth error: ${resp.error}`, 'error');
            return;
        }
        isAuthorized = true;
        await appendToGoogleSheet();
    };

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({prompt: 'consent'});
    } else {
        tokenClient.requestAccessToken({prompt: ''});
    }
}

// Append data to Google Sheet
async function appendToGoogleSheet() {
    try {
        showStatus('Sending data to Google Sheet...', 'info');

        // Prepare data rows
        const headers = ['Page', 'Date', 'Cust Order #', 'PO Number', 'Customer Name', 'Ship To Name', 'Customer Address', 'Ship To Address', 'Phone', 'Address Type', 'Model Number', 'Internet Num', 'Qty Shipped'];

        const rows = invoiceData.map(order => [
            //order.page,
            order.date,
            order.custNum,
            order.poNumber,
            order.customerName,
            order.shipToName,
            order.customerAddress,
            order.shipToAddress,
            order.phone,
            order.addressType,
            order.modelNumber,
            order.internetNumber,
            order.quantity
        ]);

        // Append to sheet (includes headers + data)
        const response = await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: GOOGLE_SHEET_ID,
            range: 'IMPORT!A:M',
            valueInputOption: 'USER_ENTERED',
            resource: {
                //values: [headers, ...rows]
                  values: [...rows]
            },
        });

        showStatus(`✓ Successfully added ${rows.length} orders to Google Sheet!`, 'success');
        console.log('Google Sheets response:', response);
    } catch (err) {
        showStatus(`Error: ${err.message}`, 'error');
        console.error('Google Sheets error:', err);
    }
}

// Initialize Google Sheets button
function initGoogleSheetsButton() {
    const btn = document.getElementById('saveSheetBtn');
    if (!btn) {
        console.error('Save to Google Sheet button not found');
        return;
    }

    // Dim button until API loads
    btn.style.opacity = '0.5';
    btn.title = 'Loading Google Sheets API...';

    btn.addEventListener('click', () => {
        console.log('Save to Google Sheet clicked');

        if (invoiceData.length === 0) {
            showStatus('No data to save. Please upload a PDF first.', 'error');
            return;
        }

        if (!gapiInited || !gisInited) {
            showStatus('Google Sheets API is still loading. Please wait or refresh the page.', 'info');
            console.log('API Status: gapiInited=' + gapiInited + ', gisInited=' + gisInited);
            return;
        }

        if (!isAuthorized) {
            console.log('Starting authorization flow...');
            handleAuthClick();
        } else {
            console.log('Already authorized, appending to sheet...');
            appendToGoogleSheet();
        }
    });
}

// Call this after DOM loads
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initGoogleSheetsButton);
} else {
    initGoogleSheetsButton();
}
