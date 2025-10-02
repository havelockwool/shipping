// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

let invoiceData = [];
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const summary = document.getElementById('summary');
const tableContainer = document.getElementById('tableContainer');
const dataTable = document.getElementById('dataTable');
const actions = document.getElementById('actions');

// Drop zone events
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
        handleFile(file);
    } else {
        showStatus('Please drop a PDF file', 'error');
    }
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
});

async function handleFile(file) {
    showStatus(`Processing ${file.name}...`, 'info');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

        invoiceData = [];

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            // Extract order data from Home Depot invoice using coordinate-based parsing
            const order = await extractOrderFromText(page, i);
            invoiceData.push(order);
        }

        document.getElementById('orderCount').textContent = invoiceData.length;
        document.getElementById('pageCount').textContent = pdf.numPages;

        displayData();
        showStatus(`Successfully extracted ${invoiceData.length} orders from ${pdf.numPages} pages`, 'success');

        summary.style.display = 'block';
        tableContainer.style.display = 'block';
        actions.style.display = 'flex';

    } catch (error) {
        showStatus(`Error processing PDF: ${error.message}`, 'error');
        console.error(error);
    }
}

async function extractOrderFromText(page, pageNum) {
    const textContent = await page.getTextContent();
    const items = textContent.items;

    // Group items into lines using Y-coordinate
    const tolerance = 5; // pixels for line grouping
    items.sort((a, b) => a.transform[5] - b.transform[5]); // sort by Y ascending

    const lines = [];
    let currentLine = [];

    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const y = item.transform[5];

        if (currentLine.length === 0 || Math.abs(y - currentLine[0].transform[5]) < tolerance) {
            currentLine.push(item);
        } else {
            // Sort current line by X-coordinate
            lines.push(currentLine.sort((a, b) => a.transform[4] - b.transform[4]));
            currentLine = [item];
        }
    }
    if (currentLine.length) {
        lines.push(currentLine.sort((a, b) => a.transform[4] - b.transform[4]));
    }

    // Convert lines to strings
    const lineStrings = lines.map(line =>
        line.map(item => item.str).join(' ').trim()
    );

    console.log(`Page ${pageNum} lines:`);
    lineStrings.forEach((line, index) => {
        console.log(`Line ${index}: "${line}"`);
    });

    // Extract page number from "Page: X" line
    let actualPageNum = pageNum; // default to loop value
    const pageLine = lineStrings.find(l => l.includes('Page:'));
    if (pageLine) {
        const match = pageLine.match(/Page:\s*(\d+)/);
        if (match) actualPageNum = parseInt(match[1], 10);
    }

    // Extract data from specific lines
    let custNum = `ORDER-${actualPageNum}`;
    let poNumber = '';
    let orderDate = '';
    let customerName = '';
    let customerAddress = '';
    let address = '';
    let city = '';
    let state = '';
    let zip = '';
    let phone = '';
    let addressType = 'Residential';
    let shipVia = 'Misc. Common Carrier';
    let modelNumber = '';
    let description = '';
    let quantity = '1';
    let internetNumber = '';

    // Line 1: Extract PO#, Customer Order #, and Customer Name
    const headerLine = lineStrings.find(l => l.includes('PO #') && l.includes('Customer Order #:') && l.includes('Customer Name:'));
    if (headerLine) {
        const poMatch = headerLine.match(/PO #\s*(\S+)/);
        if (poMatch) poNumber = poMatch[1];

        const custMatch = headerLine.match(/Customer Order #:\s*([A-Z0-9-]+)/);
        if (custMatch) custNum = custMatch[1];

        const nameMatch = headerLine.match(/Customer Name:\s*(.+)$/);
        if (nameMatch) customerName = nameMatch[1].trim();
    }

    // Extract Ship To information from Line 2
    const shipToLine = lineStrings.find(l => l.includes('Ship To:'));
    if (shipToLine) {
        // Extract name - between "Ship To:" and first number
        const nameMatch = shipToLine.match(/Ship To:\s+(.+?)\s+\d+/);
        if (nameMatch) customerAddress = nameMatch[1].trim();

        // Extract street address - starts with number
        const streetMatch = shipToLine.match(/(\d+\s+[A-Za-z\s]+(?:Way|Rd|St|Street|Avenue|Ave|Court|Drive|Dr|Lane|Ln|Boulevard|Blvd|Circle|Cir)[^\d]*(?:Ste|Suite|Apt|Unit)?\s*\d*)/i);
        if (streetMatch) address = streetMatch[1].trim();

        // Extract city, state, zip - look for pattern "City, ST Zip"
        const cityMatch = shipToLine.match(/([A-Za-z\s]+),\s*([A-Z]{2})\s+(\d{5})/);
        if (cityMatch) {
            city = cityMatch[1].trim();
            state = cityMatch[2].trim();
            zip = cityMatch[3].trim();
        }

        // Extract phone - pattern (###) ###-####
        const phoneMatch = shipToLine.match(/\((\d{3})\)\s*(\d{3})-(\d{4})/);
        if (phoneMatch) phone = `(${phoneMatch[1]}) ${phoneMatch[2]}-${phoneMatch[3]}`;
    }

    // Line 7: Model Number
    const modelLine = lineStrings.find((l, idx) => idx > 5 && l.trim().startsWith('Model Number') && l.length < 30);
    if (modelLine) {
        const modelMatch = modelLine.match(/Model Number\s+(.+)$/);
        if (modelMatch) modelNumber = modelMatch[1].trim();
    }

    // Line 9: Description
    const descLine = lineStrings.find(l => l.includes('R-22') || (l.includes('Insulation') && l.length > 20));
    if (descLine) {
        description = descLine.trim();
    }

    // Line 15: Date
    const dateMatch = lineStrings.find(l => /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(l.trim()));
    if (dateMatch) {
        orderDate = dateMatch.trim();
    }

    // Line 17: Ship Via
    const shipViaLine = lineStrings.find(l => l.includes('Misc. Common Carrier'));
    if (shipViaLine) {
        if (shipViaLine.includes('Misc. Common Carrier')) {
            shipVia = 'Misc. Common Carrier';
        }
    }

    // Line 19: Address Type
    const addrTypeLine = lineStrings.find(l => l.trim() === 'Residential' || l.trim() === 'Commercial');
    if (addrTypeLine) {
        addressType = addrTypeLine.trim();
    }

    // Line 24: Quantity
    const qtyLine = lineStrings.findIndex(l => l.trim() === 'Qty Shipped');
    if (qtyLine !== -1 && lineStrings[qtyLine + 1]) {
        const qtyMatch = lineStrings[qtyLine + 1].trim();
        if (/^\d+$/.test(qtyMatch)) {
            quantity = qtyMatch;
        }
    }

    // Extract Internet Number (appears on multiple lines)
    const internetLine = lineStrings.find(l => l.includes('Internet Number'));
    if (internetLine) {
        const internetMatch = internetLine.match(/Internet Number\s+(\d+)/);
        if (internetMatch) internetNumber = internetMatch[1];
    }

    const fullAddress = [address, city, state, zip].filter(Boolean).join(', ');

    console.log(`Page ${pageNum} extracted:`, {
        custNum, poNumber, orderDate, customerName, customerAddress, fullAddress, phone, addressType, modelNumber, internetNumber, quantity
    });

    return {
        custNum: custNum,
        poNumber: poNumber,
        date: orderDate,
        page: actualPageNum,
        customerName: customerName,
        customerAddress: customerAddress,
        address: fullAddress,
        street: address,
        city: city,
        state: state,
        zip: zip,
        phone: phone,
        addressType: addressType,
        modelNumber: modelNumber,
        internetNumber: internetNumber,
        quantity: quantity
    };
}

function displayData() {
    dataTable.innerHTML = '';

    invoiceData.forEach((order, index) => {
        const row = document.createElement('tr');
        // Order matches HTML headers: Page, Date, Cust Order #, PO Number, Customer Name, Ship To Name, Ship To Address, Phone, Address Type, Model Number, Internet Num, Qty Shipped
        row.innerHTML = `
            <td><input type="text" value="${order.page}" data-index="${index}" data-field="page"></td>
            <td><input type="text" value="${order.date}" data-index="${index}" data-field="date"></td>
            <td><input type="text" value="${order.custNum}" data-index="${index}" data-field="custNum"></td>
            <td><input type="text" value="${order.poNumber}" data-index="${index}" data-field="poNumber"></td>
            <td><input type="text" value="${order.customerName}" data-index="${index}" data-field="customerName"></td>
            <td><input type="text" value="${order.customerAddress}" data-index="${index}" data-field="customerAddress"></td>
            <td><input type="text" value="${order.address}" data-index="${index}" data-field="address"></td>
            <td><input type="text" value="${order.phone}" data-index="${index}" data-field="phone"></td>
            <td><input type="text" value="${order.addressType}" data-index="${index}" data-field="addressType"></td>
            <td><input type="text" value="${order.modelNumber}" data-index="${index}" data-field="modelNumber"></td>
            <td><input type="text" value="${order.internetNumber}" data-index="${index}" data-field="internetNumber"></td>
            <td><input type="text" value="${order.quantity}" data-index="${index}" data-field="quantity"></td>
        `;
        dataTable.appendChild(row);
    });

    // Add event listeners for edits
    document.querySelectorAll('#dataTable input').forEach(input => {
        input.addEventListener('change', (e) => {
            const index = parseInt(e.target.dataset.index);
            const field = e.target.dataset.field;
            invoiceData[index][field] = e.target.value;
        });
    });
}

function showStatus(message, type) {
    status.textContent = message;
    status.className = `status ${type}`;
    status.style.display = 'block';

    if (type === 'success') {
        setTimeout(() => {
            status.style.display = 'none';
        }, 5000);
    }
}

// Button actions
document.getElementById('downloadCsvBtn').addEventListener('click', () => {
    // Prepare data for CSV with all fields
    const exportData = invoiceData.map(order => ({
        'Page': order.page,
        'Date': order.date,
        'Cust Order #': order.custNum,
        'PO Number': order.poNumber,
        'Customer Name': order.customerName,
        'Ship To Name': order.customerAddress,
        'Ship To Address': order.address,
        'Phone': order.phone,
        'Address Type': order.addressType,
        'Model Number': order.modelNumber,
        'Internet Num': order.internetNumber,
        'Qty Shipped': order.quantity
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    XLSX.utils.book_append_sheet(wb, ws, 'Orders');

    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `home_depot_orders_${date}.csv`);

    showStatus('CSV file downloaded successfully!', 'success');
});

document.getElementById('saveSheetBtn').addEventListener('click', () => {
    // TODO: Implement Google Sheets API integration
    showStatus('Google Sheets integration coming soon! For now, use the CSV download.', 'info');

    // Log data for debugging
    console.log('Data ready for Google Sheets:', invoiceData);
});

document.getElementById('resetBtn').addEventListener('click', () => {
    invoiceData = [];
    dataTable.innerHTML = '';
    tableContainer.style.display = 'none';
    actions.style.display = 'none';
    summary.style.display = 'none';
    status.style.display = 'none';
    fileInput.value = '';
});


