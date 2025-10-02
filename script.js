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
            const textContent = await page.getTextContent();
            const text = textContent.items.map(item => item.str).join(' ');
            
            // Simple extraction logic - you'll customize this based on your PDF structure
            const order = extractOrderFromText(text, i);
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

function extractOrderFromText(text, pageNum) {
    // DEMO LOGIC - This is where you'll add your specific extraction rules
    // For now, we'll create sample data
    
    // Look for common patterns (customize based on your PDF)
    const addressMatch = text.match(/(\d+\s+[\w\s]+(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Boulevard|Blvd))/i);
    const typeMatch = text.match(/(Ground|Air|Express|Standard|Overnight)/i);
    
    return {
        orderNum: pageNum,
        page: pageNum,
        address: addressMatch ? addressMatch[0] : `Sample Address ${pageNum}`,
        shipmentType: typeMatch ? typeMatch[0] : 'Ground',
        carrier: text.includes('FedEx') ? 'FedEx' : text.includes('UPS') ? 'UPS' : 'FedEx'
    };
}

function displayData() {
    dataTable.innerHTML = '';
    
    invoiceData.forEach((order, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${order.orderNum}</td>
            <td>${order.page}</td>
            <td><input type="text" value="${order.address}" data-index="${index}" data-field="address"></td>
            <td><input type="text" value="${order.shipmentType}" data-index="${index}" data-field="shipmentType"></td>
            <td><input type="text" value="${order.carrier}" data-index="${index}" data-field="carrier"></td>
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
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(invoiceData);
    XLSX.utils.book_append_sheet(wb, ws, 'Orders');
    
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `invoices_${date}.csv`);
    
    showStatus('CSV file downloaded successfully!', 'success');
});

document.getElementById('saveSheetBtn').addEventListener('click', () => {
    // TODO: Implement Google Sheets API integration
    showStatus('Google Sheets integration coming soon! For now, use the CSV download.', 'info');
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