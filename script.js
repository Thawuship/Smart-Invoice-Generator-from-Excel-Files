let productList = [];
let invoiceID = 1;

document.getElementById('fileInput').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            productList = json.slice(1); // Remove header
            displayData(json);
        };
        reader.readAsArrayBuffer(file);
    }
});

function displayData(data) {
    const tableBody = document.querySelector('#excelTable tbody');
    tableBody.innerHTML = ''; // Clear existing data

    // Display the first 5 rows (excluding header)
    const maxRows = 5;
    for (let i = 1; i <= maxRows && i < data.length; i++) {
        const row = data[i];
        const tr = document.createElement('tr');
        row.forEach(cellValue => {
            const td = document.createElement('td');
            td.textContent = cellValue;
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    }
}

document.getElementById('createInvoiceBtn').addEventListener('click', function() {
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('invoiceSection').style.display = 'block';
});

document.getElementById('addRecordBtn').addEventListener('click', function() {
    addInvoiceRecord();
});

document.getElementById('generateInvoiceBtn').addEventListener('click', function() {
    generateInvoice();
});

function addInvoiceRecord() {
    const tableBody = document.querySelector('#invoiceTable tbody');
    const tr = document.createElement('tr');

    const idTd = document.createElement('td');
    idTd.className = 'id-column';
    idTd.textContent = invoiceID++;
    tr.appendChild(idTd);

    const productTd = document.createElement('td');
    const select = document.createElement('select');
    productList.forEach(product => {
        const option = document.createElement('option');
        option.value = product[1];
        option.textContent = product[1];
        select.appendChild(option);
    });
    productTd.appendChild(select);
    tr.appendChild(productTd);

    const quantityTd = document.createElement('td');
    quantityTd.className = 'quantity-column';
    const quantityInput = document.createElement('input');
    quantityInput.type = 'number';
    quantityInput.min = '1';
    quantityInput.value = '1';
    quantityInput.addEventListener('input', updateTotal);
    quantityTd.appendChild(quantityInput);
    tr.appendChild(quantityTd);

    const priceTd = document.createElement('td');
    priceTd.className = 'price-column';
    const priceInput = document.createElement('input');
    priceInput.type = 'number';
    priceInput.min = '0.01';
    priceInput.step = '0.01';
    priceInput.value = '0.00';
    priceInput.addEventListener('input', updateTotal);
    priceTd.appendChild(priceInput);
    tr.appendChild(priceTd);

    const totalTd = document.createElement('td');
    totalTd.className = 'total-column';
    totalTd.textContent = '0.00';
    tr.appendChild(totalTd);

    const actionTd = document.createElement('td');
    const deleteBtn = document.createElement('i');
    deleteBtn.className = 'fas fa-trash delete-icon';
    deleteBtn.addEventListener('click', () => {
        tr.remove();
        updateTotal();
    });
    actionTd.appendChild(deleteBtn);
    tr.appendChild(actionTd);

    tableBody.appendChild(tr);
    updateTotal();
}

function updateTotal() {
    const tableBody = document.querySelector('#invoiceTable tbody');
    let totalAmount = 0;

    tableBody.querySelectorAll('tr').forEach(row => {
        const quantity = parseFloat(row.children[2].querySelector('input').value);
        const price = parseFloat(row.children[3].querySelector('input').value);
        const total = quantity * price;
        row.children[4].textContent = total.toFixed(2);
        totalAmount += total;
    });

    document.getElementById('totalAmount').textContent = `Total Amount: $${totalAmount.toFixed(2)}`;
}

function generateInvoice() {
    const invoiceWindow = window.open('', '', 'width=800,height=600');
    const invoiceTable = document.createElement('table');
    const originalTable = document.getElementById('invoiceTable');

    // Copy headers
    const thead = originalTable.querySelector('thead').cloneNode(true);
    const actionIndex = Array.from(thead.rows[0].cells).findIndex(cell => cell.textContent === 'Action');
    if (actionIndex >= 0) thead.rows[0].deleteCell(actionIndex);
    invoiceTable.appendChild(thead);

    // Copy rows
    const tbody = document.createElement('tbody');
    originalTable.querySelectorAll('tbody tr').forEach(originalRow => {
        const row = document.createElement('tr');
        Array.from(originalRow.children).forEach((cell, index) => {
            if (index !== actionIndex) { // Exclude the action column
                const td = document.createElement('td');
                if (index === 2 || index === 3) { // Quantity or Price columns
                    td.textContent = cell.querySelector('input').value;
                } else if (index === 1) { // Product column
                    td.textContent = cell.querySelector('select').selectedOptions[0].textContent;
                } else {
                    td.textContent = cell.textContent;
                }
                if (index === 0) {
                    td.className = 'id-column';
                } else if (index === 2) {
                    td.className = 'quantity-column';
                } else if (index === 3) {
                    td.className = 'price-column';
                } else if (index === 4) {
                    td.className = 'total-column';
                }
                row.appendChild(td);
            }
        });
        tbody.appendChild(row);
    });
    invoiceTable.appendChild(tbody);

    const totalAmount = document.getElementById('totalAmount').textContent;

    invoiceWindow.document.write('<html><head><title>Invoice</title>');
    invoiceWindow.document.write('<link rel="stylesheet" href="styles.css" type="text/css" />');
    invoiceWindow.document.write('</head><body>');
    invoiceWindow.document.write('<h1>Invoice</h1>');
    invoiceWindow.document.write(invoiceTable.outerHTML);
    invoiceWindow.document.write('<p>' + totalAmount + '</p>');
    invoiceWindow.document.write('</body></html>');
    invoiceWindow.document.close();
    invoiceWindow.print();
}
