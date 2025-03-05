// Guardar y cargar el folio
function loadFolio() {
    const storedFolio = localStorage.getItem('folio');
    return storedFolio ? parseInt(storedFolio, 10) : 1;
}

function saveFolio(folio) {
    localStorage.setItem('folio', folio);
}

// Procesar archivo de entrada
document.querySelector('#file-input').addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const fileData = e.target.result;
        let data;
        const ext = file.name.split('.').pop().toLowerCase();

        if (ext === 'xlsx' || ext === 'xls') {
            data = XLSX.read(fileData, { type: 'binary' });
            const sheet = data.Sheets[data.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            populateTable(jsonData);
        }
    };
    reader.readAsBinaryString(file);
}

function populateTable(data) {
    const tableBody = document.querySelector('#data-table tbody');
    tableBody.innerHTML = '';  
    let folio = loadFolio();
    const rows = data.slice(1);

    rows.forEach((row) => {
        const barcodeData = row[0];
        const barcodeText = `${row[0]} - ${row[1]}`;
        const barcodeCanvas = document.createElement('canvas');

        JsBarcode(barcodeCanvas, barcodeData, {
            format: "CODE128",
            displayValue: true,
            text: barcodeText,
            lineColor: "#000",
            width: 2,
            height: 50
        });

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatDate(row[2])}</td>
            <td>${row[3]}</td>
            <td>${row[4]}</td>
            <td>${row[5]}</td>
            <td>${row[6]}</td>
            <td>${row[7]}</td>
        `;

        const barcodeTd = document.createElement('td');
        barcodeTd.appendChild(barcodeCanvas);
        tr.appendChild(barcodeTd);

        tableBody.appendChild(tr);
    });

    saveFolio(folio);
}

// Función para formatear la fecha en formato dd/mm/yyyy
function formatDate(date) {
    // Verificar si la fecha es un número, lo que indica que proviene de Excel
    if (typeof date === 'number') {
        const excelDate = new Date((date - 25569) * 86400 * 1000); // Convertir número de Excel a fecha
        return formatToDDMMYYYY(excelDate);
    }

    const parsedDate = new Date(date);
    if (isNaN(parsedDate.getTime())) return date;  // Si no es una fecha válida, retornamos tal cual

    // Extraemos el día, mes y año
    return formatToDDMMYYYY(parsedDate);
}

function formatToDDMMYYYY(parsedDate) {
    const day = parsedDate.getDate().toString().padStart(2, '0');
    const month = (parsedDate.getMonth() + 1).toString().padStart(2, '0');
    const year = parsedDate.getFullYear();
    return `${day}/${month}/${year}`;
}

// Generar todos los códigos cuando se hace clic en el botón
document.querySelector('#generate-all-codes-btn').addEventListener('click', generateAllBarcodes);

function generateAllBarcodes() {
    const rows = document.querySelectorAll('#data-table tbody tr');
    rows.forEach((row, index) => {
        const barcodeData = row.children[0].innerText;

        let barcodeCanvas = row.querySelector('canvas');
        if (!barcodeCanvas) {
            barcodeCanvas = document.createElement('canvas');
            row.children[row.children.length - 1].appendChild(barcodeCanvas);
        }

        JsBarcode(barcodeCanvas, barcodeData, {
            format: "CGS1-128",
            displayValue: true,
            lineColor: "#000",
            width: 2,
            height: 50
        });
    });
}

// Exportar tabla a archivo XLS
document.querySelector('#export-btn').addEventListener('click', exportToXLS);

function exportToXLS() {
    const table = document.querySelector('#data-table');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['Código', 'Descripción', 'Fecha', 'Factura', 'OC', 'Registro', 'Proveedor', 'Cliente', 'Código de Barras']]);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    html2canvas(table).then((canvas) => {
        const img = canvas.toDataURL('image/png');

        ws['!images'] = [{ image: img, from: { col: 0, row: 1 }, to: { col: 8, row: rows.length } }];
        XLSX.writeFile(wb, 'table_data.xlsx');
    });
}

// Imprimir códigos de barras
document.querySelector('#print-barcode-btn').addEventListener('click', printBarcodes);

function printBarcodes() {
    const rows = document.querySelectorAll('#data-table tbody tr');
    const barcodeData = [];

    rows.forEach((row) => {
        const barcodeCanvas = row.querySelector('canvas');
        const code = row.children[0].innerText;
        const date = row.children[2].innerText;
        const url = row.children[6].innerText;
        if (barcodeCanvas) {
            const barcodeImage = barcodeCanvas.toDataURL('image/png');
            barcodeData.push({ image: barcodeImage, code: code, date: date, url: url });
        }
    });

    if (barcodeData.length > 0) {
        const printWindow = window.open('', '', 'width=800,height=600');
        printWindow.document.write('<html><head><title>Impresión de Códigos de Barras</title>');
        // Modificador de pagina y contenedor de códigos
        printWindow.document.write(`
            <style>
                @media print {
                    @page {
                        margin: 0; /* Aseguramos que no haya margen de la impresora */
                    }
                    body { font-family: Arial, sans-serif; margin: 0; }
                    .page { 
                        display: grid; 
                        grid-template-columns: repeat(3, 7.8cm); 
                        grid-template-rows: repeat(10, 2.87cm); 
                        column-gap: 6mm; 
                        row-gap: 0mm;
                        width: 21.59cm;
                        height: 27.94cm;
                        page-break-after: always; 
                        padding: 15mm 3mm; 
                        box-sizing: border-box; 
                    }
                    .barcode-container { 
                        display: flex; 
                        flex-direction: column; 
                        justify-content: center; 
                        align-items: center; 
                        border: 1px solid #ccc; 
                        padding: 0; 
                        margin: 2mm; 
                        height: 2.87cm; 
                        width: 7.8cm; 
                    }
                    img { 
                        width: 93%; 
                        height: auto; 
                    }
                    .barcode-code { 
                        font-size: 5px; 
                        text-align: center; 
                        margin-top: 2px; 
                    }
                }
            </style>
        `);
        printWindow.document.write('</head><body>');

        const barcodesPerPage = 30;
        for (let i = 0; i < barcodeData.length; i += barcodesPerPage) {
            const pageBarcodeData = barcodeData.slice(i, i + barcodesPerPage);
            printWindow.document.write('<div class="page">');
            pageBarcodeData.forEach((barcode) => {
                printWindow.document.write(`
                    <div class="barcode-container">
                        <img src="${barcode.image}" alt="Código de Barras">
                        <div class="barcode-code" style="font-size: small">${barcode.date}</div>
                    </div>
                `);
            });
            printWindow.document.write('</div>');
        }

        printWindow.document.write('</body></html>');
        printWindow.document.close();

        printWindow.onload = function () {
            printWindow.print();
            printWindow.close();
        };
    } else {
        alert('No se encontraron códigos de barras para imprimir.');
    }
}