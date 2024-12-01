document.getElementById('fileUpload').addEventListener('change', handleFileUpload);
document.getElementById('manualSearch').addEventListener('click', handleSearch);
document.getElementById('addColor').addEventListener('click', addNewColor);
document.getElementById('exportResults').addEventListener('click', exportResults);

let excelData = []; // Aquí almacenaremos los datos del archivo Excel

// Carga y procesamiento del archivo Excel
function handleFileUpload(event) {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convertimos la hoja a un arreglo

        alert('Archivo cargado exitosamente');
        console.log(excelData); // Verifica los datos en la consola
    };

    reader.readAsArrayBuffer(file);
}

// Búsqueda manual del código
function handleSearch() {
    const code = document.getElementById('codeInput').value;

    if (code) {
        const result = findCodeInExcel(code);

        if (result) {
            const { color, zone } = result;
            alert(El paquete con código ${code} está en la lista. Está en la Zona ${zone} (${color}).);
            addToResultsTable(code, color, zone);
        } else {
            alert(El paquete con código ${code} no está en la lista.);
        }
    } else {
        alert('Por favor, ingresa un código.');
    }
}

// Buscar código en el Excel cargado
function findCodeInExcel(code) {
    for (let i = 1; i < excelData.length; i++) { // Comenzamos en 1 para omitir los encabezados
        const row = excelData[i];
        const codeInRow = row[0]; // Columna A
        const colorInRow = row[10]; // Columna K

        if (codeInRow == code) {
            // Analizamos el color
            const rgb = colorInRow.trim();
            let zone = '';

            if (rgb === '255,0,0') zone = '1 (Rojo)';
            else if (rgb === '0,0,255') zone = '2 (Azul)';
            else zone = 'Zona Desconocida';

            return { color: rgb, zone };
        }
    }
    return null;
}

// Agregar colores personalizados
function addNewColor() {
    const newColor = prompt("Ingrese el color en formato RGB (ejemplo: 128,128,128):");
    if (newColor) {
        const colorSelect = document.getElementById('colorSelect');
        const option = document.createElement('option');
        option.value = newColor;
        option.textContent = Color Personalizado (${newColor});
        colorSelect.appendChild(option);
    }
}

// Agregar resultado a la tabla
function addToResultsTable(code, color, zone) {
    const tableBody = document.getElementById('resultsTable').querySelector('tbody');
    const newRow = document.createElement('tr');

    newRow.innerHTML = `
        <td>${code}</td>
        <td>${color}</td>
        <td>${zone}</td>
    `;

    tableBody.appendChild(newRow);
}

// Exportar resultados a CSV
function exportResults() {
    const table = document.getElementById('resultsTable');
    const rows = Array.from(table.querySelectorAll('tr')).map(row => 
        Array.from(row.querySelectorAll('td, th')).map(cell => cell.textContent)
    );

    const csvContent = rows.map(row => row.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const link = document.createElement('a');

    link.href = URL.createObjectURL(blob);
    link.download = 'resultados.csv';
    link.click();
}