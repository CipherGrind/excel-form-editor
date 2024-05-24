let workbook;
let worksheet;
let sheetName = "Sheet1"; // Change to the name of your sheet

document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        worksheet = workbook.Sheets[sheetName];
        displayForm();
    };
    reader.readAsArrayBuffer(file);
}

function displayForm() {
    const formContainer = document.getElementById('formContainer');
    formContainer.innerHTML = '';

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    // Set grid template to accommodate row numbers and column headers
    formContainer.style.gridTemplateColumns = `50px repeat(${headers.length}, 150px)`; // Adjusted for better fit

    // Add column letters
    formContainer.appendChild(createCell('', 'header-cell')); // Top-left empty cell
    headers.forEach((_, index) => {
        formContainer.appendChild(createCell(columnToLetter(index), 'header-cell'));
    });

    // Add Excel column headers
    formContainer.appendChild(createCell('', 'header-cell')); // Top-left empty cell again for the row headers
    headers.forEach(header => {
        formContainer.appendChild(createCell(header, 'header-cell'));
    });

    // Add rows with row numbers and input cells
    rows.forEach((row, rowIndex) => {
        formContainer.appendChild(createCell(rowIndex + 1, 'header-cell row-header')); // Row number cell
        row.forEach((cellValue, colIndex) => {
            const input = document.createElement('input');
            input.type = 'text';
            input.value = cellValue !== undefined ? cellValue : '';
            input.dataset.rowIndex = rowIndex + 1; // Offset by 2 to account for header row
            input.dataset.colIndex = colIndex;
            formContainer.appendChild(createCell(input));
        });
    });
}

function createCell(content, className = 'cell') {
    const cell = document.createElement('div');
    cell.className = className;
    if (typeof content === 'string' || typeof content === 'number') {
        cell.textContent = content;
    } else {
        cell.appendChild(content);
    }
    return cell;
}

function columnToLetter(column) {
    let temp;
    let letter = '';
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

function updateExcel() {
    const formGroups = document.querySelectorAll('#formContainer input');
    formGroups.forEach(input => {
        const rowIndex = parseInt(input.dataset.rowIndex);
        const colIndex = parseInt(input.dataset.colIndex);
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        if (!worksheet[cellAddress]) {
            worksheet[cellAddress] = { t: 's', v: input.value };
        } else {
            worksheet[cellAddress].v = input.value;
        }
    });

    const updatedWorkbook = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([updatedWorkbook], { type: "application/octet-stream" }), 'updated_excel.xlsx');
}

function clearForm() {
    document.getElementById('fileInput').value = '';
    document.getElementById('formContainer').innerHTML = '';
    workbook = null;
    worksheet = null;
}