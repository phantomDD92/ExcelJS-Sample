const ExcelJS = require('exceljs');

async function createExcelFile() {
    // Create a new workbook
    const workbook = new ExcelJS.Workbook();

    // Add a worksheet
    const sheet = workbook.addWorksheet('Employees');

    // Add data
    sheet.addRow(['ID', 'Name', 'Department', 'Salary']);
    sheet.addRow([1, 'John Doe', 'Finance', 50000]);
    sheet.addRow([2, 'Jane Smith', 'HR', 60000]);
    sheet.addRow([3, 'John Doe', 'Finance', 50000]);
    sheet.addRow([4, 'Jane Smith', 'HR', 60000]);
    sheet.addRow([5, 'John Doe', 'Finance', 50000]);
    sheet.addRow([6, 'Jane Smith', 'HR', 60000]);
    sheet.addRow([7, 'John Doe', 'Finance', 50000]);
    sheet.addRow([8, 'Jane Smith', 'HR', 60000]);
    sheet.addRow([9, 'John Doe', 'Finance', 50000]);
    sheet.addRow([10, 'Jane Smith', 'HR', 60000]);
    sheet.addRow([11, 'John Doe', 'Finance', 50000]);
    sheet.addRow([12, 'Jane Smith', 'HR', 60000]);

    // Merge cells
    // sheet.mergeCells('A1:C1');

    // Apply styles
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' }
    };

    // Add a formula
    sheet.getCell('D3').value = { formula: 'SUM(D2:D3)', result: 110000 };

    // Set column widths
    sheet.getColumn(1).width = 20;
    sheet.getColumn(2).width = 15;
    sheet.getColumn(3).width = 15;

    // Save the workbook
    await workbook.xlsx.writeFile('complex_sample.xlsx');
    console.log('Excel file created successfully.');
}

createExcelFile().catch(err => {
    console.error('Error creating Excel file:', err);
});
