const ExcelJS = require('exceljs');

// Create a new workbook
const workbook = new ExcelJS.Workbook();

// Add worksheets
const sheet1 = workbook.addWorksheet('Sheet1');
const sheet2 = workbook.addWorksheet('Sheet2');

// Populate data
sheet1.addRow([1, 'John Doe', 25]);
sheet1.addRow([2, 'Jane Smith', 30]);

sheet2.addRow(['Product', 'Price']);
sheet2.addRow(['Apple', 1.5]);
sheet2.addRow(['Banana', 2]);

// Save the workbook
workbook.xlsx.writeFile('sample.xlsx')
    .then(() => {
        console.log('Excel file created successfully.');
    })
    .catch((err) => {
        console.error('Error creating Excel file:', err);
    });
