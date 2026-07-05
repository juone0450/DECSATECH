const xlsx = require('xlsx');
const fs = require('fs');
const workbook = xlsx.readFile('Lista para Armador pc.xlsx');
const sheetName = workbook.SheetNames[0];
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 1});
fs.writeFileSync('excel_data2.json', JSON.stringify(data, null, 2), 'utf8');
