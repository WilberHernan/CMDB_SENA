
const XLSX = require('xlsx');

function readExcelFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length > 0) {
      console.log('Columnas encontradas en el archivo XLSX:');
      console.log(data[0]);
    } else {
      console.log('El archivo XLSX está vacío o no tiene cabeceras.');
    }
  } catch (error) {
    console.error('Error al leer el archivo XLSX:', error);
  }
}

const filePath = 'CMDB CAUCA 2026 REGIONAL.xlsx';
readExcelFile(filePath);
