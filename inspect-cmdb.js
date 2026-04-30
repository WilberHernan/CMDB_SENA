const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'CMDB CAUCA 2026.xlsx');

if (!fs.existsSync(filePath)) {
  console.error('No se encontró el archivo:', filePath);
  process.exit(1);
}

console.log('📁 Archivo:', path.basename(filePath));
console.log('📏 Tamaño:', (fs.statSync(filePath).size / 1024).toFixed(2), 'KB');
console.log('');

const workbook = XLSX.readFile(filePath);

console.log('📑 Hojas encontradas:', workbook.SheetNames.length);
console.log('');

workbook.SheetNames.forEach(sheetName => {
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  console.log(`📄 Hoja: ${sheetName}`);
  console.log(`   Filas: ${data.length}`);

  if (data.length > 0) {
    console.log(`   Columnas: ${data[0].length}`);
    console.log('');
    console.log('   🏷️  Encabezados:');
    data[0].forEach((col, i) => {
      console.log(`      ${String(i + 1).padStart(2)}. ${col || '(vacía)'}`);
    });

    if (data.length > 1) {
      console.log('');
      console.log('   🔍 Primera fila de datos:');
      const firstRow = data[1];
      data[0].forEach((col, i) => {
        const val = firstRow[i] !== undefined ? firstRow[i] : '(vacío)';
        console.log(`      ${String(i + 1).padStart(2)}. ${col || '?'}: ${val}`);
      });
    }
  }

  console.log('');
});

// Estadísticas generales
const totalEquipos = workbook.SheetNames.reduce((sum, name) => {
  const ws = workbook.Sheets[name];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return sum + Math.max(0, data.length - 1);
}, 0);

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log(`📊 Total de equipos registrados: ${totalEquipos}`);
