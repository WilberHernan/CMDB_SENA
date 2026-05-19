const xlsx = require('xlsx');

// Leer el archivo Excel
const workbook = xlsx.readFile('CMDB CAUCA 2026.xlsx');

// Obtener los nombres de las hojas
const sheetNames = workbook.SheetNames;
console.log('Hojas disponibles:', sheetNames);

if (sheetNames.length >= 3) {
    const sheet3Name = sheetNames[2]; // Índice 2 es la hoja 3
    console.log(`\nLeyendo la Hoja 3: "${sheet3Name}"`);
    
    const worksheet = workbook.Sheets[sheet3Name];
    
    // Convertir a JSON (solo la primera fila para obtener los campos/headers)
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length > 0) {
        console.log('\nCampos (Columnas) de la Hoja 3:');
        data[0].forEach((col, index) => {
            console.log(`${index + 1}. ${col}`);
        });
    } else {
        console.log('La hoja 3 está vacía o no tiene encabezados.');
    }
} else {
    console.log('El archivo no tiene 3 hojas.');
}
