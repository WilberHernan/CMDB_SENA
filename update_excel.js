const XLSX = require('xlsx');

function addColumnToExcel(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convertir a JSON para manipular más fácil
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length === 0) {
      console.log('El archivo está vacío');
      return;
    }
    
    // Encontrar el índice de la columna UBICACIÓN
    const headers = data[0];
    const ubicacionIndex = headers.indexOf('UBICACIÓN');
    
    if (ubicacionIndex === -1) {
      console.log('No se encontró la columna UBICACIÓN');
      return;
    }
    
    console.log(`Columna UBICACIÓN encontrada en índice: ${ubicacionIndex}`);
    
    // Insertar la nueva columna después de UBICACIÓN
    const newColumnIndex = ubicacionIndex + 1;
    const newColumnName = 'NOMBRE DE LA OFICINA O AMBIENTE';
    
    // Insertar header
    headers.splice(newColumnIndex, 0, newColumnName);
    
    // Insertar celda vacía en cada fila de datos
    for (let i = 1; i < data.length; i++) {
      // Asegurar que la fila tenga suficientes elementos
      while (data[i].length < newColumnIndex) {
        data[i].push('');
      }
      data[i].splice(newColumnIndex, 0, '');
    }
    
    console.log(`Columna '${newColumnName}' insertada en índice: ${newColumnIndex}`);
    console.log('Nuevos headers:', headers);
    
    // Convertir de vuelta a worksheet
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Reemplazar la hoja en el workbook
    workbook.Sheets[sheetName] = newWorksheet;
    
    // Guardar el archivo
    XLSX.writeFile(workbook, filePath);
    
    console.log('Archivo Excel actualizado exitosamente');
  } catch (error) {
    console.error('Error al modificar el Excel:', error);
  }
}

const filePath = 'CMDB CAUCA 2026 REGIONAL.xlsx';
addColumnToExcel(filePath);
