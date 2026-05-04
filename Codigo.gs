const SPREADSHEET_ID = '1yGpolZYSmeIgjFeDOuN2C-cd2h-ipouZXx-PCZrVrmw';
const HOJAS_EQUIPOS = ['EquiposSena', 'EquiposTelefonica'];

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.queryParams = e.parameter || {};
  template.validaciones = obtenerValidacionesManuales();
  template.mapeoSedeId = obtenerMapeoSedeId();
  return template.evaluate()
    .setTitle('CMDB SENA CCYS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Obtiene el mapeo bidireccional Sede ↔ ID desde la Hoja3.
 * La Hoja3 tiene las columnas 6 (ID numérico) y 7 (nombre sede) invertidas respecto a sus headers.
 */
function obtenerMapeoSedeId() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Hoja3');
  if (!sheet) return { sedeAId: {}, idASede: {} };

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { sedeAId: {}, idASede: {} };

  const sedeAId = {};
  const idASede = {};

  for (let i = 1; i < data.length; i++) {
    // En Hoja3: col 6 tiene los IDs numéricos, col 7 tiene los nombres de sede
    const idRaw = data[i][6];
    const sedeRaw = data[i][7];

    if (idRaw && sedeRaw) {
      const id = idRaw.toString().trim();
      const sede = sedeRaw.toString().trim().toUpperCase();
      if (id && sede) {
        sedeAId[sede] = id;
        idASede[id] = sede;
      }
    }
  }

  return { sedeAId: sedeAId, idASede: idASede };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Lee las validaciones dinámicamente desde la hoja 'Hoja3'.
 * Cada columna de Hoja3 contiene los valores posibles para cada dropdown.
 * Los valores vacíos se ignoran automáticamente.
 */
function obtenerValidacionesManuales() {
  const resultado = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Hoja3');

  if (!sheet) {
    // Fallback si no existe Hoja3
    Logger.log('⚠️ No se encontró Hoja3, usando validaciones por defecto');
    return obtenerValidacionesPorDefecto();
  }

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return resultado;

  // Mapeo: columna en Hoja3 → índice en el formulario principal
  // Nota: columna 12 en Hoja3 tiene header "en " pero son los TIPOS DE USUARIO
  const mapeo = {
    1:  1,   // TIPO
    2:  2,   // PROPIETARIO
    3:  3,   // MARCA
    6:  7,   // ID → ID SEDE
    7:  8,   // NOMBRE DE LA SEDE
    8:  9,   // CIUDAD
    9:  10,  // UBICACIÓN
    10: 11,  // PISO
    12: 13,  // en  → TIPO DE USUARIO
    13: 14,  // TIPO DE RED
    15: 16,  // TIPO DISCO 1
    16: 17,  // TAMAÑO DISCO 1
    17: 18,  // TIPO DISCO 2
    18: 19,  // TAMAÑO DISCO 2
    19: 20,  // TIPO MEMORIA
    20: 21,  // TAMAÑO MEMORIA
    28: 32,  // SISTEMA OPERATIVO
    29: 33,  // VERSION DEL S.O.
    30: 34,  // ANTIVIRUS
    31: 42,  // ESTADO DEL EQUIPO
    32: 43,  // TIENE DOMINIO
    33: 44,  // EN QUE DOMINIO SE ENCUENTRA
    37: 48   // ASS
  };

  // Extraer valores únicos por columna (ignorando vacíos)
  Object.keys(mapeo).forEach(function(colHoja3Str) {
    const colHoja3 = parseInt(colHoja3Str);
    const indiceForm = mapeo[colHoja3];
    const valoresUnicos = [];
    const visto = {};

    for (let i = 1; i < data.length; i++) {
      const raw = data[i][colHoja3];
      if (raw === undefined || raw === null || raw.toString().trim() === '') continue;
      const val = raw.toString().trim().toUpperCase();
      if (!visto[val]) {
        visto[val] = true;
        valoresUnicos.push(val);
      }
    }

    if (valoresUnicos.length > 0) {
      resultado[indiceForm] = valoresUnicos;
    }
  });

  // Validaciones manuales para campos que NO están en Hoja3
  resultado[35] = ["SI", "NO", "N/A"]; // OFFICE
  resultado[36] = ["SI", "NO", "N/A"]; // ADOBE
  resultado[37] = ["SI", "NO", "N/A"]; // LAPS
  resultado[38] = ["SI", "NO", "N/A"]; // 7ZIP
  resultado[39] = ["SI", "NO", "N/A"]; // VPN
  resultado[40] = ["SI", "NO", "N/A"]; // JAMF
  resultado[45] = ["SI", "NO", "N/A"]; // CONTRASEÑA BIOS

  return resultado;
}

/**
 * Validaciones de respaldo en caso de que no exista Hoja3.
 */
function obtenerValidacionesPorDefecto() {
  const resultado = {};
  resultado[1] = ["DESKTOP", "PORTATIL", "TODO EN UNO", "IMAC"];
  resultado[2] = ["SENA", "TELEFONICA"];
  resultado[3] = ["LENOVO", "DELL", "HP", "ASUS", "JANUS", "ACER", "APPLE"];
  resultado[7] = ["65", "68", "69", "300", "319", "320", "321", "374", "389"];
  resultado[8] = ["REGIONAL", "CCYS", "GUAPI", "TECNOPARQUE", "SNFT", "ARCHIVO CENTRAL", "SAN JOSE", "LA PAMBA", "CIUDAD JARDIN"];
  resultado[9] = ["POPAYAN", "GUAPI"];
  resultado[10] = ["OFICINA", "AMBIENTE"];
  resultado[11] = ["1", "2", "3"];
  resultado[13] = ["ADMINISTRATIVO", "CONTRATISTA", "INSTRUCTOR", "APRENDIZ"];
  resultado[14] = ["FUNCIONARIO", "FORMACION"];
  resultado[16] = ["HDD", "SSD", "M2"];
  resultado[17] = ["120 GB", "256 GB", "512 GB", "1 TB"];
  resultado[18] = ["HDD", "SSD", "M2", "N/A"];
  resultado[19] = ["120 GB", "256 GB", "512 GB", "1 TB", "N/A"];
  resultado[20] = ["DDR3", "DDR4", "DDR5"];
  resultado[21] = ["4 GB", "8 GB", "16 GB", "32 GB", "64 GB"];
  resultado[32] = ["WINDOWS 10", "WINDOWS 11", "MAC OS MONTEREY", "MAC OS VENTURA"];
  resultado[33] = ["20H2", "21H1", "21H2", "22H2", "23H2"];
  resultado[34] = ["SI", "NO", "N/A"];
  resultado[35] = ["SI", "NO", "N/A"];
  resultado[36] = ["SI", "NO", "N/A"];
  resultado[37] = ["SI", "NO", "N/A"];
  resultado[38] = ["SI", "NO", "N/A"];
  resultado[39] = ["SI", "NO", "N/A"];
  resultado[40] = ["SI", "NO", "N/A"];
  resultado[42] = ["OPERATIVO", "PRESENTA FALLA", "DAÑADO"];
  resultado[43] = ["SI", "NO"];
  resultado[44] = ["SENA.RED", "FORMACION.RED", "N/A"];
  resultado[45] = ["SI", "NO", "N/A"];
  resultado[48] = [
    "ANDRES SEBASTIAN BRAVO PALACIOS",
    "JULIAN ANDRES NOGUERA BURGOS",
    "LEONARDO ANDRES GUITIERREZ NARVAEZ",
    "HARRY LEHANDRO PEDRAZA ARROYO",
    "LUIS FELIPE FLOREZ DORADO",
    "YESID ANTONIO BRAVO RAMIREZ",
    "JESUS ALEXIS VEGA SANCHEZ",
    "JESUS HERNAN AMAYA ROJAS",
    "JHON ALEXANDER CORTES PAZ"
  ];
  return resultado;
}

/**
 * Busca el equipo por placa en todas las hojas
 */
function buscarEquipo(placaEscaneada) {
  const placaBuscada = placaEscaneada.toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase();
  if (!placaBuscada) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const validacionesGlobales = obtenerValidacionesManuales();
  const mapeoSedeId = obtenerMapeoSedeId();

  for (let h = 0; h < HOJAS_EQUIPOS.length; h++) {
    const sheet = ss.getSheetByName(HOJAS_EQUIPOS[h]);
    if (!sheet) continue; 
    
    const data = sheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      let placaEnHoja = data[i][6] ? data[i][6].toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase() : "";
      
      if (placaEnHoja === placaBuscada) {
        return { 
          hoja: HOJAS_EQUIPOS[h], 
          fila: i + 1, 
          valores: data[i], 
          validaciones: validacionesGlobales,
          mapeoSedeId: mapeoSedeId
        };
      }
    }
  }
  return null; 
}

/**
 * Sanitiza valores: reemplaza apóstrofos por guiones (fallo común de lectores QR)
 */
function sanitizarValores(valores) {
  return valores.map(function(v) {
    if (typeof v === 'string') {
      return v.replace(/'/g, '-');
    }
    return v;
  });
}

/**
 * Guarda los cambios en la fila correspondiente.
 * Si cambió el propietario, mueve el equipo a la hoja destino correcta.
 */
function actualizarEquipo(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetOrigen = ss.getSheetByName(datos.hoja);
    if (!sheetOrigen) throw new Error("No se encontró la hoja origen: " + datos.hoja);

    const fila = parseInt(datos.fila);
    const valoresLimpios = sanitizarValores(datos.valores);
    const cantidadDatos = valoresLimpios.length;

    // Detectar si cambió el propietario (columna C = índice 2)
    const propietarioNuevo = (valoresLimpios[2] || '').toString().toUpperCase().trim();
    const hojaDestino = (propietarioNuevo === 'TELEFONICA') ? 'EquiposTelefonica' : 'EquiposSena';

    // Leer propietario actual de la hoja origen
    const valoresActuales = sheetOrigen.getRange(fila, 1, 1, cantidadDatos).getDisplayValues()[0];
    const propietarioActual = (valoresActuales[2] || '').toString().toUpperCase().trim();

    // Si cambió de propietario, mover el equipo a la hoja correspondiente
    if (hojaDestino !== datos.hoja && propietarioActual !== propietarioNuevo) {
      const sheetDestino = ss.getSheetByName(hojaDestino);
      if (!sheetDestino) throw new Error("No se encontró la hoja destino: " + hojaDestino);

      // Validar que la placa no exista ya en la hoja destino
      const placa = valoresLimpios[6]
        ? valoresLimpios[6].toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase()
        : '';

      if (placa) {
        const dataDest = sheetDestino.getDataRange().getDisplayValues();
        for (let i = 1; i < dataDest.length; i++) {
          const placaExistente = dataDest[i][6]
            ? dataDest[i][6].toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase()
            : '';
          if (placaExistente === placa) {
            return {
              exito: false,
              mensaje: 'La placa ' + placa + ' ya existe en ' + hojaDestino + ' (fila ' + (i + 1) + '). No se puede mover el equipo.'
            };
          }
        }
      }

      // Asegurar columnas en hoja destino
      if (sheetDestino.getMaxColumns() < cantidadDatos) {
        sheetDestino.insertColumnsAfter(sheetDestino.getMaxColumns(), cantidadDatos - sheetDestino.getMaxColumns());
      }

      // Insertar en hoja destino
      sheetDestino.appendRow(valoresLimpios);
      const ultimaFila = sheetDestino.getLastRow();
      const rangoNuevo = sheetDestino.getRange(ultimaFila, 1, 1, cantidadDatos);
      rangoNuevo.setDataValidation(null);
      rangoNuevo.setBackground('#dbeafe'); // azul para movido/nuevo

      // Eliminar de hoja origen
      sheetOrigen.deleteRow(fila);

      return {
        exito: true,
        mensaje: 'Equipo movido exitosamente de ' + datos.hoja + ' a ' + hojaDestino
      };
    }

    // Si no cambió de propietario, actualizar normalmente en la misma hoja
    if (sheetOrigen.getMaxColumns() < cantidadDatos) {
      sheetOrigen.insertColumnsAfter(sheetOrigen.getMaxColumns(), cantidadDatos - sheetOrigen.getMaxColumns());
    }

    const rangoActualizar = sheetOrigen.getRange(fila, 1, 1, cantidadDatos);

    // Eliminar validaciones de datos heredadas del Excel antes de escribir
    rangoActualizar.setDataValidation(null);

    rangoActualizar.setValues([valoresLimpios]);

    // Pintar fila actualizada de verde pálido sutil
    rangoActualizar.setBackground('#dcfce7');

    return { exito: true, mensaje: "¡CMDB Actualizada con éxito!" };

  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

/**
 * Crea un nuevo equipo al final de la hoja indicada.
 * Realiza validación para evitar duplicados de placa.
 */
function crearEquipo(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojaNombre = datos.hoja || 'EquiposSena';
    const sheet = ss.getSheetByName(hojaNombre);
    if (!sheet) throw new Error("No se encontró la hoja: " + hojaNombre);

    const valoresLimpios = sanitizarValores(datos.valores);
    const cantidadDatos = valoresLimpios.length;
    if (sheet.getMaxColumns() < cantidadDatos) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), cantidadDatos - sheet.getMaxColumns());
    }

    // Validar duplicado de placa (columna G = índice 6)
    const placaNueva = valoresLimpios[6]
      ? valoresLimpios[6].toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase()
      : '';

    if (placaNueva) {
      const data = sheet.getDataRange().getDisplayValues();
      for (let i = 1; i < data.length; i++) {
        const placaExistente = data[i][6]
          ? data[i][6].toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9-]/g, '').toUpperCase()
          : '';
        if (placaExistente === placaNueva) {
          return {
            exito: false,
            mensaje: 'La placa ' + placaNueva + ' ya existe en ' + hojaNombre + ' (fila ' + (i + 1) + ').'
          };
        }
      }
    }

    sheet.appendRow(valoresLimpios);

    // Pintar fila nueva de azul pálido sutil
    const ultimaFila = sheet.getLastRow();
    const rangoNuevo = sheet.getRange(ultimaFila, 1, 1, cantidadDatos);
    rangoNuevo.setBackground('#dbeafe');

    return { exito: true, mensaje: 'Equipo registrado exitosamente en ' + hojaNombre };

  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}
