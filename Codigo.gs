const SPREADSHEET_ID = '1yGpolZYSmeIgjFeDOuN2C-cd2h-ipouZXx-PCZrVrmw';
const HOJAS_EQUIPOS = ['EquiposSena', 'EquiposTelefonica'];

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.queryParams = e.parameter || {};
  template.validaciones = obtenerValidacionesManuales();
  return template.evaluate()
    .setTitle('CMDB SENA CCYS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Configuración manual de menús desplegables.
 * Mapeados a las 50 columnas del spreadsheet.
 */
function obtenerValidacionesManuales() {
  const resultado = {};

  // Índice 1 = columna "TIPO"
  resultado[1] = ["DESKTOP", "PORTATIL", "TODO EN UNO"];

  // Índice 2 = columna "PROPIETARIO"
  resultado[2] = ["SENA", "TELEFONICA"];

  // Índice 3 = columna "MARCA"
  resultado[3] = ["LENOVO", "DELL", "HP", "ASUS", "JANUS", "ACER", "APPLE"];

  // Índice 7 = columna "ID SEDE"
  resultado[7] = ["65", "68", "69", "300", "319", "320", "321", "374"];

  // Índice 8 = columna "NOMBRE DE LA SEDE"
  resultado[8] = ["REGIONAL CAUCA", "CCYS", "GUAPI", "TECNOPARQUE", "SNFT", "ARCHIVO CENTRAL", "SAN JOSE", "LA PAMBA"];

  // Índice 9 = columna "CIUDAD"
  resultado[9] = ["POPAYAN", "GUAPI"];

  // Índice 10 = columna "UBICACIÓN"
  resultado[10] = ["OFICINA", "AMBIENTE"];

  // Índice 11 = columna "PISO"
  resultado[11] = ["1", "2", "3"];

  // Índice 13 = columna "TIPO DE USUARIO"
  resultado[13] = ["ADMINISTRATIVO", "CONTRATISTA", "INSTRUCTOR", "APRENDIZ"];

  // Índice 14 = columna "TIPO DE RED"
  resultado[14] = ["FUNCIONARIO", "FORMACION"];

  // Índice 15 = columna "PROCESADOR" (texto libre - nombre completo del CPU)
  // resultado[15] = ["RYZEN", "INTEL"];

  // Índice 16 = columna "TIPO DISCO 1"
  resultado[16] = ["HDD", "SSD", "M2"];

  // Índice 17 = columna "TAMAÑO DISCO 1"
  resultado[17] = ["120 GB", "256 GB", "512 GB", "1 TB"];

  // Índice 18 = columna "TIPO DISCO 2"
  resultado[18] = ["HDD", "SSD", "M2", "N/A"];

  // Índice 19 = columna "TAMAÑO DISCO 2"
  resultado[19] = ["120 GB", "256 GB", "512 GB", "1 TB", "N/A"];

  // Índice 20 = columna "TIPO MEMORIA"
  resultado[20] = ["DDR3", "DDR4", "DDR5"];

  // Índice 21 = columna "TAMAÑO MEMORIA"
  resultado[21] = ["4 GB", "8 GB", "16 GB", "32 GB", "64 GB"];

  // Índice 22 = columna "TARJETA DE VIDEO" (texto libre - nombre completo)
  // resultado[22] = ["AMD", "NVIDIA", "INTEL"];

  // Índice 32 = columna "SISTEMA OPERATIVO"
  resultado[32] = ["WINDOWS 10", "WINDOWS 11", "MAC OS MONTEREY", "MAC OS VENTURA"];

  // Índice 33 = columna "VERSION DEL S.O."
  resultado[33] = ["20H2", "21H1", "21H2", "22H2", "23H2", "25H2"];

  // Índice 34 = columna "ANTIVIRUS"
  resultado[34] = ["SI", "NO", "N/A"];

  // Índice 35 = columna "OFFICE"
  resultado[35] = ["SI", "NO", "N/A"];

  // Índice 36 = columna "ADOBE"
  resultado[36] = ["SI", "NO", "N/A"];

  // Índice 37 = columna "LAPS"
  resultado[37] = ["SI", "NO", "N/A"];

  // Índice 38 = columna "7ZIP"
  resultado[38] = ["SI", "NO", "N/A"];

  // Índice 39 = columna "VPN"
  resultado[39] = ["SI", "NO", "N/A"];

  // Índice 40 = columna "JAMF"
  resultado[40] = ["SI", "NO", "N/A"];

  // Índice 42 = columna "ESTADO DEL EQUIPO"
  resultado[42] = ["OPERATIVO", "PRESENTA FALLA", "DAÑADO"];

  // Índice 43 = columna "TIENE DOMINIO"
  resultado[43] = ["SI", "NO"];

  // Índice 44 = columna "EN QUE DOMINIO SE ENCUENTRA"
  resultado[44] = ["SENA.RED", "FORMACION.RED", "N/A"];

  // Índice 45 = columna "CONTRASEÑA BIOS"
  resultado[45] = ["SI", "NO", "N/A"];

  // Índice 48 = columna "ASS"
  resultado[48] = [
    "ANDRES SEBASTIAN BRAVO PALACIOS",
    "JULIAN ANDRES NOGUERA BURGOS",
    "LEONARDO ANDRES GUITIERREZ NARVAEZ",
    "HARRY LEHANDRO PEDRAZA ARROYO",
    "LUIS FELIPE FLOREZ DORADO",
    "YESID ANTONIO BRAVO RAMIREZ",
    "JHON ALEXANDER CORTES PAZ",
    "JESUS ALEXIS VEGA SANCHEZ",
    "JESUS HERNAN AMAYA ROJAS"
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
          validaciones: validacionesGlobales 
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
 * Guarda los cambios en la fila correspondiente
 */
function actualizarEquipo(datos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(datos.hoja);
    if (!sheet) throw new Error("No se encontró la hoja seleccionada.");

    const fila = parseInt(datos.fila);
    const valoresLimpios = sanitizarValores(datos.valores);
    const cantidadDatos = valoresLimpios.length;

    // Asegurar que la hoja tenga suficientes columnas
    if (sheet.getMaxColumns() < cantidadDatos) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), cantidadDatos - sheet.getMaxColumns());
    }

    sheet.getRange(fila, 1, 1, cantidadDatos).setValues([valoresLimpios]);

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
    return { exito: true, mensaje: 'Equipo registrado exitosamente en ' + hojaNombre };

  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}
