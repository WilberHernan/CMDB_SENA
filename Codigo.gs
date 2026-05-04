const SPREADSHEET_ID = '1yGpolZYSmeIgjFeDOuN2C-cd2h-ipouZXx-PCZrVrmw';
const HOJAS_EQUIPOS = ['EquiposSena', 'EquiposTelefonica'];

/* ================================================================
   CONFIGURACION ROBUSTA: Mapeo de campos del formulario principal
   ================================================================ */
const FORM_FIELDS = {
  "HOSTNAME": 0, "TIPO": 1, "PROPIETARIO": 2, "MARCA": 3, "MODELO": 4,
  "SERIAL": 5, "PLACA": 6, "ID SEDE": 7, "NOMBRE DE LA SEDE": 8,
  "CIUDAD": 9, "UBICACIÓN": 10, "PISO": 11, "NOMBRE DEL USUARIO": 12,
  "TIPO DE USUARIO": 13, "TIPO DE RED": 14, "PROCESADOR": 15,
  "TIPO DISCO 1": 16, "TAMAÑO DISCO 1": 17, "TIPO DISCO 2": 18,
  "TAMAÑO DISCO 2": 19, "TIPO MEMORIA": 20, "TAMAÑO MEMORIA": 21,
  "TARJETA DE VIDEO": 22, "CAMBIO DE PARTE": 23, "CAMBIO DE PARTE 2": 24,
  "# DE CASO PARA REPUESTO": 25, "PLACA MONITOR": 26, "PLACA MOUSE": 27,
  "PLACA TECLADO": 28, "PLACA CARGADOR": 29, "MAC:RED CABLEADA": 30,
  "MAC RED INALAMBRICA": 31, "SISTEMA OPERATIVO": 32, "VERSION DEL S.O.": 33,
  "ANTIVIRUS": 34, "OFFICE": 35, "ADOBE": 36, "LAPS": 37, "7ZIP": 38,
  "VPN": 39, "JAMF": 40, "OTRO SOFTWARE": 41, "ESTADO DEL EQUIPO": 42,
  "TIENE DOMINIO": 43, "EN QUE DOMINIO SE ENCUENTRA": 44,
  "CONTRASEÑA BIOS": 45, "FECHA ULTIMO MANTENIMIENTO": 46,
  "FECHA IMPACTO MAQUINA": 47, "ASS": 48, "Observaciones": 49
};

const HEADER_ALIASES = {
  "EN": "TIPO DE USUARIO",
  "EN ": "TIPO DE USUARIO",
  "TIPO USUARIO": "TIPO DE USUARIO",
  "USUARIO": "TIPO DE USUARIO",
  "ID": "ID SEDE",
  "NUMERO SEDE": "ID SEDE",
  "NOMBRE SEDE": "NOMBRE DE LA SEDE",
  "SEDE": "NOMBRE DE LA SEDE",
  "UBICACION": "UBICACIÓN",
  "VERSION SO": "VERSION DEL S.O.",
  "VERSION S.O.": "VERSION DEL S.O.",
  "VERSION SISTEMA OPERATIVO": "VERSION DEL S.O.",
  "SO VERSION": "VERSION DEL S.O.",
  "ESTADO": "ESTADO DEL EQUIPO",
  "ESTADO EQUIPO": "ESTADO DEL EQUIPO",
  "DOMINIO": "EN QUE DOMINIO SE ENCUENTRA",
  "CONTRASENA BIOS": "CONTRASEÑA BIOS",
  "PASSWORD BIOS": "CONTRASEÑA BIOS",
  "FECHA MANTENIMIENTO": "FECHA ULTIMO MANTENIMIENTO",
  "FECHA IMPACTO": "FECHA IMPACTO MAQUINA",
  "OBSERVACIONES": "Observaciones"
};

const HARDCODED_VALIDATIONS = {
  35: ["SI", "NO", "N/A"],
  36: ["SI", "NO", "N/A"],
  37: ["SI", "NO", "N/A"],
  38: ["SI", "NO", "N/A"],
  39: ["SI", "NO", "N/A"],
  40: ["SI", "NO", "N/A"],
  45: ["SI", "NO", "N/A"]
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.queryParams = e.parameter || {};
  template.validaciones = obtenerValidacionesManuales();
  template.validacionesIndices = obtenerValidacionesIndices();
  template.mapeoSedeId = obtenerMapeoSedeId();
  return template.evaluate()
    .setTitle('CMDB SENA CCYS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Obtiene el mapeo bidireccional Sede ↔ ID desde la Hoja3.
 * Auto-descubre qué columna contiene IDs numéricos y cuál nombres de sede
 * analizando el tipo de contenido, sin depender de la posición ni el header.
 */
function obtenerMapeoSedeId() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Hoja3');
  if (!sheet) return { sedeAId: {}, idASede: {} };

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { sedeAId: {}, idASede: {} };

  // Auto-descubrir columnas por tipo de contenido
  let idCandidates = [];
  let sedeCandidates = [];

  for (let c = 0; c < data[0].length && c < 20; c++) {
    let numericValues = [];
    let textValues = [];

    for (let r = 1; r < Math.min(data.length, 25); r++) {
      const val = data[r][c];
      if (!val || val.toString().trim() === '') continue;
      const str = val.toString().trim();
      if (/^\d+$/.test(str)) numericValues.push(str);
      else textValues.push(str);
    }

    // Candidato ID: mayoría numérica, todos de 2+ dígitos
    if (numericValues.length >= 2 && textValues.length === 0) {
      const allMultiDigit = numericValues.every(function(v) { return v.length >= 2; });
      if (allMultiDigit) {
        idCandidates.push({ col: c, count: numericValues.length });
      }
    }

    // Candidato Sede: mayoría texto, longitud promedio > 3
    if (textValues.length >= 2 && numericValues.length === 0) {
      const avgLen = textValues.reduce(function(s, v) { return s + v.length; }, 0) / textValues.length;
      if (avgLen > 3) {
        sedeCandidates.push({ col: c, count: textValues.length });
      }
    }
  }

  // Elegir el candidato con más valores
  idCandidates.sort(function(a, b) { return b.count - a.count; });
  sedeCandidates.sort(function(a, b) { return b.count - a.count; });

  const colId = idCandidates.length > 0 ? idCandidates[0].col : 6;
  const colSede = sedeCandidates.length > 0 ? sedeCandidates[0].col : 7;

  Logger.log('🔍 Mapeo Sede-ID detectado: col ID=' + colId + ', col Sede=' + colSede);

  const sedeAId = {};
  const idASede = {};

  for (let i = 1; i < data.length; i++) {
    const idRaw = data[i][colId];
    const sedeRaw = data[i][colSede];

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
 * Normaliza un header para matching (mayúsculas, sin caracteres especiales, espacios unificados)
 */
function normalizarHeader(header) {
  return header.toString().toUpperCase()
    .replace(/[^A-Z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Intenta matchear un header de Hoja3 con un campo del formulario principal.
 * Usa aliases conocidos, matching exacto y matching parcial.
 */
function matchHeaderConFormField(hoja3Header) {
  const hNorm = normalizarHeader(hoja3Header);

  // 1. Alias directo
  if (HEADER_ALIASES[hNorm]) {
    return HEADER_ALIASES[hNorm];
  }

  // 2. Matching exacto contra FORM_FIELDS
  for (const formName of Object.keys(FORM_FIELDS)) {
    if (normalizarHeader(formName) === hNorm) {
      return formName;
    }
  }

  // 3. Matching parcial (uno contiene al otro)
  for (const formName of Object.keys(FORM_FIELDS)) {
    const fNorm = normalizarHeader(formName);
    if (fNorm.includes(hNorm) || hNorm.includes(fNorm)) {
      return formName;
    }
  }

  return null;
}

/**
 * Lee las validaciones dinámicamente desde la hoja 'Hoja3'.
 * Auto-descubre el mapeo de columnas leyendo los headers y matcheándolos
 * con los campos del formulario principal. No depende de índices hardcodeados.
 */
function obtenerValidacionesManuales() {
  const resultado = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Hoja3');

  if (!sheet) {
    Logger.log('⚠️ No se encontró Hoja3, usando validaciones por defecto');
    return obtenerValidacionesPorDefecto();
  }

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return resultado;

  const headers = data[0];
  const mapeoDetectado = {};
  const noMatcheados = [];

  // Descubrir mapeo leyendo headers
  for (let c = 0; c < headers.length; c++) {
    const hRaw = headers[c];
    if (!hRaw) continue;

    const formField = matchHeaderConFormField(hRaw);
    if (formField && FORM_FIELDS[formField] !== undefined) {
      mapeoDetectado[c] = FORM_FIELDS[formField];
    } else {
      noMatcheados.push(hRaw);
    }
  }

  if (noMatcheados.length > 0) {
    Logger.log('⚠️ Headers de Hoja3 no matcheados: ' + noMatcheados.join(', '));
  }

  Logger.log('🔍 Mapeo Hoja3 detectado: ' + JSON.stringify(mapeoDetectado));

  // Extraer valores únicos por columna mapeada
  Object.keys(mapeoDetectado).forEach(function(colHoja3Str) {
    const colHoja3 = parseInt(colHoja3Str);
    const indiceForm = mapeoDetectado[colHoja3];
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
  Object.keys(HARDCODED_VALIDATIONS).forEach(function(idx) {
    resultado[parseInt(idx)] = HARDCODED_VALIDATIONS[idx];
  });

  return resultado;
}

/**
 * Devuelve el array de índices que deben renderizarse como dropdowns en el frontend.
 * Se genera dinámicamente a partir de las validaciones activas.
 */
function obtenerValidacionesIndices() {
  const validaciones = obtenerValidacionesManuales();
  return Object.keys(validaciones).map(Number).sort(function(a, b) { return a - b; });
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
  const validacionesIdx = obtenerValidacionesIndices();
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
          validacionesIndices: validacionesIdx,
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
