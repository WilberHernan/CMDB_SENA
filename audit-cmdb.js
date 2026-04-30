const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'CMDB CAUCA 2026.xlsx');

if (!fs.existsSync(filePath)) {
  console.error('❌ No se encontró:', filePath);
  process.exit(1);
}

const workbook = XLSX.readFile(filePath);

// ─── Helpers ─────────────────────────────────────────────────────
function isEmpty(val) {
  if (val === undefined || val === null) return true;
  if (typeof val === 'string' && val.trim() === '') return true;
  return false;
}

function normalizePlaca(val) {
  if (isEmpty(val)) return '';
  return val.toString().replace(/'/g, '-').replace(/[^a-zA-Z0-9\-]/gi, '').toUpperCase();
}

function excelDateToJSDate(serial) {
  if (!serial || isNaN(serial)) return null;
  // Excel cuenta días desde 30/12/1899 (con bug del año bisiesto 1900)
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const fractionalDay = serial - Math.floor(serial) + 0.0000001;
  let totalSeconds = Math.floor(86400 * fractionalDay);
  const seconds = totalSeconds % 60;
  totalSeconds -= seconds;
  const hours = Math.floor(totalSeconds / (60 * 60));
  const minutes = Math.floor(totalSeconds / 60) % 60;
  return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate(), hours, minutes, seconds);
}

function formatDate(serial) {
  const d = excelDateToJSDate(serial);
  if (!d) return null;
  return d.toLocaleDateString('es-CO');
}

// ─── Configuración de hojas a auditar ────────────────────────────
const EQUIPOS_SHEETS = ['EquiposSena', 'EquiposTelefonica', 'Impresoras'];

const REPORT = {
  archivo: path.basename(filePath),
  fechaAuditoria: new Date().toLocaleString('es-CO'),
  hojas: {},
  duplicadosGlobal: {},
  resumen: {}
};

// ─── Procesar cada hoja ──────────────────────────────────────────
EQUIPOS_SHEETS.forEach(sheetName => {
  if (!workbook.SheetNames.includes(sheetName)) {
    console.log(`⚠️  Hoja no encontrada: ${sheetName}`);
    return;
  }

  const worksheet = workbook.Sheets[sheetName];
  const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  if (rawData.length === 0) return;

  const headers = rawData[0].map(h => h?.toString().trim() || '');
  const rows = rawData.slice(1).filter(r => r.some(c => !isEmpty(c))); // ignorar filas totalmente vacías

  const stats = {
    totalFilas: rows.length,
    placasDuplicadas: [],
    filasConCamposVacios: [],
    fechasNumericas: [],
    porSede: {},
    porEstado: {},
    porPropietario: {},
    porASS: {}
  };

  const placasVistas = new Map(); // placaNormalizada -> [filaOriginal, ...]

  rows.forEach((row, idx) => {
    const filaReal = idx + 2; // +1 por header, +1 porque idx empieza en 0

    // Mapear columnas por índice
    const getCol = (name) => {
      const i = headers.indexOf(name);
      return i >= 0 ? row[i] : undefined;
    };

    const placaRaw = getCol('PLACA');
    const placaNorm = normalizePlaca(placaRaw);

    // --- Duplicados de placa ---
    if (placaNorm) {
      if (!placasVistas.has(placaNorm)) placasVistas.set(placaNorm, []);
      placasVistas.get(placaNorm).push({ hoja: sheetName, fila: filaReal, placaOriginal: placaRaw });
    }

    // --- Campos obligatorios vacíos ---
    const obligatorios = ['HOSTNAME', 'MARCA', 'MODELO', 'SERIAL', 'PLACA'];
    const vacios = obligatorios.filter(name => isEmpty(getCol(name)));
    if (vacios.length > 0) {
      stats.filasConCamposVacios.push({ fila: filaReal, vacios, placa: placaRaw || '(sin placa)' });
    }

    // --- Fechas numéricas ---
    ['FECHA ULTIMO MANTENIMIENTO', 'FECHA IMPACTO MAQUINA'].forEach(fechaCol => {
      const val = getCol(fechaCol);
      if (typeof val === 'number' && val > 30000) {
        stats.fechasNumericas.push({
          fila: filaReal,
          columna: fechaCol,
          valorSerial: val,
          convertida: formatDate(val)
        });
      }
    });

    // --- Agregaciones ---
    const sede = getCol('NOMBRE DE LA SEDE') || 'Sin sede';
    const estado = getCol('ESTADO DEL EQUIPO') || getCol('Estado - Observaciones') || 'Sin estado';
    const propietario = getCol('PROPIETARIO') || 'Sin propietario';
    const ass = getCol('ASS') || 'Sin ASS';

    stats.porSede[sede] = (stats.porSede[sede] || 0) + 1;
    stats.porEstado[estado] = (stats.porEstado[estado] || 0) + 1;
    stats.porPropietario[propietario] = (stats.porPropietario[propietario] || 0) + 1;
    stats.porASS[ass] = (stats.porASS[ass] || 0) + 1;
  });

  // --- Resolver duplicados ---
  placasVistas.forEach((ocurrencias, placa) => {
    if (ocurrencias.length > 1) {
      stats.placasDuplicadas.push({ placa, ocurrencias });
    }
  });

  REPORT.hojas[sheetName] = stats;
});

// ─── Duplicados entre hojas (global) ─────────────────────────────
const todasLasPlacas = new Map();
EQUIPOS_SHEETS.forEach(sheetName => {
  const hoja = REPORT.hojas[sheetName];
  if (!hoja) return;
  hoja.placasDuplicadas.forEach(dup => {
    if (!todasLasPlacas.has(dup.placa)) todasLasPlacas.set(dup.placa, []);
    dup.ocurrencias.forEach(oc => {
      if (!todasLasPlacas.get(dup.placa).some(x => x.hoja === oc.hoja && x.fila === oc.fila)) {
        todasLasPlacas.get(dup.placa).push(oc);
      }
    });
  });
});

REPORT.duplicadosGlobal = Array.from(todasLasPlacas.entries())
  .filter(([placa, ocs]) => ocs.length > 1)
  .map(([placa, ocs]) => ({ placa, ocs }));

// ─── Resumen global ──────────────────────────────────────────────
REPORT.resumen.totalEquipos = Object.values(REPORT.hojas).reduce((s, h) => s + h.totalFilas, 0);
REPORT.resumen.totalDuplicados = REPORT.duplicadosGlobal.length;
REPORT.resumen.totalFilasConVacios = Object.values(REPORT.hojas).reduce((s, h) => s + h.filasConCamposVacios.length, 0);
REPORT.resumen.totalFechasNumericas = Object.values(REPORT.hojas).reduce((s, h) => s + h.fechasNumericas.length, 0);

// ─── Output ──────────────────────────────────────────────────────
console.log(`\n╔══════════════════════════════════════════════════════════════╗`);
console.log(`║        AUDITORÍA CMDB — ${REPORT.fechaAuditoria.padEnd(31)}║`);
console.log(`╚══════════════════════════════════════════════════════════════╝\n`);

console.log(`📁 Archivo: ${REPORT.archivo}`);
console.log(`📊 Total equipos: ${REPORT.resumen.totalEquipos}`);
console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

Object.entries(REPORT.hojas).forEach(([name, s]) => {
  console.log(`\n📄 ${name}`);
  console.log(`   Filas con datos: ${s.totalFilas}`);

  if (s.placasDuplicadas.length > 0) {
    console.log(`   ❌ Placas duplicadas: ${s.placasDuplicadas.length}`);
    s.placasDuplicadas.forEach(d => {
      console.log(`      - "${d.placa}" aparece ${d.ocurrencias.length} veces:`);
      d.ocurrencias.forEach(o => console.log(`         • ${o.hoja} fila ${o.fila}`));
    });
  } else {
    console.log(`   ✅ Sin placas duplicadas`);
  }

  if (s.filasConCamposVacios.length > 0) {
    console.log(`   ⚠️  Filas con campos obligatorios vacíos: ${s.filasConCamposVacios.length}`);
    s.filasConCamposVacios.slice(0, 5).forEach(f => {
      console.log(`      • Fila ${f.fila} (placa: ${f.placa}) — faltan: ${f.vacios.join(', ')}`);
    });
    if (s.filasConCamposVacios.length > 5) {
      console.log(`      ... y ${s.filasConCamposVacios.length - 5} más`);
    }
  } else {
    console.log(`   ✅ Todas las filas tienen campos obligatorios`);
  }

  if (s.fechasNumericas.length > 0) {
    console.log(`   📅 Fechas en formato numérico (serial Excel): ${s.fechasNumericas.length}`);
    s.fechasNumericas.slice(0, 3).forEach(f => {
      console.log(`      • Fila ${f.fila} [${f.columna}]: ${f.valorSerial} → ${f.convertida}`);
    });
    if (s.fechasNumericas.length > 3) {
      console.log(`      ... y ${s.fechasNumericas.length - 3} más`);
    }
  }

  console.log(`\n   🏢 Por sede:`);
  Object.entries(s.porSede)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .forEach(([sede, count]) => console.log(`      ${sede}: ${count}`));

  console.log(`   🏷️  Por estado:`);
  Object.entries(s.porEstado)
    .sort((a, b) => b[1] - a[1])
    .forEach(([estado, count]) => console.log(`      ${estado}: ${count}`));
});

console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
console.log(`\n🌍 DUPLICADOS GLOBALES (entre hojas): ${REPORT.resumen.totalDuplicados}`);
if (REPORT.duplicadosGlobal.length > 0) {
  REPORT.duplicadosGlobal.forEach(d => {
    console.log(`   ❌ "${d.placa}" aparece en:`);
    d.ocs.forEach(o => console.log(`      • ${o.hoja} fila ${o.fila}`));
  });
} else {
  console.log(`   ✅ No hay placas repetidas entre hojas`);
}

console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
console.log(`\n📋 RESUMEN DE PROBLEMAS`);
console.log(`   • Duplicados de placa: ${REPORT.resumen.totalDuplicados}`);
console.log(`   • Filas con campos vacíos: ${REPORT.resumen.totalFilasConVacios}`);
console.log(`   • Fechas sin convertir: ${REPORT.resumen.totalFechasNumericas}`);

const tieneProblemas = REPORT.resumen.totalDuplicados > 0 || REPORT.resumen.totalFilasConVacios > 0 || REPORT.resumen.totalFechasNumericas > 0;
if (tieneProblemas) {
  console.log(`\n   ⚠️  Se encontraron problemas que deberías revisar antes de sincronizar.`);
} else {
  console.log(`\n   ✅ Todo limpio. El Excel está listo para sincronizar con Google Sheets.`);
}

// Guardar reporte JSON por si lo querés procesar después
const reportPath = path.join(__dirname, 'cmdb-audit-report.json');
fs.writeFileSync(reportPath, JSON.stringify(REPORT, null, 2));
console.log(`\n💾 Reporte detallado guardado en: ${reportPath}\n`);
