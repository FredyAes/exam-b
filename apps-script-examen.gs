// ════════════════════════════════════════════════════════════════════
//  📊  APPS SCRIPT — EXAMEN DE BIOLOGÍA
//  Recibe datos vía GET (parámetros en URL) desde GitHub Pages.
//  Esto evita problemas de CORS que ocurren con POST + no-cors.
// ════════════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ────────────────────────────────────────────────────
const SHEET_ID   = '10pWnT5tPmRNjzdSgXLsUcLE2nb__CZWDNCx8rq0sM3k';
const SHEET_NAME = 'Respuestas';
// ─────────────────────────────────────────────────────────────────────

const HEADERS = [
  'Fecha y hora de registro',
  'Nombre',
  'No. de Lista',
  'Grupo',
  'Hora inicio examen',
  'Hora término examen',
  'Duración (seg)',
  'Calificación (%)',
  'Respuestas correctas',
  'Respuestas incorrectas',
  'Detalle de respuestas'
];

// ── RECIBIR DATOS (GET con parámetros en URL) ────────────────────────
function doGet(e) {
  try {
    const p = e.parameter;

    // Si no vienen parámetros del examen, devolver mensaje simple
    if (!p || !p.nombre) {
      return ContentService.createTextOutput('Servicio activo.');
    }

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Crear encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      const hr = sheet.getRange(1, 1, 1, HEADERS.length);
      hr.setBackground('#0d3b2e');
      hr.setFontColor('#ffffff');
      hr.setFontWeight('bold');
      hr.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    // ── VALIDACIONES ─────────────────────────────────────────────────

    const nombre = String(p.nombre || '').trim();
    const lista  = Number(p.numeroDeLista);
    const grupo  = String(p.grupo  || '').trim();

    if (!/^[A-Za-zÁÉÍÓÚáéíóúÑñÜü\s]+$/.test(nombre)) {
      return jsonResponse({ ok: false, mensaje: 'Nombre inválido.' });
    }
    if (isNaN(lista) || lista < 1 || lista > 50) {
      return jsonResponse({ ok: false, mensaje: 'Número de lista inválido.' });
    }
    if (!grupo) {
      return jsonResponse({ ok: false, mensaje: 'Grupo inválido.' });
    }

    // ── EVITAR DUPLICADOS (lista + grupo) ────────────────────────────
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (Number(allData[i][2]) === lista && String(allData[i][3]) === grupo) {
        return jsonResponse({
          ok: false,
          mensaje: `Ya existe un registro para lista #${lista} grupo ${grupo}.`
        });
      }
    }

    // ── GUARDAR ──────────────────────────────────────────────────────
    const ahora = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'
    );
    const cal = Number(p.calificacion || 0);

    sheet.appendRow([
      ahora,
      nombre,
      lista,
      grupo,
      p.horaInicio        || '',
      p.horaTermino       || '',
      Number(p.duracionSegundos || 0),
      cal,
      Number(p.correctas  || 0),
      Number(p.incorrectas|| 0),
      p.respuestas        || ''
    ]);

    // ── FORMATO DE CALIFICACIÓN ──────────────────────────────────────
    const newRow  = sheet.getLastRow();
    const calCell = sheet.getRange(newRow, 8);
    if      (cal >= 70) calCell.setBackground('#d6f5e8').setFontColor('#0d3b2e');
    else if (cal >= 50) calCell.setBackground('#fff3cd').setFontColor('#856404');
    else                calCell.setBackground('#fde8e8').setFontColor('#8b0000');
    calCell.setFontWeight('bold');

    // Ajustar anchos solo en la primera fila de datos
    if (newRow === 2) {
      [170, 250, 90, 80, 120, 120, 100, 110, 120, 130, 280]
        .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    }

    return jsonResponse({ ok: true, mensaje: 'Registro guardado correctamente.' });

  } catch (err) {
    return jsonResponse({ ok: false, mensaje: err.toString() });
  }
}

// ── Respuesta JSON con cabecera CORS ────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── PRUEBA MANUAL (ejecutar desde el editor de Apps Script) ─────────
function testDoGet() {
  const mockEvent = {
    parameter: {
      nombre:           'JUAN PEREZ GARCIA',
      numeroDeLista:    '12',
      grupo:            '3',
      horaInicio:       '11:00:00',
      horaTermino:      '11:22:45',
      duracionSegundos: '1365',
      calificacion:     '85',
      correctas:        '17',
      incorrectas:      '3',
      respuestas:       '0,0,0,0,1,1,1,1,0,0,0,1,0,1,0,1,0,1,1,1'
    }
  };
  Logger.log(doGet(mockEvent).getContent());
}


// ════════════════════════════════════════════════════════════════════
//
//  📋  POR QUÉ SE CAMBIÓ DE doPost A doGet
//  ─────────────────────────────────────────
//  Cuando el HTML está en GitHub Pages (dominio distinto a Google),
//  el navegador bloquea POST con Content-Type: application/json
//  debido a CORS. Con mode:'no-cors' el body llega vacío a Apps Script.
//  La solución más confiable es enviar los datos como parámetros GET
//  en la URL (?nombre=...&grupo=...) — los parámetros GET sí pasan
//  con no-cors sin ningún problema.
//
//  PASOS PARA ACTUALIZAR TU IMPLEMENTACIÓN
//  ─────────────────────────────────────────
//  1. Reemplaza TODO el código en tu editor de Apps Script con este.
//  2. Guarda (Ctrl+S).
//  3. Ve a Implementar → Administrar implementaciones.
//  4. Haz clic en el lápiz ✏️ de tu implementación actual.
//  5. En "Versión" selecciona "Nueva versión".
//  6. Clic en "Implementar".
//  7. La URL /exec NO cambia — no necesitas actualizar el HTML.
//  8. Prueba ejecutando testDoGet() desde el editor.
//
//  IMPORTANTE: "Quién tiene acceso" debe ser "Cualquier usuario".
//
// ════════════════════════════════════════════════════════════════════