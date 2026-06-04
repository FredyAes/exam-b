// ════════════════════════════════════════════════════════════════════
//  📊  APPS SCRIPT — EXAMEN DE BIOLOGÍA
//  Instrucciones de instalación al final de este archivo.
// ════════════════════════════════════════════════════════════════════

// ── CONFIGURACIÓN ────────────────────────────────────────────────────
// Pega aquí el ID de tu Google Sheet (está en la URL entre /d/ y /edit)
// Ejemplo: https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
const SHEET_ID   = 'PEGA_AQUI_EL_ID_DE_TU_GOOGLE_SHEET';
const SHEET_NAME = 'Respuestas';   // Nombre de la hoja (pestaña)
// ─────────────────────────────────────────────────────────────────────

// Encabezados de la hoja (se crean automáticamente si la hoja está vacía)
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

// ── RECIBIR DATOS (POST) ─────────────────────────────────────────────
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // Crear la hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Agregar encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      // Estilo de encabezados
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground('#0d3b2e');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    // ── VALIDACIONES DEL LADO DEL SERVIDOR ──────────────────────────

    // 1. Verificar que el número de lista + grupo no se hayan registrado antes
    const listaCol = 3; // columna "No. de Lista"
    const grupoCol = 4; // columna "Grupo"
    const lastRow  = sheet.getLastRow();
    if (lastRow > 1) {
      const listaVals = sheet.getRange(2, listaCol, lastRow - 1, 1).getValues().flat().map(Number);
      const grupoVals = sheet.getRange(2, grupoCol, lastRow - 1, 1).getValues().flat().map(String);
      for (let i = 0; i < listaVals.length; i++) {
        if (listaVals[i] === Number(data.numeroDeLista) && grupoVals[i] === String(data.grupo)) {
          return jsonResponse({ ok: false, error: 'DUPLICADO_LISTA',
            mensaje: `El alumno con lista #${data.numeroDeLista} del grupo ${data.grupo} ya tiene una respuesta registrada.` });
        }
      }
    }

    // 2. Validar que venga nombre (solo letras/espacios)
    if (!data.nombre || !/^[A-ZÁÉÍÓÚÑ\s]{3,}$/i.test(data.nombre)) {
      return jsonResponse({ ok: false, error: 'NOMBRE_INVALIDO',
        mensaje: 'El nombre contiene caracteres no permitidos.' });
    }

    // 3. Validar número de lista (1-50)
    const listaNum = Number(data.numeroDeLista);
    if (!data.numeroDeLista || isNaN(listaNum) || listaNum < 1 || listaNum > 50) {
      return jsonResponse({ ok: false, error: 'LISTA_INVALIDA',
        mensaje: 'El número de lista no es válido (debe ser entre 1 y 50).' });
    }

    // 4. Validar grupo (1-9)
    if (!data.grupo || !['1','2','3','4','5','6','7','8','9'].includes(String(data.grupo))) {
      return jsonResponse({ ok: false, error: 'GRUPO_INVALIDO',
        mensaje: 'El grupo no es válido.' });
    }

    // 5. Validar rango de calificación
    const cal = Number(data.calificacion);
    if (isNaN(cal) || cal < 0 || cal > 100) {
      return jsonResponse({ ok: false, error: 'CALIFICACION_INVALIDA',
        mensaje: 'La calificación recibida no es válida.' });
    }

    // ── GUARDAR FILA ────────────────────────────────────────────────
    const now = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'
    );

    const row = [
      now,
      data.nombre           || '',
      listaNum,
      data.grupo            || '',
      data.horaInicio       || '',
      data.horaTermino      || '',
      Number(data.duracionSegundos) || 0,
      cal,
      Number(data.correctas)   || 0,
      Number(data.incorrectas) || 0,
      data.respuestas       || ''
    ];

    sheet.appendRow(row);

    // ── DAR FORMATO A LA FILA RECIÉN AÑADIDA ────────────────────────
    const newRow = sheet.getLastRow();

    // Colorear según calificación
    const calCell = sheet.getRange(newRow, 8); // columna Calificación
    if (cal >= 70)      calCell.setBackground('#d6f5e8').setFontColor('#0d3b2e');
    else if (cal >= 50) calCell.setBackground('#fff3cd').setFontColor('#856404');
    else                calCell.setBackground('#fde8e8').setFontColor('#8b0000');

    calCell.setFontWeight('bold');

    // Ajustar ancho de columnas solo una vez (cuando hay 2 filas: encabezado + primera respuesta)
    if (newRow === 2) {
      const widths = [160, 200, 220, 60, 100, 110, 100, 100, 100, 110, 250];
      widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    }

    return jsonResponse({ ok: true, mensaje: 'Respuesta registrada correctamente.' });

  } catch (err) {
    return jsonResponse({ ok: false, error: 'ERROR_SERVIDOR', mensaje: err.toString() });
  }
}

// ── RESPUESTA JSON CON CORS ──────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── PRUEBA MANUAL (ejecutar desde el editor) ────────────────────────
function testDoPost() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        nombre:           'JUAN PÉREZ GARCÍA',
        numeroDeLista:    12,
        grupo:            '3',
        horaInicio:       '11:00:00',
        horaTermino:      '11:22:45',
        duracionSegundos: 1365,
        calificacion:     85,
        correctas:        17,
        incorrectas:      3,
        respuestas:       '0,0,0,0,1,1,1,1,0,0,0,1,0,1,0,1,0,1,1,1'
      })
    }
  };
  const result = doPost(mockEvent);
  Logger.log(result.getContent());
}


// ════════════════════════════════════════════════════════════════════
//
//  📋  INSTRUCCIONES DE INSTALACIÓN PASO A PASO
//  ─────────────────────────────────────────────
//
//  1. CREAR LA HOJA DE CÁLCULO
//     • Ve a https://sheets.google.com y crea una nueva hoja.
//     • Copia el ID de la URL (la cadena entre /d/ y /edit).
//     • Pégalo en la variable SHEET_ID al inicio de este script.
//
//  2. ABRIR EL EDITOR DE APPS SCRIPT
//     • En la hoja de cálculo ve a: Extensiones → Apps Script
//     • Borra el código que aparece por defecto.
//     • Pega TODO el contenido de este archivo.
//
//  3. GUARDAR
//     • Ctrl+S (o el ícono de guardar).
//     • Ponle un nombre al proyecto, ej: "Examen Biología".
//
//  4. PROBAR LOCALMENTE (opcional pero recomendado)
//     • Selecciona la función "testDoPost" en el menú desplegable.
//     • Haz clic en "Ejecutar".
//     • Acepta los permisos que pida (acceso a Sheets).
//     • Revisa los registros (Ver → Registros) y verifica que
//       aparezca {"ok":true,...}.
//     • Abre tu hoja y confirma que se creó la fila de prueba.
//
//  5. PUBLICAR COMO WEB APP
//     • Clic en "Implementar" → "Nueva implementación".
//     • Tipo: "Aplicación web".
//     • Ejecutar como: "Yo (tu correo)".
//     • Quién tiene acceso: "Cualquier usuario" (Anyone).
//     • Haz clic en "Implementar".
//     • COPIA la URL que aparece (termina en /exec).
//
//  6. PEGAR LA URL EN EL EXAMEN HTML
//     • Abre el archivo examen-biologia.html.
//     • Busca la línea:
//         const SHEET_URL = 'https://script.google.com/macros/s/TU_APPS_SCRIPT_ID/exec';
//     • Reemplaza toda la URL por la que copiaste en el paso 5.
//
//  7. SUBIR A GITHUB PAGES
//     • Crea un repositorio en GitHub (puede ser privado o público).
//     • Sube el archivo examen-biologia.html como "index.html".
//     • Ve a Settings → Pages → Branch: main → /root → Save.
//     • Tu examen estará disponible en:
//         https://TU_USUARIO.github.io/NOMBRE_REPOSITORIO/
//
//  ⚠️  NOTAS IMPORTANTES
//     • Cada vez que modifiques el Apps Script y quieras que los
//       cambios surtan efecto, debes crear una NUEVA implementación
//       (Implementar → Administrar implementaciones → editar ✏️
//       → Versión: "Nueva versión" → Implementar).
//     • La URL /exec NO cambia entre versiones, solo el código interno.
//     • Si ves errores CORS en el navegador, es normal con "no-cors";
//       los datos llegan igual, solo no puedes leer la respuesta JSON
//       desde el navegador por seguridad de Google.
//
// ════════════════════════════════════════════════════════════════════
