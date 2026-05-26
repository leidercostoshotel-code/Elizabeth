// =============================================
// GOOGLE APPS SCRIPT - Control de Evidencias
// Autor: LTM Soluciones Digitales
// =============================================
//
// INSTRUCCIONES DE CONFIGURACION:
// 1. Crea un nuevo proyecto en script.google.com
// 2. Pega este codigo completo
// 3. Cambia SPREADSHEET_ID por el ID de tu Google Sheets
// 4. Despliega como Web App:
//    Deploy > New Deployment > Web App
//    Execute as: Me
//    Who has access: Anyone
// 5. Copia la URL del deployment y pégala en la app HTML (variable SCRIPT_URL)
// =============================================

const SPREADSHEET_ID = '1QtJ6JWHAW29eNOVBsnszKj5c2teiFk4KoRXT91-z6DY';
const CARPETA_NOMBRE = 'Evidencias Fotograficas';
const HOJA_NOMBRE    = 'Evidencias';

// ---- CABECERAS de la hoja ----
const CABECERAS = [
  'ID',
  'Fecha y Hora',
  'Responsable',
  'Area',
  'Estado',
  'Observacion',
  'Latitud',
  'Longitud',
  'URL Foto',
  'ID Foto Drive'
];

// =============================================
// CORS - Respuesta OPTIONS (preflight)
// =============================================
function doOptions(e) {
  return buildResponse({});
}

// =============================================
// POST - Recibe datos desde la app HTML
// =============================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Guardar foto en Drive
    const fotoInfo = guardarFotoEnDrive(payload.foto, payload.fecha);

    // Guardar registro en Sheets
    const id = guardarEnSheets({
      fecha:       payload.fecha || new Date().toLocaleString('es-PE'),
      responsable: payload.responsable || '',
      area:        payload.area || '',
      estado:      payload.estado || '',
      observacion: payload.observacion || '',
      latitud:     payload.latitud || '',
      longitud:    payload.longitud || '',
      urlFoto:     fotoInfo.url,
      idFoto:      fotoInfo.id
    });

    return buildResponse({ resultado: 'ok', id: id, urlFoto: fotoInfo.url });

  } catch (err) {
    return buildResponse({ resultado: 'error', error: err.message });
  }
}

// =============================================
// Obtiene o crea la carpeta de fotos en Drive
// =============================================
function obtenerCarpeta() {
  const carpetas = DriveApp.getFoldersByName(CARPETA_NOMBRE);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(CARPETA_NOMBRE);
}

// =============================================
// Guarda la foto en Google Drive
// =============================================
function guardarFotoEnDrive(base64, fecha) {
  const base64Limpio = base64.replace(/^data:image\/\w+;base64,/, '');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Limpio),
    'image/jpeg',
    'evidencia_' + generarId() + '.jpg'
  );

  const folder = obtenerCarpeta();
  const file = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const urlFoto = 'https://drive.google.com/uc?id=' + fileId;

  return { id: fileId, url: urlFoto };
}

// =============================================
// Guarda el registro en Google Sheets
// =============================================
function guardarEnSheets(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja = ss.getSheetByName(HOJA_NOMBRE);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_NOMBRE);
    hoja.appendRow(CABECERAS);
    formatearCabeceras(hoja);
  }

  const id = generarId();

  hoja.appendRow([
    id,
    datos.fecha,
    datos.responsable,
    datos.area,
    datos.estado,
    datos.observacion,
    datos.latitud,
    datos.longitud,
    datos.urlFoto,
    datos.idFoto
  ]);

  const ultimaFila = hoja.getLastRow();
  colorearEstado(hoja, ultimaFila, datos.estado);

  return id;
}

// =============================================
// Colores segun estado en la columna E
// =============================================
function colorearEstado(hoja, fila, estado) {
  const celda = hoja.getRange(fila, 5);
  const colores = {
    'Conforme':    { bg: '#E6F7F0', font: '#1A7F54' },
    'No conforme': { bg: '#FCF0EF', font: '#C0392B' },
    'Pendiente':   { bg: '#FEF5E7', font: '#B7770D' },
    'Informativo': { bg: '#EEF4FB', font: '#1A3C5E' }
  };
  const c = colores[estado];
  if (c) {
    celda.setBackground(c.bg).setFontColor(c.font).setFontWeight('bold');
  }
}

// =============================================
// Formato de cabeceras
// =============================================
function formatearCabeceras(hoja) {
  const rango = hoja.getRange(1, 1, 1, CABECERAS.length);
  rango
    .setBackground('#1A3C5E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.setColumnWidths(1, CABECERAS.length, 160);
  hoja.setColumnWidth(1, 80);
  hoja.setColumnWidth(2, 180);
  hoja.setColumnWidth(9, 220);
  hoja.setColumnWidth(6, 280);
}

// =============================================
// Generar ID unico
// =============================================
function generarId() {
  return 'EV-' + new Date().getTime().toString(36).toUpperCase();
}

// =============================================
// Respuesta con CORS
// =============================================
function buildResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// =============================================
// TEST - Ejecuta desde el IDE de Apps Script
// =============================================
function testManual() {
  const resultado = guardarEnSheets({
    fecha:       new Date().toLocaleString('es-PE'),
    responsable: 'Leider T.',
    area:        'Almacen',
    estado:      'Conforme',
    observacion: 'Prueba manual desde el IDE',
    latitud:     '-12.0464',
    longitud:    '-77.0428',
    urlFoto:     'https://via.placeholder.com/400',
    idFoto:      'test_id'
  });
  Logger.log('Registro guardado: ' + resultado);
}
