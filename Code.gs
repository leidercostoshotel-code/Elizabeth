// =============================================
// GOOGLE APPS SCRIPT - Control de Evidencias
// Autor: LTM Soluciones Digitales
// =============================================
//
// INSTRUCCIONES DE CONFIGURACION:
// 1. Crea un nuevo proyecto en script.google.com
// 2. Pega este codigo completo
// 3. Despliega como Web App:
//    Deploy > New Deployment > Web App
//    Execute as: Me
//    Who has access: Anyone
// 4. Copia la URL del deployment y pégala en la app HTML (variable SCRIPT_URL)
// =============================================

const SPREADSHEET_ID = '1QtJ6JWHAW29eNOVBsnszKj5c2teiFk4KoRXT91-z6DY';
const CARPETA_NOMBRE = 'Evidencias Fotograficas';
const HOJA_NOMBRE    = 'Evidencias';
const HOTEL_NOMBRE   = 'HOTEL ELIZABETH';

// ---- CABECERAS de la hoja ----
const CABECERAS = [
  'ID',
  'Fecha y Hora',
  'Responsable',
  'Área',
  'Estado',
  'Observación',
  'Latitud',
  'Longitud',
  'Foto',
  'ID Drive'
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

    const fotoInfo = guardarFotoEnDrive(payload.foto);

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
function guardarFotoEnDrive(base64) {
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
  return { id: fileId, url: 'https://drive.google.com/uc?id=' + fileId };
}

// =============================================
// Guarda el registro en Google Sheets
// =============================================
function guardarEnSheets(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja = ss.getSheetByName(HOJA_NOMBRE);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_NOMBRE);
    inicializarHoja(hoja);
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

  const fila = hoja.getLastRow();

  // Imagen visible directamente en la celda
  hoja.getRange(fila, 9).setFormula('=IMAGE("' + datos.urlFoto + '",4,100,140)');
  hoja.setRowHeight(fila, 110);

  // Alineacion vertical centrada en toda la fila
  hoja.getRange(fila, 1, 1, CABECERAS.length)
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  colorearEstado(hoja, fila, datos.estado);

  // Alternar color de fila para legibilidad
  if (fila % 2 === 0) {
    hoja.getRange(fila, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
  }

  return id;
}

// =============================================
// Inicializa la hoja con titulo y cabeceras
// =============================================
function inicializarHoja(hoja) {
  // Fila 1: Titulo
  hoja.appendRow([HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS']);
  const titulo = hoja.getRange(1, 1, 1, CABECERAS.length);
  titulo.merge()
    .setValue(HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS')
    .setBackground('#1A3C5E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(1, 55);

  // Fila 2: Cabeceras
  hoja.appendRow(CABECERAS);
  formatearCabeceras(hoja);
}

// =============================================
// Formato de cabeceras (fila 2)
// =============================================
function formatearCabeceras(hoja) {
  const rango = hoja.getRange(2, 1, 1, CABECERAS.length);
  rango
    .setBackground('#2E5F8A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setRowHeight(2, 40);
  hoja.setFrozenRows(2);

  // Anchos de columna
  hoja.setColumnWidth(1, 100);  // ID
  hoja.setColumnWidth(2, 170);  // Fecha y Hora
  hoja.setColumnWidth(3, 140);  // Responsable
  hoja.setColumnWidth(4, 140);  // Area
  hoja.setColumnWidth(5, 120);  // Estado
  hoja.setColumnWidth(6, 260);  // Observacion
  hoja.setColumnWidth(7, 90);   // Latitud
  hoja.setColumnWidth(8, 90);   // Longitud
  hoja.setColumnWidth(9, 170);  // Foto
  hoja.setColumnWidth(10, 0);   // ID Drive (oculto)
}

// =============================================
// Colores segun estado en la columna E
// =============================================
function colorearEstado(hoja, fila, estado) {
  const celda = hoja.getRange(fila, 5);
  const colores = {
    'Conforme':    { bg: '#D4EDDA', font: '#1A7F54' },
    'No conforme': { bg: '#FADBD8', font: '#C0392B' },
    'Pendiente':   { bg: '#FEF5E7', font: '#B7770D' },
    'Informativo': { bg: '#D6EAF8', font: '#1A5276' }
  };
  const c = colores[estado];
  if (c) {
    celda.setBackground(c.bg)
      .setFontColor(c.font)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
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

// =============================================
// REFORMATEAR - Corre esto una vez si ya tienes
// datos en la hoja y quieres aplicar el nuevo diseño
// =============================================
function reformatearHoja() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(HOJA_NOMBRE);
  if (!hoja) { Logger.log('Hoja no encontrada'); return; }

  // Insertar fila de titulo si no existe
  const primerValor = hoja.getRange(1, 1).getValue();
  if (primerValor !== HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS') {
    hoja.insertRowBefore(1);
    const titulo = hoja.getRange(1, 1, 1, CABECERAS.length);
    titulo.merge()
      .setValue(HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS')
      .setBackground('#1A3C5E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    hoja.setRowHeight(1, 55);
  }

  formatearCabeceras(hoja);

  // Reformatear filas de datos existentes
  const ultimaFila = hoja.getLastRow();
  for (let i = 3; i <= ultimaFila; i++) {
    hoja.setRowHeight(i, 110);
    hoja.getRange(i, 1, 1, CABECERAS.length)
      .setVerticalAlignment('middle')
      .setFontSize(10)
      .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    if (i % 2 === 0) hoja.getRange(i, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
    const estado = hoja.getRange(i, 5).getValue();
    if (estado) colorearEstado(hoja, i, estado);
  }

  hoja.setColumnWidth(10, 0);
  Logger.log('Hoja reformateada correctamente');
}
