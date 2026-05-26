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
const HOTEL_NOMBRE   = 'HOTEL ELIZABETH';

const MESES = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

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
// Nombre de la hoja segun mes y año actual
// =============================================
function nombreHojaMes() {
  const ahora = new Date();
  return MESES[ahora.getMonth()] + ' ' + ahora.getFullYear();
}

// =============================================
// Guarda el registro en la hoja del mes actual
// =============================================
function guardarEnSheets(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const nombreHoja = nombreHojaMes();
  let hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    inicializarHoja(hoja, nombreHoja);
    // Mover la hoja al inicio
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(1);
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

  hoja.getRange(fila, 9).setFormula('=IMAGE("' + datos.urlFoto + '",4,100,140)');
  hoja.setRowHeight(fila, 110);

  hoja.getRange(fila, 1, 1, CABECERAS.length)
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  if (fila % 2 !== 0) {
    hoja.getRange(fila, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
  }

  colorearEstado(hoja, fila, datos.estado);

  return id;
}

// =============================================
// Inicializa hoja con titulo y cabeceras
// =============================================
function inicializarHoja(hoja, nombreHoja) {
  // Fila 1: Titulo con mes
  hoja.appendRow([HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS - ' + nombreHoja.toUpperCase()]);
  const titulo = hoja.getRange(1, 1, 1, CABECERAS.length);
  titulo.merge()
    .setValue(HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS - ' + nombreHoja.toUpperCase())
    .setBackground('#1A3C5E')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(15)
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

  hoja.setColumnWidth(1, 100);
  hoja.setColumnWidth(2, 170);
  hoja.setColumnWidth(3, 140);
  hoja.setColumnWidth(4, 140);
  hoja.setColumnWidth(5, 120);
  hoja.setColumnWidth(6, 260);
  hoja.setColumnWidth(7, 90);
  hoja.setColumnWidth(8, 90);
  hoja.setColumnWidth(9, 170);
  hoja.hideColumns(10);
}

// =============================================
// Colores segun estado
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
// REFORMATEAR - Aplica diseño a hoja existente
// Selecciona esta funcion y presiona Ejecutar
// =============================================
function reformatearHoja() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const nombreHoja = nombreHojaMes();
  let hoja = ss.getSheetByName(nombreHoja);

  // Si no existe la hoja del mes, busca 'Evidencias' (hoja anterior)
  if (!hoja) hoja = ss.getSheetByName('Evidencias');
  if (!hoja) { Logger.log('No se encontro ninguna hoja'); return; }

  // Renombrar si es la hoja antigua
  if (hoja.getName() === 'Evidencias') {
    hoja.setName(nombreHoja);
  }

  // Insertar titulo si no existe
  const primerValor = String(hoja.getRange(1, 1).getValue());
  if (!primerValor.includes('REGISTRO DE EVIDENCIAS')) {
    hoja.insertRowBefore(1);
    const titulo = hoja.getRange(1, 1, 1, CABECERAS.length);
    titulo.merge()
      .setValue(HOTEL_NOMBRE + ' - REGISTRO DE EVIDENCIAS - ' + nombreHoja.toUpperCase())
      .setBackground('#1A3C5E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(15)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    hoja.setRowHeight(1, 55);
  }

  // Insertar cabeceras si no existe la fila
  const segundoValor = String(hoja.getRange(2, 1).getValue());
  if (segundoValor !== 'ID') {
    hoja.insertRowBefore(2);
    hoja.getRange(2, 1, 1, CABECERAS.length).setValues([CABECERAS]);
  }

  formatearCabeceras(hoja);

  // Reformatear filas de datos
  const ultimaFila = hoja.getLastRow();
  for (let i = 3; i <= ultimaFila; i++) {
    hoja.setRowHeight(i, 110);
    hoja.getRange(i, 1, 1, CABECERAS.length)
      .setVerticalAlignment('middle')
      .setFontSize(10)
      .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    if (i % 2 !== 0) hoja.getRange(i, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
    const estado = hoja.getRange(i, 5).getValue();
    if (estado) colorearEstado(hoja, i, estado);
  }

  Logger.log('Hoja reformateada: ' + nombreHoja);
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
