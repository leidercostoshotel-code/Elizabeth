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

// Fila 1 = Titulo | Fila 2 = Filtro mes | Fila 3 = Cabeceras | Fila 4+ = Datos
const FILA_TITULO    = 1;
const FILA_FILTRO    = 2;
const FILA_CABECERAS = 3;
const FILA_DATOS     = 4;

const CABECERAS = [
  'ID', 'Fecha y Hora', 'Responsable', 'Estado',
  'Observación', 'Latitud', 'Longitud', 'Foto', 'ID Drive'
];

// =============================================
// CORS
// =============================================
function doOptions(e) { return buildResponse({}); }

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
      area:        payload.area || 'General',
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
// Drive
// =============================================
function obtenerCarpeta() {
  const c = DriveApp.getFoldersByName(CARPETA_NOMBRE);
  if (c.hasNext()) return c.next();
  return DriveApp.createFolder(CARPETA_NOMBRE);
}

function guardarFotoEnDrive(base64) {
  const limpio = base64.replace(/^data:image\/\w+;base64,/, '');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(limpio), 'image/jpeg',
    'evidencia_' + generarId() + '.jpg'
  );
  const file = obtenerCarpeta().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), url: 'https://drive.google.com/uc?id=' + file.getId() };
}

// =============================================
// Guardar en la pestaña del area correspondiente
// =============================================
function guardarEnSheets(datos) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const area = datos.area || 'General';
  let hoja   = ss.getSheetByName(area);

  if (!hoja) {
    hoja = ss.insertSheet(area);
    inicializarHoja(hoja, area);
  }

  // Mostrar todas las filas antes de agregar
  const ultima = hoja.getLastRow();
  if (ultima >= FILA_DATOS) hoja.showRows(FILA_DATOS, ultima - FILA_DATOS + 1);

  const id = generarId();
  hoja.appendRow([
    id, datos.fecha, datos.responsable, datos.estado,
    datos.observacion, datos.latitud, datos.longitud,
    datos.urlFoto, datos.idFoto
  ]);

  const fila = hoja.getLastRow();
  hoja.getRange(fila, 8).setFormula('=IMAGE("' + datos.urlFoto + '",4,100,140)');
  hoja.setRowHeight(fila, 110);
  hoja.getRange(fila, 1, 1, CABECERAS.length)
    .setVerticalAlignment('middle').setFontSize(10)
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  if (fila % 2 !== 0) hoja.getRange(fila, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
  colorearEstado(hoja, fila, datos.estado);

  // Re-aplicar filtro de mes activo
  aplicarFiltroMes(hoja);

  return id;
}

// =============================================
// Inicializar pestaña de area
// =============================================
function inicializarHoja(hoja, area) {
  // Fila 1: Titulo
  hoja.appendRow([HOTEL_NOMBRE + ' - ' + area.toUpperCase()]);
  hoja.getRange(FILA_TITULO, 1, 1, CABECERAS.length).merge()
    .setBackground('#1A3C5E').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(FILA_TITULO, 55);

  // Fila 2: Filtro de mes
  hoja.appendRow(['']);
  configurarFiltroMes(hoja);

  // Fila 3: Cabeceras
  hoja.appendRow(CABECERAS);
  formatearCabeceras(hoja);
}

// =============================================
// Fila de filtro por mes (fila 2)
// =============================================
function configurarFiltroMes(hoja) {
  hoja.getRange(FILA_FILTRO, 1).setValue('FILTRAR POR MES:')
    .setFontWeight('bold').setFontColor('#1A3C5E').setHorizontalAlignment('right');

  const mesesOpc = ['Todos', ...MESES];
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(mesesOpc, true).setAllowInvalid(false).build();

  hoja.getRange(FILA_FILTRO, 2).setDataValidation(regla).setValue('Todos')
    .setBackground('#EAF4FB').setFontWeight('bold').setHorizontalAlignment('center');

  hoja.getRange(FILA_FILTRO, 3).setValue('← Selecciona un mes para filtrar')
    .setFontColor('#888888').setFontStyle('italic');

  hoja.getRange(FILA_FILTRO, 1, 1, CABECERAS.length).setBackground('#D6EAF8');
  hoja.setRowHeight(FILA_FILTRO, 38);
}

// =============================================
// Filtrar filas por mes seleccionado
// =============================================
function aplicarFiltroMes(hoja) {
  const mesFiltro  = hoja.getRange(FILA_FILTRO, 2).getValue() || 'Todos';
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_DATOS) return;

  hoja.showRows(FILA_DATOS, ultimaFila - FILA_DATOS + 1);
  if (mesFiltro === 'Todos') return;

  for (let i = FILA_DATOS; i <= ultimaFila; i++) {
    const fecha = String(hoja.getRange(i, 2).getValue());
    let mesRegistro = '';
    const partes = fecha.split('/');
    if (partes.length >= 2) {
      const numMes = parseInt(partes[1]);
      if (numMes >= 1 && numMes <= 12) mesRegistro = MESES[numMes - 1];
    }
    if (mesRegistro !== mesFiltro) hoja.hideRows(i);
  }
}

// =============================================
// Trigger onEdit - filtra al cambiar el desplegable
// =============================================
function onEdit(e) {
  const hoja = e.source.getActiveSheet();
  if (e.range.getRow() === FILA_FILTRO && e.range.getColumn() === 2) {
    aplicarFiltroMes(hoja);
  }
}

// =============================================
// Cabeceras (fila 3)
// =============================================
function formatearCabeceras(hoja) {
  hoja.getRange(FILA_CABECERAS, 1, 1, CABECERAS.length)
    .setBackground('#2E5F8A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(FILA_CABECERAS, 40);
  hoja.setFrozenRows(FILA_CABECERAS);

  hoja.setColumnWidth(1, 100);  // ID
  hoja.setColumnWidth(2, 170);  // Fecha
  hoja.setColumnWidth(3, 140);  // Responsable
  hoja.setColumnWidth(4, 120);  // Estado
  hoja.setColumnWidth(5, 260);  // Observacion
  hoja.setColumnWidth(6, 90);   // Latitud
  hoja.setColumnWidth(7, 90);   // Longitud
  hoja.setColumnWidth(8, 170);  // Foto
  hoja.hideColumns(9);          // ID Drive oculto
}

// =============================================
// Colores por estado
// =============================================
function colorearEstado(hoja, fila, estado) {
  const colores = {
    'Conforme':    { bg: '#D4EDDA', font: '#1A7F54' },
    'No conforme': { bg: '#FADBD8', font: '#C0392B' },
    'Pendiente':   { bg: '#FEF5E7', font: '#B7770D' },
    'Informativo': { bg: '#D6EAF8', font: '#1A5276' }
  };
  const c = colores[estado];
  if (c) {
    hoja.getRange(fila, 4)
      .setBackground(c.bg).setFontColor(c.font)
      .setFontWeight('bold').setHorizontalAlignment('center');
  }
}

// =============================================
// Generar ID
// =============================================
function generarId() {
  return 'EV-' + new Date().getTime().toString(36).toUpperCase();
}

// =============================================
// Respuesta CORS
// =============================================
function buildResponse(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// =============================================
// REFORMATEAR - Aplica diseño a hoja existente
// Selecciona esta funcion y presiona Ejecutar
// =============================================
function reformatearHoja() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  // Aplica a todas las pestanas existentes excepto hoja inicial vacia
  const hojas = ss.getSheets().filter(h => h.getLastRow() > 0);

  hojas.forEach(hoja => {
    const val1 = String(hoja.getRange(1, 1).getValue());
    const val2 = String(hoja.getRange(2, 1).getValue());
    const val3 = String(hoja.getRange(3, 1).getValue());

    if (!val1.includes(HOTEL_NOMBRE)) {
      hoja.insertRowBefore(1);
      hoja.getRange(1, 1, 1, CABECERAS.length).merge()
        .setValue(HOTEL_NOMBRE + ' - ' + hoja.getName().toUpperCase())
        .setBackground('#1A3C5E').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(16)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      hoja.setRowHeight(1, 55);
    }

    if (val2 !== 'FILTRAR POR MES:') {
      hoja.insertRowBefore(2);
      configurarFiltroMes(hoja);
    }

    const val3nuevo = String(hoja.getRange(3, 1).getValue());
    if (val3nuevo !== 'ID') {
      hoja.insertRowBefore(3);
      hoja.getRange(3, 1, 1, CABECERAS.length).setValues([CABECERAS]);
    }

    formatearCabeceras(hoja);

    const ultimaFila = hoja.getLastRow();
    for (let i = FILA_DATOS; i <= ultimaFila; i++) {
      hoja.setRowHeight(i, 110);
      hoja.getRange(i, 1, 1, CABECERAS.length)
        .setVerticalAlignment('middle').setFontSize(10)
        .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      if (i % 2 !== 0) hoja.getRange(i, 1, 1, CABECERAS.length).setBackground('#F7F9FC');
      const estado = hoja.getRange(i, 4).getValue();
      if (estado) colorearEstado(hoja, i, estado);
    }

    Logger.log('Reformateada: ' + hoja.getName());
  });
}

// =============================================
// TEST
// =============================================
function testManual() {
  const r = guardarEnSheets({
    fecha: new Date().toLocaleString('es-PE'), responsable: 'Leider T.',
    area: 'Banquetes', estado: 'Conforme',
    observacion: 'Prueba manual desde el IDE',
    latitud: '-12.0464', longitud: '-77.0428',
    urlFoto: 'https://via.placeholder.com/400', idFoto: 'test_id'
  });
  Logger.log('Registro guardado: ' + r);
}
