// =============================================
// GOOGLE APPS SCRIPT - Control de Evidencias
// Autor: LTM Soluciones Digitales
// =============================================

const SPREADSHEET_ID = '1QtJ6JWHAW29eNOVBsnszKj5c2teiFk4KoRXT91-z6DY';
const CARPETA_NOMBRE = 'Evidencias Fotograficas';
const HOTEL_NOMBRE   = 'Swissotel Lima Peru';

const MESES = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

// Estructura: Fila 1=Titulo | Fila 2=Filtro mes | Fila 3=Cabeceras | Fila 4+=Datos
const FILA_TITULO    = 1;
const FILA_FILTRO    = 2;
const FILA_CABECERAS = 3;
const FILA_DATOS     = 4;

// Columnas identicas al formato Walk Through
const CABECERAS = [
  'AREA',
  'DESCRIPCION DE HALLAZGO',
  'REGISTRO FOTOGRAFICO',
  'RESPONSABLE',
  'DEPARTAMENTO',
  'Comentario - Respuesta del area responsable',
  'REGISTRO FOTOGRAFICO DEL LEVANTAMIENTO',
  'ESTADO'
];

// Columnas internas (ocultas)
const COL_ID_INTERNO   = 9;
const COL_FECHA        = 10;

// =============================================
function doOptions(e) { return buildResponse({}); }

// =============================================
// POST
// =============================================
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const fotoInfo = guardarFotoEnDrive(p.foto);
    const id = guardarEnSheets({
      fecha:        p.fecha || new Date().toLocaleString('es-PE'),
      area:         p.area || '',
      descripcion:  p.descripcion || '',
      responsable:  p.responsable || '',
      departamento: p.departamento || '',
      estado:       p.estado || '',
      urlFoto:      fotoInfo.url,
      idFoto:       fotoInfo.id
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
// Guardar en la pestaña del area
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

  // Columnas: AREA | DESC HALLAZGO | FOTO | RESPONSABLE | DEPTO | COMENTARIO | FOTO_LEV | ESTADO | ID | FECHA
  hoja.appendRow([
    datos.area,
    datos.descripcion,
    datos.urlFoto,      // se reemplaza por IMAGE() abajo
    datos.responsable,
    datos.departamento,
    '',                 // Comentario - se llena manualmente en el sheet
    '',                 // Foto levantamiento - se llena manualmente
    datos.estado,
    id,
    datos.fecha
  ]);

  const fila = hoja.getLastRow();

  // Imagen visible en columna 3
  hoja.getRange(fila, 3).setFormula('=IMAGE("' + datos.urlFoto + '",4,100,140)');
  hoja.setRowHeight(fila, 120);

  // Formato de la fila
  hoja.getRange(fila, 1, 1, 8)
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  if (fila % 2 !== 0) hoja.getRange(fila, 1, 1, 8).setBackground('#F7F9FC');

  colorearEstado(hoja, fila, datos.estado);
  aplicarFiltroMes(hoja);

  return id;
}

// =============================================
// Inicializar pestaña nueva
// =============================================
function inicializarHoja(hoja, area) {
  const mesActual = MESES[new Date().getMonth()] + ' ' + new Date().getFullYear();

  // Fila 1: Titulo estilo Walk Through
  hoja.appendRow([HOTEL_NOMBRE + ' - WALK THROUGH - ' + mesActual.toUpperCase()]);
  hoja.getRange(FILA_TITULO, 1, 1, 8).merge()
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
// Filtro de mes (fila 2)
// =============================================
function configurarFiltroMes(hoja) {
  hoja.getRange(FILA_FILTRO, 1).setValue('FILTRAR POR MES:')
    .setFontWeight('bold').setFontColor('#FFFFFF').setHorizontalAlignment('right');

  const mesesOpc = ['Todos', ...MESES];
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(mesesOpc, true).setAllowInvalid(false).build();

  hoja.getRange(FILA_FILTRO, 2).setDataValidation(regla).setValue('Todos')
    .setBackground('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  hoja.getRange(FILA_FILTRO, 3).setValue('← Selecciona un mes')
    .setFontColor('#AAAAAA').setFontStyle('italic');

  hoja.getRange(FILA_FILTRO, 1, 1, 8).setBackground('#7F8C8D');
  hoja.setRowHeight(FILA_FILTRO, 35);
}

// =============================================
// Aplicar filtro por mes
// =============================================
function aplicarFiltroMes(hoja) {
  const mesFiltro  = hoja.getRange(FILA_FILTRO, 2).getValue() || 'Todos';
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_DATOS) return;

  hoja.showRows(FILA_DATOS, ultimaFila - FILA_DATOS + 1);
  if (mesFiltro === 'Todos') return;

  for (let i = FILA_DATOS; i <= ultimaFila; i++) {
    const fecha = String(hoja.getRange(i, COL_FECHA).getValue());
    const partes = fecha.split('/');
    let mesRegistro = '';
    if (partes.length >= 2) {
      const n = parseInt(partes[1]);
      if (n >= 1 && n <= 12) mesRegistro = MESES[n - 1];
    }
    if (mesRegistro !== mesFiltro) hoja.hideRows(i);
  }
}

// =============================================
// onEdit trigger
// =============================================
function onEdit(e) {
  const hoja = e.source.getActiveSheet();
  if (e.range.getRow() === FILA_FILTRO && e.range.getColumn() === 2) {
    aplicarFiltroMes(hoja);
    actualizarTitulo(hoja);
  }
}

// =============================================
// Actualiza el titulo segun el mes seleccionado
// =============================================
function actualizarTitulo(hoja) {
  const mes = hoja.getRange(FILA_FILTRO, 2).getValue();
  const anio = new Date().getFullYear();
  const titulo = mes === 'Todos'
    ? HOTEL_NOMBRE + ' - WALK THROUGH - ' + anio
    : HOTEL_NOMBRE + ' - WALK THROUGH - ' + mes.toUpperCase() + ' ' + anio;
  hoja.getRange(FILA_TITULO, 1).setValue(titulo);
}

// =============================================
// Formato cabeceras (fila 3) - identico al Walk Through
// =============================================
function formatearCabeceras(hoja) {
  hoja.getRange(FILA_CABECERAS, 1, 1, 8)
    .setBackground('#1A3C5E').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  hoja.setRowHeight(FILA_CABECERAS, 50);
  hoja.setFrozenRows(FILA_CABECERAS);

  // Anchos de columna = Walk Through
  hoja.setColumnWidth(1, 140);  // AREA
  hoja.setColumnWidth(2, 220);  // DESCRIPCION DE HALLAZGO
  hoja.setColumnWidth(3, 180);  // REGISTRO FOTOGRAFICO
  hoja.setColumnWidth(4, 130);  // RESPONSABLE
  hoja.setColumnWidth(5, 130);  // DEPARTAMENTO
  hoja.setColumnWidth(6, 220);  // COMENTARIO RESPUESTA
  hoja.setColumnWidth(7, 180);  // FOTO LEVANTAMIENTO
  hoja.setColumnWidth(8, 110);  // ESTADO

  // Ocultar columnas internas
  hoja.hideColumns(9);
  hoja.hideColumns(10);
}

// =============================================
// Colores por estado (columna 8)
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
    hoja.getRange(fila, 8)
      .setBackground(c.bg).setFontColor(c.font)
      .setFontWeight('bold').setHorizontalAlignment('center');
  }
}

// =============================================
// ID y CORS
// =============================================
function generarId() {
  return 'EV-' + new Date().getTime().toString(36).toUpperCase();
}

function buildResponse(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// =============================================
// REFORMATEAR - Ejecuta una vez desde el IDE
// =============================================
function reformatearHoja() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojas = ss.getSheets().filter(h => h.getLastRow() > 0);
  const mesActual = MESES[new Date().getMonth()] + ' ' + new Date().getFullYear();

  hojas.forEach(hoja => {
    const v1 = String(hoja.getRange(1, 1).getValue());

    if (!v1.includes('WALK THROUGH') && !v1.includes(HOTEL_NOMBRE)) {
      hoja.insertRowBefore(1);
      hoja.getRange(1, 1, 1, 8).merge()
        .setValue(HOTEL_NOMBRE + ' - WALK THROUGH - ' + mesActual.toUpperCase())
        .setBackground('#1A3C5E').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(16)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      hoja.setRowHeight(1, 55);
    }

    if (String(hoja.getRange(2, 1).getValue()) !== 'FILTRAR POR MES:') {
      hoja.insertRowBefore(2);
      configurarFiltroMes(hoja);
    }

    if (String(hoja.getRange(3, 1).getValue()) !== 'AREA') {
      hoja.insertRowBefore(3);
      hoja.getRange(3, 1, 1, 8).setValues([CABECERAS]);
    }

    formatearCabeceras(hoja);

    const ultimaFila = hoja.getLastRow();
    for (let i = FILA_DATOS; i <= ultimaFila; i++) {
      hoja.setRowHeight(i, 120);
      hoja.getRange(i, 1, 1, 8)
        .setVerticalAlignment('middle').setFontSize(10).setWrap(true)
        .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      if (i % 2 !== 0) hoja.getRange(i, 1, 1, 8).setBackground('#F7F9FC');
      const estado = hoja.getRange(i, 8).getValue();
      if (estado) colorearEstado(hoja, i, estado);
    }

    Logger.log('Reformateada: ' + hoja.getName());
  });
}

// =============================================
// TEST
// =============================================
function testManual() {
  guardarEnSheets({
    fecha: new Date().toLocaleString('es-PE'), area: 'Banquetes',
    descripcion: 'Cortinas sucias en salon principal',
    responsable: 'Leider T.', departamento: 'Housekeeping',
    estado: 'Pendiente',
    urlFoto: 'https://via.placeholder.com/400', idFoto: 'test_id'
  });
  Logger.log('Test completado');
}
