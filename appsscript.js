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
  'AREA', 'DESCRIPCION DE HALLAZGO', 'REGISTRO FOTOGRAFICO',
  'RESPONSABLE', 'PARTICIPANTES', 'DEPARTAMENTO',
  'Comentario - Respuesta del area responsable',
  'REGISTRO FOTOGRAFICO DEL LEVANTAMIENTO', 'ESTADO',
  'CÓDIGO', 'QR'
];

// Columnas visibles: 1-11 | Ocultas: 12 (fecha)
const COL_ID_INTERNO = 10;  // CÓDIGO legible visible
const COL_QR         = 11;  // QR image visible
const COL_FECHA      = 12;  // Fecha oculta

// Abreviaciones para código legible (2 letras por subárea)
const ABREV = (nombre) => {
  const limpio = nombre.replace(/[^A-Za-z]/g, '');
  return limpio.substring(0, 2).toUpperCase() || 'XX';
};

function generarCodigo(subarea, hoja) {
  const prefijo = ABREV(subarea);
  const filas   = Math.max(0, hoja.getLastRow() - FILA_DATOS + 1);
  const seq     = String(filas + 1).padStart(3, '0');
  const hoy     = new Date();
  const dd  = String(hoy.getDate()).padStart(2, '0');
  const mm  = String(hoy.getMonth() + 1).padStart(2, '0');
  const yyyy = hoy.getFullYear();
  return prefijo + seq + '-' + dd + mm + yyyy;  // ej. BA001-27052026
}

// =============================================
function doOptions(e) { return buildResponse({}); }

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    if (action === 'buscar') {
      const codigo = (e.parameter.codigo || '').trim().toUpperCase();
      if (!codigo) return buildResponse({ resultado: 'error', error: 'Código requerido' });
      const encontrado = buscarPorCodigo(codigo);
      return buildResponse(encontrado
        ? { resultado: 'ok', ...encontrado }
        : { resultado: 'no_encontrado' });
    }
    if (action === 'exportExcel') {
      const tipo = (e.parameter.tipo || 'completo').toLowerCase(); // completo|dia|mes|rango
      const fecha = e.parameter.fecha || '';
      const mes = e.parameter.mes || '';
      const desde = e.parameter.desde || '';
      const hasta = e.parameter.hasta || '';
      const out = exportarInformeExcel({ tipo, fecha, mes, desde, hasta });
      return buildResponse({ resultado: 'ok', ...out });
    }
    return buildResponse({ resultado: 'ok', mensaje: 'Servicio activo' });
  } catch (err) {
    return buildResponse({ resultado: 'error', error: err.message });
  }
}

// =============================================
// POST
// =============================================
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    // Acción: registrar usuario (nombre + email + sección + PIN para recuperación multi-dispositivo)
    if (p.action === 'registrarUsuario') {
      return buildResponse(registrarUsuarioEnSheets(p.nombre || '', p.seccion || '', p.pin || '', p.email || ''));
    }

    // Acción: recuperar perfil desde otro dispositivo (verifica nombre + PIN)
    if (p.action === 'recuperarPerfil') {
      return buildResponse(recuperarPerfilDesdeSheets(p.email || '', p.pin || ''));
    }

    // Acción: levantar observación existente
    if (p.action === 'levantar') {
      const out = levantarObservacion(
        (p.codigo || '').trim().toUpperCase(),
        p.comentario || '',
        p.foto || '',
        p.estado || 'Corregido'
      );
      return buildResponse(out);
    }

    const fotoInfo = guardarFotoEnDrive(p.foto);
    const id = guardarEnSheets({
      fecha:        p.fecha || new Date().toLocaleString('es-PE'),
      area:         p.area || 'General',
      subarea:      p.subarea || '',
      descripcion:  p.descripcion || '',
      responsable:   p.responsable || '',
      participantes: p.participantes || '',
      departamento:  p.departamento || '',
      estado:        p.estado || '',
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

  const codigo = generarCodigo(datos.subarea || datos.area, hoja);

  // Cols: AREA|DESC|FOTO|RESP|PART|DEPTO|COMMENT|FOTO_LEV|ESTADO|CÓDIGO|QR_placeholder|FECHA
  hoja.appendRow([
    datos.subarea, datos.descripcion, datos.urlFoto,
    datos.responsable, datos.participantes, datos.departamento,
    '', '', datos.estado,
    codigo, '', datos.fecha
  ]);

  const fila = hoja.getLastRow();

  // Imagen evidencia (col 3)
  hoja.getRange(fila, 3).setFormula('=IMAGE("' + datos.urlFoto + '",4,130,200)');
  // QR legible (col 11)
  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=90x90&data=' + encodeURIComponent(codigo);
  hoja.getRange(fila, COL_QR).setFormula('=IMAGE("' + qrUrl + '",4,80,80)');

  hoja.setRowHeight(fila, 150);

  hoja.getRange(fila, 1, 1, 9)
    .setVerticalAlignment('middle').setFontSize(10).setWrap(true)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Código: negrita centrado
  hoja.getRange(fila, COL_ID_INTERNO)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setFontSize(10);

  if (fila % 2 !== 0) hoja.getRange(fila, 1, 1, 9).setBackground('#F7F9FC');

  colorearEstado(hoja, fila, datos.estado);
  aplicarFiltroMes(hoja);

  return codigo;
}

// =============================================
// Inicializar pestaña nueva
// =============================================
function inicializarHoja(hoja, area) {
  const mesActual = MESES[new Date().getMonth()] + ' ' + new Date().getFullYear();

  // Fila 1: Titulo estilo Walk Through
  hoja.appendRow([HOTEL_NOMBRE + ' - WALK THROUGH - ' + mesActual.toUpperCase()]);
  hoja.getRange(FILA_TITULO, 1, 1, 11).merge()
    .setBackground('#7B1827').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  hoja.setRowHeight(FILA_TITULO, 65);

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

  hoja.getRange(FILA_FILTRO, 1, 1, 11).setBackground('#7F8C8D');
  hoja.setRowHeight(FILA_FILTRO, 40);
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
  hoja.getRange(FILA_CABECERAS, 1, 1, 11)
    .setBackground('#7B1827').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  hoja.setRowHeight(FILA_CABECERAS, 58);
  hoja.setFrozenRows(FILA_CABECERAS);

  // Anchos de columna
  hoja.setColumnWidth(1,  180);  // AREA
  hoja.setColumnWidth(2,  280);  // DESCRIPCION DE HALLAZGO
  hoja.setColumnWidth(3,  220);  // REGISTRO FOTOGRAFICO
  hoja.setColumnWidth(4,  160);  // RESPONSABLE
  hoja.setColumnWidth(5,  200);  // PARTICIPANTES
  hoja.setColumnWidth(6,  160);  // DEPARTAMENTO
  hoja.setColumnWidth(7,  280);  // COMENTARIO RESPUESTA
  hoja.setColumnWidth(8,  220);  // FOTO LEVANTAMIENTO
  hoja.setColumnWidth(9,  140);  // ESTADO
  hoja.setColumnWidth(10, 140);  // CÓDIGO
  hoja.setColumnWidth(11, 110);  // QR

  // Mostrar cols visibles, ocultar fecha (col 12)
  hoja.showColumns(10, 2);
  hoja.hideColumns(12);
}

// =============================================
// Colores por estado (columna 8)
// =============================================
function colorearEstado(hoja, fila, estado) {
  const colores = {
    'Pendiente':   { bg: '#FEF9C3', font: '#92400E' },
    'Recurrente':  { bg: '#FEE2E2', font: '#B91C1C' },
    'Corregido':   { bg: '#D1FAE5', font: '#065F46' }
  };
  const c = colores[estado];
  if (c) {
    hoja.getRange(fila, 9)
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
// EXPORTAR EXCEL (idéntico a base de datos)
// =============================================
function exportarInformeExcel(opts) {
  const marca = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const nombreCopia = 'Informe_Evidencias_' + marca;
  const archivoOrigen = DriveApp.getFileById(SPREADSHEET_ID);
  const archivoCopia = archivoOrigen.makeCopy(nombreCopia);
  const ssCopia = SpreadsheetApp.openById(archivoCopia.getId());

  if (opts.tipo === 'dia' || opts.tipo === 'mes' || opts.tipo === 'rango') {
    filtrarCopiaParaExportacion(ssCopia, opts);
  }
  prepararImagenesParaExcel(ssCopia);
  SpreadsheetApp.flush();

  const urlExport = 'https://docs.google.com/spreadsheets/d/' + ssCopia.getId() + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(urlExport, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  // Eliminar copia temporal
  archivoCopia.setTrashed(true);

  if (resp.getResponseCode() !== 200) {
    throw new Error('No se pudo exportar Excel. Código: ' + resp.getResponseCode());
  }

  // Devolver contenido en base64 para descarga directa en el navegador
  return {
    nombre: nombreCopia + '.xlsx',
    base64: Utilities.base64Encode(resp.getContent()),
    tipo: opts.tipo
  };
}

function filtrarCopiaParaExportacion(ss, opts) {
  const tz = Session.getScriptTimeZone();
  const hoy = new Date();
  const fechaObjetivo = opts.fecha || Utilities.formatDate(hoy, tz, 'yyyy-MM-dd');
  const mesObjetivo   = opts.mes   || Utilities.formatDate(hoy, tz, 'yyyy-MM');
  const desdeDate = opts.desde ? new Date(opts.desde + 'T00:00:00') : null;
  const hastaDate = opts.hasta ? new Date(opts.hasta + 'T23:59:59') : null;

  ss.getSheets().forEach(hoja => {
    // Nunca incluir hoja de usuarios en el reporte
    if (hoja.getName() === 'Usuarios') { ss.deleteSheet(hoja); return; }

    const ultima = hoja.getLastRow();
    if (ultima < FILA_DATOS) return;

    // Recolectar filas a eliminar (de abajo hacia arriba para no desplazar índices)
    const eliminar = [];
    for (let i = FILA_DATOS; i <= ultima; i++) {
      const raw = hoja.getRange(i, COL_FECHA).getValue();
      if (!raw) { eliminar.push(i); continue; }

      let fecha = raw instanceof Date ? raw : new Date(raw);
      if (isNaN(fecha.getTime())) {
        // Formato es-PE: "27/5/2026, 2:36:38 p.m." → quitar coma, tomar la parte de fecha
        const soloFecha = String(raw).replace(/,/g, '').trim().split(/\s+/)[0]; // "27/5/2026"
        const p = soloFecha.split('/');
        if (p.length === 3) fecha = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
      }
      if (isNaN(fecha.getTime())) { eliminar.push(i); continue; }

      const ymd = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
      const ym  = Utilities.formatDate(fecha, tz, 'yyyy-MM');
      let visible;
      if      (opts.tipo === 'dia')   visible = ymd === fechaObjetivo;
      else if (opts.tipo === 'mes')   visible = ym  === mesObjetivo;
      else if (opts.tipo === 'rango') visible = desdeDate && hastaDate && fecha >= desdeDate && fecha <= hastaDate;
      else visible = true;
      if (!visible) eliminar.push(i);
    }

    for (let i = eliminar.length - 1; i >= 0; i--) {
      hoja.deleteRow(eliminar[i]);
    }
  });
}

// Convierte =IMAGE("url") a imagen embebida real para que se exporte en xlsx/Office 365
function prepararImagenesParaExcel(ss) {
  const COLS_FOTO = [3, 8]; // REGISTRO FOTOGRAFICO y FOTO DEL LEVANTAMIENTO
  ss.getSheets().forEach(hoja => {
    const ultima = hoja.getLastRow();
    if (ultima < FILA_DATOS) return;
    COLS_FOTO.forEach(col => {
      for (let i = FILA_DATOS; i <= ultima; i++) {
        const cell    = hoja.getRange(i, col);
        const formula = cell.getFormula();
        if (!formula) continue;
        const match = formula.match(/=IMAGE\("([^"]+)"/i);
        if (!match) continue;
        try {
          // CellImage embebida: se exporta como imagen real en xlsx
          const img = SpreadsheetApp.newCellImage()
            .setSourceUrl(match[1])
            .setAltTextTitle('Foto evidencia')
            .build();
          cell.setValue(img);
        } catch (e) {
          // Fallback: si no carga la imagen, dejar la URL como texto clicable
          cell.setValue(match[1]);
        }
      }
    });
  });
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
      hoja.getRange(1, 1, 1, 11).merge()
        .setValue(HOTEL_NOMBRE + ' - WALK THROUGH - ' + mesActual.toUpperCase())
        .setBackground('#7B1827').setFontColor('#FFFFFF')
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
      hoja.getRange(3, 1, 1, 11).setValues([CABECERAS]);
    }

    formatearCabeceras(hoja);

    const ultimaFila = hoja.getLastRow();
    for (let i = FILA_DATOS; i <= ultimaFila; i++) {
      hoja.setRowHeight(i, 120);
      hoja.getRange(i, 1, 1, 9)
        .setVerticalAlignment('middle').setFontSize(10).setWrap(true)
        .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
      if (i % 2 !== 0) hoja.getRange(i, 1, 1, 9).setBackground('#F7F9FC');
      const estado = hoja.getRange(i, 9).getValue();
      if (estado) colorearEstado(hoja, i, estado);
    }

    Logger.log('Reformateada: ' + hoja.getName());
  });
}

// =============================================
// BUSCAR OBSERVACIÓN POR CÓDIGO
// =============================================
function buscarPorCodigo(codigo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  for (const hoja of ss.getSheets()) {
    const ultima = hoja.getLastRow();
    if (ultima < FILA_DATOS) continue;
    for (let i = FILA_DATOS; i <= ultima; i++) {
      const cod = String(hoja.getRange(i, COL_ID_INTERNO).getValue()).trim().toUpperCase();
      if (cod === codigo) {
        const fotoFormula = hoja.getRange(i, 3).getFormula();
        const fotoMatch   = fotoFormula.match(/=IMAGE\("([^"]+)"/i);
        const levFormula  = hoja.getRange(i, 8).getFormula();
        const levMatch    = levFormula  ? levFormula.match(/=IMAGE\("([^"]+)"/i) : null;
        return {
          codigo:      cod,
          hoja:        hoja.getName(),
          fila:        i,
          area:        hoja.getRange(i, 1).getValue(),
          descripcion: hoja.getRange(i, 2).getValue(),
          urlFoto:     fotoMatch ? fotoMatch[1] : '',
          responsable: hoja.getRange(i, 4).getValue(),
          estado:      hoja.getRange(i, 9).getValue(),
          comentario:  hoja.getRange(i, 7).getValue(),
          urlFotoLev:  levMatch  ? levMatch[1]  : ''
        };
      }
    }
  }
  return null;
}

// =============================================
// LEVANTAR OBSERVACIÓN (actualiza fila existente)
// =============================================
function levantarObservacion(codigo, comentario, fotoBase64, estado) {
  const encontrado = buscarPorCodigo(codigo);
  if (!encontrado) return { resultado: 'no_encontrado' };

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(encontrado.hoja);
  const fila = encontrado.fila;

  // Comentario (col 7)
  if (comentario) hoja.getRange(fila, 7).setValue(comentario);

  // Foto levantamiento (col 8)
  if (fotoBase64) {
    const fotoInfo = guardarFotoEnDrive(fotoBase64);
    hoja.getRange(fila, 8).setFormula('=IMAGE("' + fotoInfo.url + '",4,130,200)');
  }

  // Estado (col 9)
  const nuevoEstado = estado || 'Corregido';
  hoja.getRange(fila, 9).setValue(nuevoEstado);
  colorearEstado(hoja, fila, nuevoEstado);

  return { resultado: 'ok', codigo, hoja: encontrado.hoja, fila };
}

// =============================================
// REGISTRO DE USUARIOS
// =============================================
function registrarUsuarioEnSheets(nombre, seccion, pin, email) {
  if (!nombre || !seccion) return { resultado: 'error', error: 'Nombre y sección requeridos' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja    = ss.getSheetByName('Usuarios');

  if (!hoja) {
    hoja = ss.insertSheet('Usuarios');
    hoja.setTabColor('#1A7F54');

    hoja.getRange(1, 1, 1, 6).merge()
        .setValue('USUARIOS REGISTRADOS — Walk Through · Swissotel Lima Peru')
        .setBackground('#1A7F54')
        .setFontColor('#FFFFFF')
        .setFontFamily('Montserrat')
        .setFontWeight('bold')
        .setFontSize(13)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    hoja.setRowHeight(1, 36);

    const cabs = ['NOMBRE', 'CORREO ELECTRÓNICO', 'SECCIÓN', 'PIN', 'FECHA REGISTRO', 'ÚLTIMO ACCESO'];
    hoja.getRange(2, 1, 1, 6).setValues([cabs])
        .setBackground('#145A32')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center');
    hoja.setRowHeight(2, 30);

    hoja.setColumnWidth(1, 200);
    hoja.setColumnWidth(2, 220);
    hoja.setColumnWidth(3, 130);
    hoja.setColumnWidth(4, 70);
    hoja.setColumnWidth(5, 170);
    hoja.setColumnWidth(6, 170);
    hoja.setFrozenRows(2);
  }

  const ahora  = new Date().toLocaleString('es-PE');
  const ultimo = hoja.getLastRow();

  if (ultimo >= 3) {
    const datos = hoja.getRange(3, 1, ultimo - 2, 6).getValues();
    for (let i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toLowerCase() === nombre.trim().toLowerCase()) {
        const fila = i + 3;
        if (email)  hoja.getRange(fila, 2).setValue(email);
        hoja.getRange(fila, 3).setValue(seccion);
        if (pin)    hoja.getRange(fila, 4).setValue(pin);
        hoja.getRange(fila, 6).setValue(ahora);
        return { resultado: 'ok', accion: 'actualizado' };
      }
    }
  }

  const nuevaFila = hoja.getLastRow() + 1;
  hoja.appendRow([nombre, email || '', seccion, pin || '', ahora, ahora]);
  const color = (nuevaFila % 2 === 0) ? '#EAF7EF' : '#FFFFFF';
  hoja.getRange(nuevaFila, 1, 1, 6).setBackground(color).setFontSize(11);

  return { resultado: 'ok', accion: 'registrado' };
}

function recuperarPerfilDesdeSheets(email, pin) {
  if (!email || !pin) return { resultado: 'error', error: 'Correo y PIN requeridos' };

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Usuarios');
  if (!hoja || hoja.getLastRow() < 3)
    return { resultado: 'error', error: 'No se encontró el perfil. Crea uno nuevo.' };

  const datos = hoja.getRange(3, 1, hoja.getLastRow() - 2, 6).getValues();
  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];
    // cols: 0=nombre, 1=email, 2=seccion, 3=pin, 4=fecha reg, 5=último acceso
    if (String(fila[1]).trim().toLowerCase() === email.trim().toLowerCase()) {
      if (String(fila[3]).trim() !== String(pin).trim())
        return { resultado: 'error', error: 'PIN incorrecto' };
      hoja.getRange(i + 3, 6).setValue(new Date().toLocaleString('es-PE'));
      return { resultado: 'ok', nombre: fila[0], email: fila[1], seccion: fila[2], pin: fila[3] };
    }
  }
  return { resultado: 'error', error: 'Correo no encontrado. Verifica o crea un perfil nuevo.' };
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
