// ============================================================
//  PATAGONIA COMIDAS — POS Google Apps Script v3
//  1. Borrá todo el código del proyecto actual
//  2. Pegá este código
//  3. Implementar → Administrar implementaciones → editar → Nueva versión → Implementar
// ============================================================

const SHEET_ID = ''; // ← Pegá el ID de tu Google Sheet entre las comillas

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SHEET_ID
      ? SpreadsheetApp.openById(SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    if (body.tipo === 'venta')      registrarVenta(ss, body.data);
    if (body.tipo === 'movimiento') registrarMovimiento(ss, body.data);
    if (body.tipo === 'precios')    actualizarProductos(ss, body.data);
    if (body.tipo === 'reset')      resetPlanilla(ss);

    return ok();
  } catch (err) {
    return error(err.message);
  }
}

function doGet(e) {
  return ok('POS Patagonia v3 activo');
}

// ─────────────────────────────────────────────
//  VENTAS
//  Columnas: Fecha | Hora | Vendedor | Total ($) | Cant. items | Productos
// ─────────────────────────────────────────────
function registrarVenta(ss, data) {
  const sv = getOrCreate(ss, 'Ventas');
  if (sv.getLastRow() === 0) {
    sv.appendRow(['Fecha', 'Hora', 'Vendedor', 'Total ($)', 'Cant. items', 'Productos']);
    estilo(sv, 1, 6, '#1D9E75');
    sv.setFrozenRows(1);
    sv.setColumnWidth(1, 100);
    sv.setColumnWidth(2, 90);
    sv.setColumnWidth(3, 110);
    sv.setColumnWidth(4, 100);
    sv.setColumnWidth(5, 90);
    sv.setColumnWidth(6, 320);
  }

  const items     = data.items || [];
  const cantItems = items.reduce(function(a, it) { return a + it.qty; }, 0);
  const resumen   = items.map(function(it) { return it.name + ' x' + it.qty; }).join(', ');
  const vendedor  = data.vendor || 'Sin vendedor';

  sv.appendRow([data.date, data.time, vendedor, data.total, cantItems, resumen]);
  sv.getRange(sv.getLastRow(), 4).setNumberFormat('$#,##0');

  // Detalle
  registrarDetalle(ss, data, items, vendedor);

  // Resumen diario
  actualizarResumenDiario(ss, data.date, data.total, cantItems, vendedor);
}

// ─────────────────────────────────────────────
//  DETALLE VENTAS
//  Columnas: Fecha | Hora | Vendedor | Producto | Cantidad | Precio unit. ($) | Subtotal ($)
// ─────────────────────────────────────────────
function registrarDetalle(ss, data, items, vendedor) {
  const sd = getOrCreate(ss, 'Detalle ventas');
  if (sd.getLastRow() === 0) {
    sd.appendRow(['Fecha', 'Hora', 'Vendedor', 'Producto', 'Cantidad', 'Precio unit. ($)', 'Subtotal ($)']);
    estilo(sd, 1, 7, '#1D9E75');
    sd.setFrozenRows(1);
  }
  items.forEach(function(it) {
    sd.appendRow([data.date, data.time, vendedor, it.name, it.qty, it.price, it.sub]);
    var r = sd.getLastRow();
    sd.getRange(r, 6).setNumberFormat('$#,##0');
    sd.getRange(r, 7).setNumberFormat('$#,##0');
  });
}

// ─────────────────────────────────────────────
//  RESUMEN DIARIO
//  Columnas: Fecha | Vendedor | Total ventas ($) | Transacciones | Items vendidos | Ticket promedio ($)
// ─────────────────────────────────────────────
function actualizarResumenDiario(ss, fecha, total, cantItems, vendedor) {
  const sr = getOrCreate(ss, 'Resumen diario');
  if (sr.getLastRow() === 0) {
    sr.appendRow(['Fecha', 'Vendedor', 'Total ventas ($)', 'Transacciones', 'Items vendidos', 'Ticket promedio ($)']);
    estilo(sr, 1, 6, '#534AB7');
    sr.setFrozenRows(1);
    sr.setColumnWidth(1, 100);
    sr.setColumnWidth(2, 110);
    sr.setColumnWidth(3, 130);
    sr.setColumnWidth(4, 110);
    sr.setColumnWidth(5, 110);
    sr.setColumnWidth(6, 130);
  }

  var datos = sr.getDataRange().getValues();
  var fila  = -1;

  // Buscar fila existente con misma fecha + vendedor
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(fecha) && datos[i][1] === vendedor) {
      fila = i + 1;
      break;
    }
  }

  if (fila === -1) {
    // Nueva fila
    sr.appendRow([fecha, vendedor, total, 1, cantItems, total]);
    var nr = sr.getLastRow();
    sr.getRange(nr, 3).setNumberFormat('$#,##0');
    sr.getRange(nr, 6).setNumberFormat('$#,##0');
  } else {
    // Actualizar fila existente
    var totalAcum  = datos[fila - 1][2] + total;
    var txnAcum    = datos[fila - 1][3] + 1;
    var itemsAcum  = datos[fila - 1][4] + cantItems;
    var ticketProm = Math.round(totalAcum / txnAcum);
    sr.getRange(fila, 3, 1, 4).setValues([[totalAcum, txnAcum, itemsAcum, ticketProm]]);
    sr.getRange(fila, 3).setNumberFormat('$#,##0');
    sr.getRange(fila, 6).setNumberFormat('$#,##0');
  }
}

// ─────────────────────────────────────────────
//  MOVIMIENTOS DE STOCK
//  Columnas: Fecha | Hora | Vendedor | Producto | Tipo | Cantidad | Nota
// ─────────────────────────────────────────────
function registrarMovimiento(ss, data) {
  const sm = getOrCreate(ss, 'Movimientos stock');
  if (sm.getLastRow() === 0) {
    sm.appendRow(['Fecha', 'Hora', 'Vendedor', 'Producto', 'Tipo', 'Cantidad', 'Nota']);
    estilo(sm, 1, 7, '#185FA5');
    sm.setFrozenRows(1);
  }

  var tipo = data.tipo === 'ingreso' ? 'Ingreso'
           : data.tipo === 'merma'   ? 'Merma'
           : 'Consumo interno';
  var color = data.tipo === 'ingreso' ? '#EAF3DE'
            : data.tipo === 'merma'   ? '#FCEBEB'
            : '#FAEEDA';

  sm.appendRow([data.date, data.time, data.vendor || '', data.prod, tipo, data.cant, data.nota || '']);
  sm.getRange(sm.getLastRow(), 1, 1, 7).setBackground(color);
}

// ─────────────────────────────────────────────
//  PRODUCTOS — sincroniza precios y stock desde la app
//  Columnas: ID | Categoría | Producto | Precio ($) | Stock actual | Stock mínimo | Activo
// ─────────────────────────────────────────────
function actualizarProductos(ss, data) {
  var sp = getOrCreate(ss, 'Productos');
  sp.clearContents();
  sp.clearFormats();

  sp.appendRow(['ID', 'Categoría', 'Producto', 'Precio ($)', 'Stock actual', 'Stock mínimo', 'Activo']);
  estilo(sp, 1, 7, '#1D9E75');
  sp.setFrozenRows(1);
  sp.setColumnWidth(1, 45);
  sp.setColumnWidth(2, 110);
  sp.setColumnWidth(3, 170);
  sp.setColumnWidth(4, 100);
  sp.setColumnWidth(5, 100);
  sp.setColumnWidth(6, 100);
  sp.setColumnWidth(7, 70);

  var catColors = {
    'Platos':     '#E1F5EE',
    'Pizzas':     '#E6F1FB',
    'Snacks':     '#FAEEDA',
    'Dulces':     '#FBEAF0',
    'Sándwiches': '#EEEDFE',
    'Ensaladas':  '#EAF3DE',
    'Bebidas':    '#FAECE7',
    'Café':       '#F1EFE8'
  };

  var prods = data.productos || [];
  prods.forEach(function(p, i) {
    var row = i + 2;
    sp.appendRow([i + 1, p.cat, p.name, p.price || 0, p.stock || 0, 5, 'Sí']);
    var color = catColors[p.cat] || '#FFFFFF';
    sp.getRange(row, 1, 1, 7).setBackground(color);
    sp.getRange(row, 4).setNumberFormat('$#,##0');
    sp.getRange(row, 1).setHorizontalAlignment('center');
    sp.getRange(row, 4).setHorizontalAlignment('center');
    sp.getRange(row, 5).setHorizontalAlignment('center');
    sp.getRange(row, 6).setHorizontalAlignment('center');
    sp.getRange(row, 7).setHorizontalAlignment('center');
  });

  // Borde general
  sp.getRange(1, 1, prods.length + 1, 7)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

// ─────────────────────────────────────────────
//  RESET PLANILLA — borra hojas de datos
// ─────────────────────────────────────────────
function resetPlanilla(ss) {
  var hojas = ['Ventas', 'Detalle ventas', 'Movimientos stock', 'Resumen diario'];
  hojas.forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    if (h) {
      h.clearContents();
      h.clearFormats();
    }
  });
  // Agregar fila de log
  var log = getOrCreate(ss, 'sistema');
  if (log.getLastRow() === 0) {
    log.appendRow(['Fecha', 'Hora', 'Evento']);
    log.getRange(1,1,1,3).setFontWeight('bold').setBackground('#888888').setFontColor('#ffffff');
  }
  log.appendRow([
    new Date().toLocaleDateString('es-AR'),
    new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'}),
    'Reset total — datos borrados'
  ]);
}

// ─────────────────────────────────────────────
//  UTILIDADES
// ─────────────────────────────────────────────
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function estilo(sheet, fila, cols, color) {
  var r = sheet.getRange(fila, 1, 1, cols);
  r.setFontWeight('bold')
   .setBackground(color)
   .setFontColor('#FFFFFF')
   .setHorizontalAlignment('center');
}

function ok(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: msg || '' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
