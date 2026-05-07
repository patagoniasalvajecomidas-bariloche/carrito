// ======================
// PATAGONIA SALVAJE - SISTEMA COMPLETO v7.8 (fork de Copihue)
// v7.8: Fix getVentasProducto — strip (XXXgr)\/(XXXkg) de nomFila antes de comparar
// v7.7: vender() ahora devuelve ticket en la respuesta — guardarFiado() lo recibe
//       y lo graba en columna C de FIADOS (antes siempre quedaba vacía)
// v7.6: getDetalleTicket — devuelve items reales de hoja Ventas por número de ticket
//       cobrarFiado — fechaPago ahora graba fecha+hora (yyyy-MM-dd HH:mm) para mejor filtrado
// v7.5: upsertCliente — guardarFiado crea/actualiza hoja Clientes automáticamente
// v7.4: listarClientes — lee hoja Clientes (nombre + teléfono) para autocomplete modal fiado
// v7.3: listarFiados incluye pagados últimos 30 días + fechaPago/metodoPago
// v7.2: Fix getVentasProducto — strip (XXXgr) para productos por peso
// v7.1: Fix guardarFiado — eliminado DATEVALUE() en fórmula columna K de FIADOS
//       (DATEVALUE rompe cuando J ya es fecha real, no texto)
// v7.0: Fix categoría en producto nuevo — usa datos.categoria en vez de hardcodear ALMACEN
//       + actualizarStockMinimos() — calcula col N con promedio 7 días × 2
//       + actualizarEstadisticas() — rotación + stock mínimos en un paso
//       + doGet action 'actualizarEstadisticas'
// v6.9: Fix gananciaReal histórica — qty en gramos (ej: 800) se convierte a kg (0.8) para cálculo correcto
// v6.8: Fix costo productos por peso — nomLimpio no sacaba "(800gr)", no matcheaba con inventario
// v6.7: Fix top histórico — ticket "184" vs "0184" no matcheaban
// v6.6: qty por peso guarda número puro (0.8) en vez de texto (0.8kg)
// v6.5: Fix stock por peso — usaba id "peso_TIMESTAMP" en vez de idOriginal, nunca encontraba la fila
// v6.4: Fix qty esPeso — guardaba gramos (800) en vez de kg (0.8), inflaba top productos
// v6.3: Carrito temporal — guardarCarritoTemp + getCarritoTemp (backup en planilla)
// v6.2: Jueves Cervecero — 1 solo gancho (mejor margen) al 20%, resto 10-15%
// v6.1: Jueves Cervecero — estrategia gancho top 2
// v5.9: setConfig + logInterruptor + rebaja stock kg productos por peso
// v5.8: Nueva hoja Ajuste_Rapido separada de Historial
// ======================

const SS_ID = '1K5pm1TEIWJf14uVrajtPENHBk7VSkvR-j3xNes8aWLQ'; // ← Planilla Patagonia Salvaje
const HOJA_INVENTARIO = 'inventario';
const HOJA_CONFIG     = 'config_sistema';
const HOJA_HISTORIAL = 'Historial';
const HOJA_AJUSTE_RAPIDO = 'Ajuste_Rapido';
const HOJA_PROVEEDORES = 'Proveedores';
const MP_ACCESS_TOKEN = 'TU_ACCESS_TOKEN_MP_AQUI'; // ← Reemplazá con tu token de Mercado Pago

// ========== MENÚ UNIFICADO ==========
// ⚠️ UN SOLO onOpen en todo el proyecto. motorInventario.gs NO tiene onOpen propio.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛒 PATAGONIA SALVAJE')
    .addItem('📦 Ingresar Productos', 'abrirFormularioIngreso')
    .addSeparator()
    .addItem('📊 Ver Historial', 'irAHistorial')
    .addItem('🔄 Actualizar Todo', 'actualizarTodo')
    .addSeparator()
    .addItem('▶ Actualizar motor inventario', 'actualizarMotorInventario')
    .addItem('📋 Generar lista de compra', 'actualizarListaCompra')
    .addToUi();
  console.log('✅ Menú Patagonia Salvaje creado');
}

// ========== ABRIR FORMULARIO DE INGRESO ==========
function abrirFormularioIngreso() {
  const html = HtmlService.createHtmlOutputFromFile('ingreso')
    .setWidth(520)
    .setHeight(750)
    .setTitle('📦 Ingreso de Mercadería');
  SpreadsheetApp.getUi().showModalDialog(html, '📦 Ingreso de Mercadería - Patagonia Salvaje');
}

// ========== IR A HOJA HISTORIAL ==========
function irAHistorial() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetHistorial = ss.getSheetByName(HOJA_HISTORIAL);
  if (sheetHistorial) {
    ss.setActiveSheet(sheetHistorial);
  } else {
    SpreadsheetApp.getUi().alert('❌ No se encontró la hoja "Historial"');
  }
}

// ========== ACTUALIZAR TODO ==========
// ⚠️ Las funciones del motor muestran su propio alert — no agregar un tercero acá
function actualizarTodo() {
  actualizarMotorInventario(); // motor de inventario
  actualizarListaCompra();    // genera lista de compra
  actualizarListaCompra();
}

// ========== OBTENER DATOS INICIALES (para formulario ingreso) ==========
function obtenerDatosIniciales() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const datos = sheetInventario.getDataRange().getValues();

    const productos = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        productos.push(String(datos[i][0]).trim());
      }
    }

    const proveedoresSet = new Set();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][4]) {
        proveedoresSet.add(String(datos[i][4]).trim());
      }
    }

    return {
      productos: productos,
      proveedores: Array.from(proveedoresSet).sort()
    };
  } catch (error) {
    console.error('Error obtenerDatosIniciales:', error);
    return { productos: [], proveedores: [] };
  }
}

// ========== OBTENER INFO DE UN PRODUCTO ==========
function obtenerInfoProducto(nombreProducto) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INVENTARIO);
    const datos = sheet.getDataRange().getValues();
    const nombreBuscado = nombreProducto.trim().toUpperCase().replace(/\s+/g, ' ');

    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      const nombreEnSheet = String(datos[i][0]).trim().toUpperCase().replace(/\s+/g, ' ');
      if (nombreEnSheet === nombreBuscado) {
        return {
          encontrado: true,
          fila: i + 1,
          precioVenta: datos[i][1] || 0,
          stockActual: datos[i][5] || 0,
          proveedor: datos[i][4] || '',
          categoria: datos[i][2] || '',
          precioCosto: obtenerUltimoCosto(nombreProducto)
        };
      }
    }

    return { encontrado: false, mensaje: 'Producto no existe en inventario' };
  } catch (error) {
    console.error('Error obtenerInfoProducto:', error);
    return { encontrado: false, error: error.toString() };
  }
}

// ========== OBTENER ÚLTIMO PRECIO DE COSTO DESDE HISTORIAL ==========
function obtenerUltimoCosto(nombreProducto) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetHistorial = ss.getSheetByName(HOJA_HISTORIAL);
    if (!sheetHistorial) return null;

    const datos = sheetHistorial.getDataRange().getValues();
    const nombreBuscado = nombreProducto.trim().toUpperCase();

    for (let i = datos.length - 1; i >= 1; i--) {
      const nombreEnHistorial = String(datos[i][1] || '').trim().toUpperCase();
      if (nombreEnHistorial === nombreBuscado) {
        const costo = datos[i][6];
        if (costo && !isNaN(costo)) return Number(costo);
      }
    }
    return null;
  } catch (error) {
    console.error('Error obtenerUltimoCosto:', error);
    return null;
  }
}

// ========== LEER PROVEEDORES ==========
function leerProveedores() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sheet = ss.getSheetByName(HOJA_PROVEEDORES);

    if (!sheet) {
      sheet = ss.insertSheet(HOJA_PROVEEDORES);
      sheet.getRange(1, 1, 1, 5).setValues([['PROVEEDOR', 'TELEFONO_WA', 'MENSAJE', 'ACTIVO', 'NOTAS']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);

      const invSheet = ss.getSheetByName(HOJA_INVENTARIO);
      const invDatos = invSheet.getDataRange().getValues();
      const provSet = new Set();
      for (let i = 1; i < invDatos.length; i++) {
        if (invDatos[i][4]) provSet.add(String(invDatos[i][4]).trim().toUpperCase());
      }
      let fila = 2;
      provSet.forEach(function(nombre) {
        sheet.getRange(fila, 1).setValue(nombre);
        sheet.getRange(fila, 3).setValue('Hola {nombre}! 🛒 Te hago el pedido:\n\n{productos}\n\n¡Gracias!');
        sheet.getRange(fila, 4).setValue('NO');
        fila++;
      });
      console.log('✅ Hoja Proveedores creada con ' + provSet.size + ' proveedores');
    }

    const datos = sheet.getDataRange().getValues();
    const proveedores = [];
    for (let i = 1; i < datos.length; i++) {
      const nombre = String(datos[i][0] || '').trim();
      if (!nombre) continue;
      proveedores.push({
        nombre:   nombre,
        telefono: String(datos[i][1] || '').replace(/\D/g, ''),
        mensaje:  String(datos[i][2] || '').trim(),
        activo:   String(datos[i][3] || '').toUpperCase() === 'SI',
        notas:    String(datos[i][4] || '').trim()
      });
    }

    return { success: true, proveedores: proveedores };
  } catch(e) {
    console.error('Error leerProveedores:', e);
    return { success: false, mensaje: e.toString(), proveedores: [] };
  }
}

// ========== GUARDAR PROVEEDOR ==========
function guardarProveedor(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
    if (!sheet) return { success: false, mensaje: 'Hoja Proveedores no existe' };

    const rows = sheet.getDataRange().getValues();
    const nombreBuscado = String(datos.nombre || '').trim().toUpperCase();
    let filaEncontrada = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toUpperCase() === nombreBuscado) {
        filaEncontrada = i + 1;
        break;
      }
    }

    const fila = [
      datos.nombre.trim().toUpperCase(),
      String(datos.telefono || '').replace(/\D/g, ''),
      datos.mensaje || '',
      datos.activo ? 'SI' : 'NO',
      datos.notas || ''
    ];

    if (filaEncontrada > 0) {
      sheet.getRange(filaEncontrada, 1, 1, 5).setValues([fila]);
    } else {
      sheet.appendRow(fila);
    }

    return { success: true, mensaje: '✅ Proveedor guardado: ' + datos.nombre };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ========== INGRESAR MERCADERÍA (desde formulario) ==========
function ingresarMercaderia(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const sheetHistorial = ss.getSheetByName(HOJA_HISTORIAL);

    const producto = datos.producto.trim();
    const cantidad = Number(datos.cantidad);
    const precioCosto = datos.precioCosto;
    const precioVenta = datos.precioVenta;
    const proveedor = datos.proveedor;
    const fechaVencimiento = datos.fechaVencimiento;

    if (!producto) return { success: false, mensaje: 'Nombre de producto vacío' };
    const esAjusteDirecto = datos.stockDirecto !== undefined && datos.stockDirecto !== null;
    if (!esAjusteDirecto && cantidad <= 0) return { success: false, mensaje: 'Cantidad debe ser mayor a 0' };

    const datosInventario = sheetInventario.getDataRange().getValues();
    let filaProducto = -1;
    let stockActual = 0;
    const productoNormalizado = producto.trim().toUpperCase().replace(/\s+/g, ' ');

    for (let i = 1; i < datosInventario.length; i++) {
      if (!datosInventario[i][0]) continue;
      const nombreEnSheet = String(datosInventario[i][0]).trim().toUpperCase().replace(/\s+/g, ' ');
      if (nombreEnSheet === productoNormalizado) {
        filaProducto = i + 1;
        stockActual = Number(datosInventario[i][5]) || 0;
        break;
      }
    }

    // ── PRODUCTO NUEVO ──
    if (filaProducto === -1) {
      const ultimaFila = sheetInventario.getLastRow() + 1;
      const nuevoStock = cantidad;
      const todosLosDatos = sheetInventario.getDataRange().getValues();
      let maxId = 0;
      for (let i = 1; i < todosLosDatos.length; i++) {
        const idRaw = String(todosLosDatos[i][3] || '').replace('ID', '');
        const idNum = parseInt(idRaw);
        if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
      }
      const nuevoId = 'ID' + (maxId + 1);

      const categoriaFinal = (datos.categoria && datos.categoria.trim())
        ? datos.categoria.trim().toUpperCase()
        : 'ALMACEN';

      sheetInventario.getRange(ultimaFila, 1).setValue(producto.toUpperCase());
      sheetInventario.getRange(ultimaFila, 2).setValue(precioVenta || 0);
      sheetInventario.getRange(ultimaFila, 3).setValue(categoriaFinal);
      sheetInventario.getRange(ultimaFila, 4).setValue(nuevoId);
      if (proveedor) sheetInventario.getRange(ultimaFila, 5).setValue(proveedor);
      sheetInventario.getRange(ultimaFila, 6).setValue(nuevoStock);

      // Rellenar columnas automáticas con 0 y fórmulas
      var f = ultimaFila; // alias corto
      sheetInventario.getRange(f, 7).setValue(0);   // G: Relampago
      sheetInventario.getRange(f, 8).setValue(0);   // H: Destacada
      sheetInventario.getRange(f, 9).setValue(0);   // I: Especial
      sheetInventario.getRange(f, 12).setValue(0);  // L: promos/antojo_popup

      // ── J: categoriaOferta — fórmula en INGLÉS ──────────────────
      sheetInventario.getRange(f, 10).setFormula(
        '=IF(A'+f+'="","",IF(F'+f+'<=0,"sin stock",IF(NOT(OR(I'+f+'=0,I'+f+'=1)),"ERROR ESPECIAL",' +
        'IF(NOT(ISNUMBER(MATCH(G'+f+',{0,1,2,3,4,5,6,10,11,12,13,14},0))),"ERROR RELAMPAGO",' +
        'IF(OR(H'+f+'>70,AND(H'+f+'>0,H'+f+'<15)),"ERROR DESTACADA",' +
        'IF((I'+f+'=1)+(H'+f+'>0)+(G'+f+'>0)>1,"ERROR DOBLE OFERTA",' +
        'IF(I'+f+'=1,"especial",' +
        'IF(H'+f+'>0,"destacada",' +
        'IF(G'+f+'>0,' +
          'IF(OR(AND(G'+f+'<=6,F'+f+'<(G'+f+'+1)),AND(G'+f+'=10,F'+f+'<2),AND(G'+f+'>=11,F'+f+'<1)),' +
            '"REPONER "&IF(G'+f+'<=6,G'+f+'+1-F'+f+',IF(G'+f+'=10,2-F'+f+',1-F'+f+'))&" PARA PROMO",' +
            'IF(G'+f+'<=6,"relampago "&(G'+f+'+1)&"x"&G'+f+',' +
            'IF(G'+f+'=10,"2da50off",' +
            'IF(G'+f+'>=11,"relampago "&CHOOSE(G'+f+'-10,10,15,20,25)&"%","relampago")))),' +
        '"sin oferta"))))))))))'
      );

      // ── K: PT (Precio Técnico) — fórmula en INGLÉS ──────────────
      sheetInventario.getRange(f, 11).setFormula(
        '=IF(AND(C'+f+'="PROMOS",F'+f+'>=1),"COMBO ACTIVADO EN CAT. PROMOS",' +
        'IF(OR(F'+f+'<=0,B'+f+'<=0),"SIN STOCK / CARGAR STOCK",' +
        'IF(B'+f+'="","CARGAR PRECIO",' +
        'IF(AND(ISNUMBER(G'+f+'),G'+f+'>0),' +
          'B'+f+'*VLOOKUP(G'+f+',{1,0.5;2,0.6667;3,0.75;4,0.8;5,0.8333;6,0.8571;10,0.75},2,FALSE),' +
        'IF(AND(ISNUMBER(H'+f+'),H'+f+'>0),' +
          'B'+f+'*(1-H'+f+'/100),' +
        'IF(I'+f+'=1,B'+f+'*0.9,"NORMAL / SIN OFERTA"))))))'
      );

      if (fechaVencimiento) sheetInventario.getRange(f, 16).setValue(fechaVencimiento);

      const fechaNuevo = new Date();
      if (cantidad > 0 || precioCosto || precioVenta) {
        sheetHistorial.appendRow([
          fechaNuevo, producto, cantidad, nuevoStock, nuevoId,
          proveedor || '', precioCosto || '', precioVenta || '', fechaVencimiento || ''
        ]);
      }

      let mensajeNuevo = '✅ Producto NUEVO creado!\n📦 ' + producto + '\n🆔 ' + nuevoId + '\n📊 Stock inicial: ' + nuevoStock;
      if (precioVenta) mensajeNuevo += '\n💰 Precio venta: $' + precioVenta;
      mensajeNuevo += categoriaFinal === 'ALMACEN' && !datos.categoria
        ? '\n\n⚠️ No se seleccionó categoría (se asignó ALMACEN por defecto)'
        : '\n\n📂 Categoría: ' + categoriaFinal;

      return { success: true, mensaje: mensajeNuevo, nuevoStock: nuevoStock, esNuevo: true };
    }

    // ── PRODUCTO EXISTENTE ──
    const nuevoStock = (datos.stockDirecto !== undefined && datos.stockDirecto !== null)
      ? Number(datos.stockDirecto)
      : stockActual + cantidad;
    sheetInventario.getRange(filaProducto, 6).setValue(nuevoStock);

    if (proveedor) sheetInventario.getRange(filaProducto, 5).setValue(proveedor);
    if (precioVenta && precioVenta > 0) sheetInventario.getRange(filaProducto, 2).setValue(precioVenta);
    if (fechaVencimiento) sheetInventario.getRange(filaProducto, 16).setValue(fechaVencimiento);
    if (datos.nombreNuevo && datos.nombreNuevo.trim()) sheetInventario.getRange(filaProducto, 1).setValue(datos.nombreNuevo.trim().toUpperCase());
    if (datos.categoria && datos.categoria.trim()) sheetInventario.getRange(filaProducto, 3).setValue(datos.categoria.trim().toUpperCase());

    const fecha = new Date();
    const tz   = Session.getScriptTimeZone();
    const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    const horaStr  = Utilities.formatDate(fecha, tz, 'HH:mm');

    // ── REGISTRAR EN SALIDAS si es retiro por vencimiento (stockDirecto=0 y cantidad=0) ──
    const esRetiroVencido = esAjusteDirecto && Number(datos.stockDirecto) === 0 && cantidad === 0 && stockActual > 0;
    if (esRetiroVencido) {
      var shSal = ss.getSheetByName('SALIDAS');
      if (!shSal) {
        shSal = ss.insertSheet('SALIDAS');
        shSal.getRange(1, 1, 1, 10).setValues([['Fecha','Hora','Producto','ID','Cantidad','Costo Unit.','Precio Venta','Motivo','Observación','Vendedor']]);
        shSal.getRange(1, 1, 1, 10).setFontWeight('bold');
      }
      // Obtener ID del producto
      var idProd = datosInventario[filaProducto - 1][3] || '';
      var costoProd = datosInventario[filaProducto - 1][18] || 0; // col S = costo
      var precioProd = datosInventario[filaProducto - 1][1] || 0; // col B = precio venta
      shSal.appendRow([
        fechaStr, horaStr,
        producto, idProd,
        stockActual, costoProd, precioProd,
        'VENCIDO', 'Auto-retiro por vencimiento', datos.vendedor || ''
      ]);
    }

    // Solo escribir en Historial si es un ingreso real (cantidad > 0) o si trae precio/costo
    const esIngresoReal = cantidad > 0 || precioCosto || precioVenta;
    if (esIngresoReal) {
      sheetHistorial.appendRow([
        fecha, producto, cantidad, nuevoStock, '',
        proveedor || '', precioCosto || '', precioVenta || '', fechaVencimiento || ''
      ]);
    }

    let mensaje = '✅ Ingreso exitoso!\n📦 ' + producto + '\n➕ Cantidad: ' + cantidad + '\n📊 Stock: ' + stockActual + ' → ' + nuevoStock;
    if (precioVenta) mensaje += '\n💰 Precio venta actualizado: $' + precioVenta;

    return { success: true, mensaje: mensaje, nuevoStock: nuevoStock };
  } catch (error) {
    console.error('Error ingresarMercaderia:', error);
    return { success: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ========== AJUSTE RÁPIDO ==========
function ajustarProducto(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetInventario = ss.getSheetByName(HOJA_INVENTARIO);

    let sheetAjuste = ss.getSheetByName(HOJA_AJUSTE_RAPIDO);
    if (!sheetAjuste) {
      sheetAjuste = ss.insertSheet(HOJA_AJUSTE_RAPIDO);
      sheetAjuste.appendRow([
        'FECHA', 'PRODUCTO', 'NOMBRE ANTES', 'NOMBRE DESPUÉS',
        'STOCK ANTES', 'STOCK DESPUÉS', 'PRECIO ANTES', 'PRECIO DESPUÉS',
        'CATEGORÍA ANTES', 'CATEGORÍA DESPUÉS', 'USUARIO'
      ]);
      sheetAjuste.getRange(1, 1, 1, 11).setBackground('#37474f').setFontColor('white').setFontWeight('bold');
    }

    const producto = String(datos.producto || '').trim();
    if (!producto) return { success: false, mensaje: 'Nombre de producto vacío' };

    const datosInv = sheetInventario.getDataRange().getValues();
    let filaProducto = -1;
    let stockActual = 0, precioActual = 0, catActual = '', nombreActual = '';
    const productoNorm = producto.toUpperCase().replace(/\s+/g, ' ');

    for (let i = 1; i < datosInv.length; i++) {
      if (!datosInv[i][0]) continue;
      const nombreEnSheet = String(datosInv[i][0]).trim().toUpperCase().replace(/\s+/g, ' ');
      if (nombreEnSheet === productoNorm) {
        filaProducto = i + 1;
        nombreActual = String(datosInv[i][0]).trim();
        precioActual = Number(datosInv[i][1]) || 0;
        catActual    = String(datosInv[i][2] || '').trim();
        stockActual  = Number(datosInv[i][5]) || 0;
        break;
      }
    }

    if (filaProducto === -1) return { success: false, mensaje: 'Producto no encontrado: ' + producto };

    const stockNuevo  = (datos.stockDespues !== undefined && datos.stockDespues !== null) ? Number(datos.stockDespues) : stockActual;
    const precioNuevo = datos.precioDespues ? Number(datos.precioDespues) : precioActual;
    const nombreNuevo = datos.nombreNuevo   ? String(datos.nombreNuevo).trim().toUpperCase() : nombreActual;
    const catNueva    = datos.categoriaDespues ? String(datos.categoriaDespues).trim().toUpperCase() : catActual;

    const stockRegistroAntes  = (datos.stockAntes !== undefined && datos.stockAntes !== null) ? Number(datos.stockAntes) : stockActual;
    const precioRegistroAntes = datos.precioAntes ? Number(datos.precioAntes) : precioActual;

    sheetInventario.getRange(filaProducto, 6).setValue(stockNuevo);
    if (precioNuevo !== precioActual)   sheetInventario.getRange(filaProducto, 2).setValue(precioNuevo);
    if (nombreNuevo !== nombreActual)   sheetInventario.getRange(filaProducto, 1).setValue(nombreNuevo);
    if (catNueva    !== catActual)      sheetInventario.getRange(filaProducto, 3).setValue(catNueva);

    const fechaAj = new Date();
    sheetAjuste.appendRow([
      fechaAj, producto,
      nombreActual,          nombreNuevo  !== nombreActual  ? nombreNuevo  : '',
      stockRegistroAntes,    stockNuevo,
      precioRegistroAntes,   precioNuevo  !== precioRegistroAntes ? precioNuevo : '',
      catActual,             catNueva     !== catActual     ? catNueva     : '',
      'POS'
    ]);

    let cambios = [];
    if (stockNuevo !== stockRegistroAntes)   cambios.push('Stock: ' + stockRegistroAntes + ' → ' + stockNuevo);
    if (precioNuevo !== precioRegistroAntes)  cambios.push('Precio: $' + precioRegistroAntes + ' → $' + precioNuevo);
    if (nombreNuevo !== nombreActual)         cambios.push('Nombre: ' + nombreActual + ' → ' + nombreNuevo);
    if (catNueva    !== catActual)            cambios.push('Categoría: ' + catActual + ' → ' + catNueva);

    return { success: true, mensaje: '✅ Ajuste guardado\n' + cambios.join('\n'), stockNuevo: stockNuevo };
  } catch (error) {
    console.error('Error ajustarProducto:', error);
    return { success: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ========== ROTACIÓN AUTOMÁTICA ==========
function calcularRotacionAuto() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);

    const sheetVentas = ss.getSheetByName('Ventas');
    const ventasMap = {};
    if (sheetVentas) {
      const filas = sheetVentas.getDataRange().getValues();
      const hace30 = new Date();
      hace30.setDate(hace30.getDate() - 30);
      for (let i = 1; i < filas.length; i++) {
        const f = filas[i];
        if (!f[0] || !f[3]) continue;
        const nombre = String(f[3]).trim().toUpperCase();
        if (nombre.includes('TOTAL TICKET') || nombre.startsWith('─')) continue;
        const fecha = f[1] instanceof Date ? f[1] : new Date(f[1]);
        if (isNaN(fecha.getTime()) || fecha < hace30) continue;
        const qty = parseFloat(String(f[4]).replace('gr','').replace('kg','')) || 0;
        ventasMap[nombre] = (ventasMap[nombre] || 0) + qty;
      }
    }

    const sheetHist = ss.getSheetByName(HOJA_HISTORIAL);
    const ingresosMap = {};
    if (sheetHist) {
      const filas = sheetHist.getDataRange().getValues();
      const hace90 = new Date();
      hace90.setDate(hace90.getDate() - 90);
      for (let i = 1; i < filas.length; i++) {
        const f = filas[i];
        if (!f[1]) continue;
        const nombre = String(f[1]).trim().toUpperCase();
        const fecha = f[0] instanceof Date ? f[0] : new Date(f[0]);
        if (isNaN(fecha.getTime()) || fecha < hace90) continue;
        ingresosMap[nombre] = (ingresosMap[nombre] || 0) + 1;
      }
    }

    const rotacionMap = {};
    const todosLosNombres = new Set([...Object.keys(ventasMap), ...Object.keys(ingresosMap)]);

    todosLosNombres.forEach(function(nombre) {
      const u = ventasMap[nombre] || 0;
      const v = ingresosMap[nombre] || 0;

      let bVentas = null;
      if      (u >= 8) bVentas = 5;
      else if (u >= 5) bVentas = 4;
      else if (u >= 3) bVentas = 3;
      else if (u >= 1) bVentas = 2;

      let bRepo = 1;
      if      (v >= 12) bRepo = 5;
      else if (v >= 6)  bRepo = 4;
      else if (v >= 3)  bRepo = 3;
      else if (v >= 1)  bRepo = 2;

      const combinado = bVentas !== null
        ? Math.round(bVentas * 0.7 + bRepo * 0.3)
        : bRepo;

      rotacionMap[nombre] = Math.max(1, Math.min(5, combinado));
    });

    if (calcularRotacionAuto._escribir) {
      const sheetInv = ss.getSheetByName(HOJA_INVENTARIO);
      const invDatos = sheetInv.getDataRange().getValues();
      const COL_ROTACION = 15;
      const actualizaciones = [];
      for (let i = 1; i < invDatos.length; i++) {
        const nombre = String(invDatos[i][0] || '').trim().toUpperCase();
        if (!nombre) continue;
        const rot = rotacionMap[nombre] || (parseInt(invDatos[i][14]) || 1);
        actualizaciones.push({ fila: i + 1, valor: rot });
      }
      actualizaciones.forEach(function(a) {
        sheetInv.getRange(a.fila, COL_ROTACION).setValue(a.valor);
      });
      console.log('✅ Rotación escrita en columna O: ' + actualizaciones.length + ' productos');
    }

    return rotacionMap;

  } catch(e) {
    console.warn('⚠️ calcularRotacionAuto falló, usando columna O:', e.toString());
    return {};
  }
}

function actualizarRotacionPlanilla() {
  calcularRotacionAuto._escribir = true;
  const mapa = calcularRotacionAuto();
  calcularRotacionAuto._escribir = false;
  const total = Object.keys(mapa).length;
  console.log('📊 Rotación actualizada para ' + total + ' productos');
  return { success: true, total: total };
}

// ========== ACTUALIZAR STOCK MÍNIMOS (columna N) ==========
function actualizarStockMinimos() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetVentas = ss.getSheetByName('Ventas');
    const sheetInv    = ss.getSheetByName(HOJA_INVENTARIO);
    if (!sheetVentas || !sheetInv) return { success: false, mensaje: 'Faltan hojas' };

    const hoy   = new Date();
    const hace7 = new Date(hoy - 7 * 24 * 60 * 60 * 1000);

    const ventas = sheetVentas.getDataRange().getValues();
    const ventasProd = {};

    for (let i = 1; i < ventas.length; i++) {
      const f = ventas[i];
      const fecha = f[1];
      if (!(fecha instanceof Date) || fecha < hace7) continue;
      const nombre = String(f[3] || '').trim();
      if (!nombre || nombre.includes('TOTAL TICKET') || nombre.startsWith('─')) continue;

      const nomLimpio = nombre
        .replace(/\[PRECIO ESP[^\]]*\]/i, '')
        .replace(/\s*\(\d+gr\)/i, '')
        .replace(/\s*\(\d+\.?\d*kg\)/i, '')
        .trim()
        .toUpperCase();

      const qty = parseFloat(String(f[4] || '0').replace('gr','').replace('kg','')) || 0;
      const esGramos = /\(\d+gr\)/i.test(nombre) && qty >= 50;
      const qtdReal  = esGramos ? qty / 1000 : qty;

      ventasProd[nomLimpio] = (ventasProd[nomLimpio] || 0) + qtdReal;
    }

    const inv = sheetInv.getDataRange().getValues();
    let actualizados = 0;
    let sinDatos = 0;

    for (let i = 1; i < inv.length; i++) {
      const nombre = String(inv[i][0] || '').trim().toUpperCase();
      if (!nombre) continue;

      const totalVendido   = ventasProd[nombre] || 0;
      const promedioDiario = totalVendido / 7;
      const stockMinimo    = promedioDiario > 0
        ? Math.ceil(promedioDiario * 2)
        : 1;

      if (totalVendido === 0) sinDatos++;

      const actualActual = parseInt(inv[i][13]) || 0;
      if (actualActual !== stockMinimo) {
        sheetInv.getRange(i + 1, 14).setValue(stockMinimo);
        actualizados++;
      }
    }

    console.log('✅ Stock mínimos actualizados: ' + actualizados + ' productos');
    return {
      success: true,
      mensaje: '✅ Stock mínimos: ' + actualizados + ' actualizados · ' + sinDatos + ' sin datos',
      actualizados,
      sinDatos
    };

  } catch(e) {
    console.error('Error actualizarStockMinimos:', e);
    return { success: false, mensaje: e.toString() };
  }
}

// ========== ACTUALIZAR ESTADÍSTICAS ==========
function actualizarEstadisticas() {
  const rotacion = actualizarRotacionPlanilla();
  const minimos  = actualizarStockMinimos();
  console.log('📊 Rotación:', rotacion.total, '| Mínimos:', minimos.actualizados);
  return { success: true, rotacion, minimos };
}

// ========== HELPERS MOTOR OFERTAS ==========
function _ofertaDiasVencer_(fila) {
  try {
    const v = fila.length > 15 ? fila[15] : '';
    if (!v) return null;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return null;
    return Math.round((d - new Date()) / (1000 * 60 * 60 * 24));
  } catch(e) { return null; }
}

function _ofertaDiasDesdePromo_(fila) {
  try {
    const v = fila.length > 17 ? fila[17] : '';
    if (!v) return null;
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return null;
    return Math.round((new Date() - d) / (1000 * 60 * 60 * 24));
  } catch(e) { return null; }
}

function _ofertaPuntaje_(fila) {
  var p = 0;
  var dv = _ofertaDiasVencer_(fila);
  var dp = _ofertaDiasDesdePromo_(fila);
  var stock = parseInt(fila[5]) || 0;
  var rotacion = fila.length > 14 ? (parseInt(fila[14]) || 0) : 0;
  var precio = parseInt(fila[1]) || 0;
  if (dp !== null && dp < 15) return -99;
  if (dv !== null) {
    if (dv <= 3) p += 5;
    else if (dv <= 7 && stock >= 5) p += 4;
  }
  if (rotacion >= 3) p += 3;
  if (precio > 1000) p += 2;
  return p;
}

function _ofertaBuildProducto_(fila, i) {
  var tz = Session.getScriptTimeZone();
  var venc = '';
  try {
    var v = fila.length > 15 ? fila[15] : '';
    if (v) {
      venc = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
    }
  } catch(e) {}
  var dpv = _ofertaDiasVencer_(fila);
  var estadoVencimiento = 'ok';
  if (dpv !== null) {
    if (dpv <= 0)  estadoVencimiento = 'vencido';
    else if (dpv <= 3)  estadoVencimiento = 'urgente';
    else if (dpv <= 15) estadoVencimiento = 'proximo';
  }
  return {
    id:                i,
    name:              String(fila[0] || '').trim(),
    price:             parseInt(fila[1]) || 0,
    category:          String(fila[2] || 'ALMACEN').trim().toUpperCase(),
    stock:             parseInt(fila[5]) || 0,
    relampago:         parseInt(fila[6]) || 0,
    destacada:         parseInt(fila[7]) || 0,
    especial:          parseInt(fila[8]) || 0,
    rotacion:          fila.length > 14 ? (parseInt(fila[14]) || 0) : 0,
    vencimiento:       venc,
    diasParaVencer:    dpv,
    estadoVencimiento: estadoVencimiento,
    puntaje:           _ofertaPuntaje_(fila)
  };
}

function calcularOfertas() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(HOJA_INVENTARIO);
  var datos = sheet.getDataRange().getValues();

  var limites = { relampago: 3, especiales: 2, destacadas: 9 };
  var horariosOferta = {
    relampago:  { inicio: 0,  cierre: 24, activo: false },
    destacadas: { inicio: 0,  cierre: 24, activo: false },
    especiales: { inicio: 0,  cierre: 24, activo: false }
  };
  var _inicioRelampago = null, _cierreRelampago = null;
  var _inicioDestacadas = null, _cierreDestacadas = null;
  var _inicioEspeciales = null, _cierreEspeciales = null;

  var diaIdxJC = [1,2,3,4,5,6,0].indexOf(new Date().getDay());
  var tzNow = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
  var horaActual = tzNow.getHours() + tzNow.getMinutes() / 60;

  // Leer todos los límites y horarios desde config_sistema al inicio
  try {
    var shCfgLim = SpreadsheetApp.openById(SS_ID).getSheetByName(HOJA_CONFIG);
    if (shCfgLim) {
      var cfgLim = shCfgLim.getDataRange().getValues();

      // Helper: parsear hora de celda (número o "HH:MM")
      function _parsearHora_(val) {
        if (val === null || val === undefined || val === '') return null;
        // ISO string: Google Sheets manda horas como "1899-12-30T21:00:00.000Z"
        if (typeof val === 'string' && val.includes('T')) {
          try {
            var d = new Date(val);
            return ((d.getUTCHours() - 3 + 24) % 24) + d.getUTCMinutes() / 60;
          } catch(e) { return null; }
        }
        // Fracción de día (0-1)
        if (typeof val === 'number') {
          if (val > 0 && val < 1) return val * 24;
          return val;
        }
        var s = String(val).trim();
        var m = s.match(/^(\d{1,2}):(\d{2})/);
        if (m) return parseInt(m[1]) + parseInt(m[2]) / 60;
        var n = parseFloat(s);
        return isNaN(n) ? null : n;
      }

      // Helper: verificar si el modo A respeta el horario
      function _estaActivoHoy_(diasVal, inicioVal, cierreVal) {
        var d = String(diasVal || '').trim().toUpperCase();
        if (d === '0') return false;           // desactivado forzado
        if (d === '1') return true;            // activado forzado
        // modo A: respetar horario
        var ini = _parsearHora_(inicioVal);
        var cie = _parsearHora_(cierreVal);
        if (ini === null || cie === null) return true;
        if (cie > ini) return horaActual >= ini && horaActual < cie;
        // horario que cruza medianoche
        return horaActual >= ini || horaActual < cie;
      }

      for (var li = 0; li < cfgLim.length; li++) {
        var lk = String(cfgLim[li][0] || '').trim().toLowerCase();
        var lv = cfgLim[li][diaIdxJC + 1];

        // Límites
        if (lk.indexOf('maximo relampago') !== -1 || lk.indexOf('máximo relampago') !== -1) { var n = parseInt(lv); if (!isNaN(n) && n > 0) limites.relampago = n; }
        if (lk.indexOf('maximo especiales') !== -1 || lk.indexOf('máximo especiales') !== -1) { var n = parseInt(lv); if (!isNaN(n) && n > 0) limites.especiales = n; }
        if (lk.indexOf('maximo destacadas') !== -1 || lk.indexOf('máximo destacadas') !== -1) { var n = parseInt(lv); if (!isNaN(n) && n > 0) limites.destacadas = n; }

        // Horarios — claves únicas y descriptivas
        if (lk === 'hora inicio relampago')         _inicioRelampago  = _parsearHora_(lv);
        if (lk === 'hora cierre relampago')         _cierreRelampago  = _parsearHora_(lv);
        if (lk === 'hora inicio destacadas')        _inicioDestacadas = _parsearHora_(lv);
        if (lk === 'hora cierre destacadas')        _cierreDestacadas = _parsearHora_(lv);
        if (lk === 'hora inicio especiales')        _inicioEspeciales = _parsearHora_(lv);
        if (lk === 'hora cierre especiales')        _cierreEspeciales = _parsearHora_(lv);

        // Flag días
        if (lk === 'dias ofertas relampago' || lk === 'dias ofertas relámpago') {
          horariosOferta.relampago.activo = _estaActivoHoy_(lv, _inicioRelampago, _cierreRelampago);
        }
        if (lk === 'dias ofertas destacadas') {
          horariosOferta.destacadas.activo = _estaActivoHoy_(lv, _inicioDestacadas, _cierreDestacadas);
        }
        if (lk === 'dias ofertas especiales') {
          horariosOferta.especiales.activo = _estaActivoHoy_(lv, _inicioEspeciales, _cierreEspeciales);
        }
      }
    }
  } catch(eLim) { Logger.log('Error leyendo config: ' + eLim); }

  var relampagoActivo = [];
  var idsUsados = {};
  var cerveceroActivoHoy = false;
  try {
    var ssCfgJC = SpreadsheetApp.openById(SS_ID);
    var shCfgJC = ssCfgJC.getSheetByName(HOJA_CONFIG);
    if (shCfgJC) {
      var cfgJC = shCfgJC.getDataRange().getValues();
      for (var cji = 0; cji < cfgJC.length; cji++) {
        var cjClave = String(cfgJC[cji][0] || '').trim().toLowerCase();
        if (cjClave.includes('jueves cervecero') || cjClave.includes('cervecero')) {
          var cjVal = cfgJC[cji][diaIdxJC + 1];
          cerveceroActivoHoy = (cjVal === 1 || cjVal === '1' || cjVal === true);
          break;
        }
      }
    }
  } catch(eCJ) {}

  function _esCerveza_(fila) {
    var nom = String(fila[0] || '').trim().toUpperCase();
    var cat = String(fila[2] || '').trim().toUpperCase();
    return cat.includes('CERV') || nom.includes('CERVEZA') || nom.includes(' LATA') ||
           nom.includes('BIRRA') || nom.includes(' IPA') || nom.includes(' STOUT') || nom.includes(' PORTER');
  }

  // Helper: detecta si la columna J dice que hay que reponer stock para la promo
  function _stockSuficienteParaOferta_(fila) {
    var catOferta = String(fila[9] || '').trim().toUpperCase();
    // Si dice REPONER → stock insuficiente para la promo
    if (catOferta.indexOf('REPONER') !== -1) return false;
    // Verificar stock mínimo según tipo de promo
    var stock = parseInt(fila[5]) || 0;
    if (catOferta.indexOf('3X2') !== -1 || catOferta.indexOf('3 X 2') !== -1) return stock >= 3;
    if (catOferta.indexOf('2X1') !== -1 || catOferta.indexOf('2 X 1') !== -1) return stock >= 2;
    if (catOferta.indexOf('4X3') !== -1 || catOferta.indexOf('4 X 3') !== -1) return stock >= 4;
    return stock >= 1;
  }

  var candidatosRelampago = [];
  for (var i3 = 1; i3 < datos.length; i3++) {
    var fila3 = datos[i3];
    if (!fila3[0]) continue;
    if (idsUsados[i3]) continue;
    var stock3 = parseInt(fila3[5]) || 0;
    var relampago3 = parseInt(fila3[6]) || 0;
    if (stock3 <= 0 || relampago3 <= 0) continue;
    if (!_stockSuficienteParaOferta_(fila3)) continue; // columna J: REPONER o stock insuficiente
    var cat3 = String(fila3[2] || '').trim().toUpperCase();
    if (cat3 === 'CERVEZAS') continue;
    if (cerveceroActivoHoy && _esCerveza_(fila3)) continue;
    candidatosRelampago.push({ fila: fila3, i: i3, p: _ofertaPuntaje_(fila3) });
  }

  candidatosRelampago.sort(function(a, b) { return b.p - a.p; });

  var relampagoCandidatosParaRotar = candidatosRelampago;
  if (relampagoCandidatosParaRotar.length > limites.relampago) {
    var tzDateR   = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var dateKeyR  = tzDateR.getFullYear() + '-' +
                    String(tzDateR.getMonth() + 1).padStart(2, '0') + '-' +
                    String(tzDateR.getDate()).padStart(2, '0');
    var argDateR  = new Date(dateKeyR);
    var dayOfYearR = Math.floor((argDateR - new Date(argDateR.getFullYear(), 0, 0)) / 86400000);
    var totalSetsR = Math.ceil(relampagoCandidatosParaRotar.length / limites.relampago);
    var setNumberR = dayOfYearR % totalSetsR;
    var startIdxR  = (setNumberR * limites.relampago) % relampagoCandidatosParaRotar.length;
    var rotadosR   = [];
    for (var rri = 0; rri < limites.relampago; rri++) {
      rotadosR.push(relampagoCandidatosParaRotar[(startIdxR + rri) % relampagoCandidatosParaRotar.length]);
    }
    relampagoCandidatosParaRotar = rotadosR;
  }

  var relampagoPoolCompleto = candidatosRelampago.map(function(c) {
    return _ofertaBuildProducto_(c.fila, c.i);
  });

  relampagoCandidatosParaRotar.forEach(function(c) {
    relampagoActivo.push(_ofertaBuildProducto_(c.fila, c.i));
  });

  var destacadasActivas = [];
  var especialesActivas = [];
  for (var idx = 1; idx < datos.length; idx++) {
    var fd = datos[idx];
    if (!fd[0]) continue;
    var stockD = parseInt(fd[5]) || 0;
    if (stockD <= 0) continue;
    if (!_stockSuficienteParaOferta_(fd)) continue; // columna J: REPONER o stock insuficiente
    if (cerveceroActivoHoy && _esCerveza_(fd)) continue;
    if ((parseInt(fd[7]) || 0) > 0) destacadasActivas.push(_ofertaBuildProducto_(fd, idx));
    if ((parseInt(fd[8]) || 0) > 0) especialesActivas.push(_ofertaBuildProducto_(fd, idx));
  }

  var especialesPool = especialesActivas;
  if (especialesActivas.length > limites.especiales) {
    var tzDate    = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var dateKey   = tzDate.getFullYear() + '-' +
                    String(tzDate.getMonth() + 1).padStart(2, '0') + '-' +
                    String(tzDate.getDate()).padStart(2, '0');
    var argDate   = new Date(dateKey);
    var dayOfYear = Math.floor((argDate - new Date(argDate.getFullYear(), 0, 0)) / 86400000);
    var totalSets = Math.ceil(especialesActivas.length / limites.especiales);
    var setNumber = dayOfYear % totalSets;
    var startIdx  = (setNumber * limites.especiales) % especialesActivas.length;
    var rotados   = [];
    for (var ri = 0; ri < limites.especiales; ri++) {
      rotados.push(especialesActivas[(startIdx + ri) % especialesActivas.length]);
    }
    especialesActivas = rotados;
  }

  var destacadasPool = destacadasActivas.slice();
  var destacadasRotadas = destacadasActivas.slice();
  if (destacadasActivas.length > limites.destacadas) {
    var tzDateD   = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Argentina/Buenos_Aires' }));
    var dateKeyD  = tzDateD.getFullYear() + '-' +
                    String(tzDateD.getMonth() + 1).padStart(2, '0') + '-' +
                    String(tzDateD.getDate()).padStart(2, '0');
    var argDateD  = new Date(dateKeyD);
    var dayOfYearD = Math.floor((argDateD - new Date(argDateD.getFullYear(), 0, 0)) / 86400000);
    var totalSetsD = Math.ceil(destacadasActivas.length / limites.destacadas);
    var setNumberD = dayOfYearD % totalSetsD;
    var startIdxD  = (setNumberD * limites.destacadas) % destacadasActivas.length;
    destacadasRotadas = [];
    for (var rdi = 0; rdi < limites.destacadas; rdi++) {
      destacadasRotadas.push(destacadasActivas[(startIdxD + rdi) % destacadasActivas.length]);
    }
  }

  var poolRecienLlegados = [];

  // Respetar horarios: vaciar pools inactivos segun config_sistema
  if (!horariosOferta.relampago.activo)  { relampagoActivo = []; relampagoPoolCompleto = []; }
  if (!horariosOferta.destacadas.activo) { destacadasActivas = []; destacadasRotadas = []; destacadasPool = []; }
  if (!horariosOferta.especiales.activo) { especialesActivas = []; especialesPool = []; }
  // Leer configuracion de Recien Llegados desde config_sistema
  var precioMinVidriera = 2000; // fallback RECIEN_LLEGADOS_PRECIO_MIN
  var limiteRecienLlegados = 12;  // fallback RECIEN_LLEGADOS_LIMITE
  try {
    var shCfgRL = SpreadsheetApp.openById(SS_ID).getSheetByName(HOJA_CONFIG);
    if (shCfgRL) {
      var cfgRL = shCfgRL.getDataRange().getValues();
      for (var ci = 1; ci < cfgRL.length; ci++) {
        var ck = String(cfgRL[ci][0] || '').trim().toUpperCase();
        if (ck === 'RECIEN_LLEGADOS_PRECIO_MIN') {
          var cv = parseInt(cfgRL[ci][1]);
          if (!isNaN(cv) && cv >= 0) precioMinVidriera = cv;
        }
        if (ck === 'RECIEN_LLEGADOS_LIMITE') {
          var cl = parseInt(cfgRL[ci][1]);
          if (!isNaN(cl) && cl > 0) limiteRecienLlegados = cl;
        }
      }
    }
  } catch(eCfgRL) { Logger.log('Error leyendo config Recien Llegados: ' + eCfgRL); }
  try {
    var ssRL = SpreadsheetApp.openById(SS_ID);
    var shHistRL = ssRL.getSheetByName(HOJA_HISTORIAL);
    var shInvRL  = ssRL.getSheetByName(HOJA_INVENTARIO);
    if (shHistRL && shInvRL) {
      var hoyRL = new Date(); hoyRL.setHours(0,0,0,0);
      var hace2 = new Date(hoyRL.getTime() - 2 * 86400000);
      var tzRL = Session.getScriptTimeZone();

      var ultimaFechaMap = {};
      var histRowsRL = shHistRL.getDataRange().getValues();
      for (var hi = 1; hi < histRowsRL.length; hi++) {
        var hfecha = histRowsRL[hi][0];
        var hnom   = String(histRowsRL[hi][1] || '').trim().toUpperCase();
        if (!hnom) continue;
        var hfechaD = hfecha instanceof Date ? hfecha : new Date(String(hfecha));
        if (isNaN(hfechaD.getTime())) continue;
        var dNorm = new Date(hfechaD); dNorm.setHours(0,0,0,0);
        if (dNorm < hace2) continue;
        if (!ultimaFechaMap[hnom] || hfechaD > ultimaFechaMap[hnom]) {
          ultimaFechaMap[hnom] = hfechaD;
        }
      }

      var idsOtrosPools = {};
      relampagoActivo.forEach(function(p){ idsOtrosPools[p.id] = true; });
      destacadasRotadas.forEach(function(p){ idsOtrosPools[p.id] = true; });
      especialesActivas.forEach(function(p){ idsOtrosPools[p.id] = true; });

      var invRowsRL = shInvRL.getDataRange().getValues();
      var candidatosRL = [];
      for (var ri = 1; ri < invRowsRL.length; ri++) {
        var rNom = String(invRowsRL[ri][0] || '').trim().toUpperCase();
        if (!rNom || !ultimaFechaMap[rNom]) continue;
        var rStock = parseInt(invRowsRL[ri][5]) || 0;
        if (rStock <= 0) continue;
        var rPrecio = parseInt(invRowsRL[ri][1]) || 0;
        if (rPrecio < precioMinVidriera) continue; // Precio minimo vidriera (desde config_sistema)
        if (idsOtrosPools[ri]) continue;
        candidatosRL.push({
          id:           ri,
          nombre:       String(invRowsRL[ri][0]).trim(),
          precio:       parseInt(invRowsRL[ri][1]) || 0,
          categoria:    String(invRowsRL[ri][2] || '').trim(),
          stock:        rStock,
          fechaIngreso: Utilities.formatDate(ultimaFechaMap[rNom], tzRL, 'dd/MM/yyyy'),
          _ts:          ultimaFechaMap[rNom].getTime()
        });
      }

      candidatosRL.sort(function(a, b){ return b._ts - a._ts; });
      poolRecienLlegados = candidatosRL.slice(0, limiteRecienLlegados + 10).map(function(p){
        return { id: p.id, nombre: p.nombre, precio: p.precio,
                 categoria: p.categoria, stock: p.stock, fechaIngreso: p.fechaIngreso };
      });
    }
  } catch(eRL) { Logger.log('Error poolRecienLlegados: ' + eRL); }

  return {
    success: true,
    relampagoActivo:      relampagoActivo,
    relampagoPool:        relampagoPoolCompleto,
    destacadasActivas:    destacadasRotadas,
    destacadasPool:       destacadasPool,
    especialesActivas:    especialesActivas,
    poolRecienLlegados:     poolRecienLlegados,
    limiteRecienLlegados:   limiteRecienLlegados,
    limites: limites,
    horariosOferta: {
      relampago:  { inicio: _inicioRelampago,  cierre: _cierreRelampago,  activo: horariosOferta.relampago.activo  },
      destacadas: { inicio: _inicioDestacadas, cierre: _cierreDestacadas, activo: horariosOferta.destacadas.activo },
      especiales: { inicio: _inicioEspeciales, cierre: _cierreEspeciales, activo: horariosOferta.especiales.activo }
    },
    fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    stats: {
      totalRelampago:          relampagoActivo.length,
      totalDestacadas:         destacadasActivas.length,
      totalEspeciales:         especialesPool.length,
      especialesMostradas:     especialesActivas.length
    }
  };
}

// ========== MOTOR DE PUNTAJE - SUGERENCIAS ==========
function calcularMotorSugerencias() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INVENTARIO);
    const datos = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const hoy = new Date();
    const sugerencias = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila[0]) continue;
      const nombre = String(fila[0] || '').trim();
      const precio = parseInt(fila[1]) || 0;
      const stock = parseInt(fila[5]) || 0;
      const relampago = parseInt(fila[6]) || 0;
      const rotacion = fila.length > 14 ? (parseInt(fila[14]) || 0) : 0;
      if (stock <= 0 || relampago <= 0) continue;

      let diasParaVencer = null;
      try {
        const v = fila.length > 15 ? fila[15] : '';
        if (v) {
          const vDate = v instanceof Date ? v : new Date(v);
          if (!isNaN(vDate.getTime())) diasParaVencer = Math.round((vDate - hoy) / (1000 * 60 * 60 * 24));
        }
      } catch(e) {}

      let diasDesdeUltimaPromo = null;
      try {
        const up = fila.length > 17 ? fila[17] : '';
        if (up) {
          const upDate = up instanceof Date ? up : new Date(up);
          if (!isNaN(upDate.getTime())) diasDesdeUltimaPromo = Math.round((hoy - upDate) / (1000 * 60 * 60 * 24));
        }
      } catch(e) {}

      if (diasDesdeUltimaPromo !== null && diasDesdeUltimaPromo < 15) continue;

      let puntaje = 0;
      if (diasParaVencer !== null) {
        if (diasParaVencer <= 3)  puntaje += 5;
        else if (diasParaVencer <= 7 && stock >= 5) puntaje += 4;
      }
      if (rotacion >= 3) puntaje += 3;
      if (precio > 1000) puntaje += 2;

      let tipo = 'relampago';

      sugerencias.push({ id: i, nombre, stock, rotacion, diasParaVencer, diasDesdeUltimaPromo, puntaje, tipo, relampago });
    }

    sugerencias.sort((a, b) => b.puntaje - a.puntaje);
    return { success: true, sugerencias: sugerencias.slice(0, 10), fecha: Utilities.formatDate(hoy, tz, 'yyyy-MM-dd HH:mm') };
  } catch (error) {
    return { success: false, mensaje: error.toString() };
  }
}

// ========== ACTIVAR RELÁMPAGO ==========
function activarRelampago(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INVENTARIO);
    const todosLosDatos = sheet.getDataRange().getValues();
    const productoId = datos.productoId;
    const tipoRelampago = datos.tipo;
    const filaProducto = parseInt(productoId) + 1;
    if (filaProducto < 2 || filaProducto > todosLosDatos.length) return { success: false, mensaje: 'Producto no encontrado' };
    sheet.getRange(filaProducto, 7).setValue(tipoRelampago);
    if (datos.limpiarOfertas && tipoRelampago > 0) {
      sheet.getRange(filaProducto, 8).setValue(0);
      sheet.getRange(filaProducto, 9).setValue(0);
    }
    const nombreProducto = todosLosDatos[filaProducto - 1][0];
    if (tipoRelampago > 0) sheet.getRange(filaProducto, 18).setValue(new Date());
    if (tipoRelampago === 0) return { success: true, mensaje: '✅ Relámpago desactivado: ' + nombreProducto };
    const descuentos = { 11: '10%', 12: '15%', 13: '20%', 14: '25%' };
    const desc = descuentos[tipoRelampago] || tipoRelampago;
    return { success: true, mensaje: '⚡ Relámpago activado! ' + nombreProducto + ' - Descuento: ' + desc };
  } catch (error) {
    return { success: false, mensaje: 'Error: ' + error.toString() };
  }
}

// ========== GET CONFIG ==========
function getConfig() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CONFIG);
    if (!sheet) return { success: false, mensaje: 'Hoja config_sistema no encontrada' };
    const datos = sheet.getDataRange().getValues();
    const config = {};
    let seccionActual = 'general';
    for (let i = 1; i < datos.length; i++) {
      const clave = String(datos[i][0] || '').trim();
      if (!clave) continue;
      if (clave.startsWith('──') || clave.startsWith('─') || clave.startsWith('📌')) {
        const c = clave.toLowerCase();
        if (c.includes('relámpago') || c.includes('relampago')) seccionActual = 'relampago';
        else if (c.includes('destacada')) seccionActual = 'destacadas';
        else if (c.includes('especial')) seccionActual = 'especiales';
        else if (c.includes('horario')) seccionActual = 'horario';
        continue;
      }
      const vals = [];
      for (let c = 1; c <= 7; c++) {
        const v = datos[i][c];
        vals.push(v !== '' && v !== null && v !== undefined ? v : null);
      }
      const claveUnica = seccionActual + '|' + clave;
      config[claveUnica] = vals;
      if (!config[clave]) config[clave] = vals;
    }
    return { success: true, config };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ========== SET CONFIG ==========
function setConfig(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_CONFIG);
    if (!sheet) return { success: false, mensaje: 'Hoja config_sistema no encontrada' };

    const seccion  = String(datos.seccion || '').toLowerCase();
    const clave    = String(datos.clave   || '').trim();
    const diaIdx   = parseInt(datos.dia);
    const valor = (datos.valor === 'A' || datos.valor === null || datos.valor === undefined) ? 'A' :
                  (datos.valor === 1 || datos.valor === '1') ? 1 : 0;

    if (isNaN(diaIdx) || diaIdx < 0 || diaIdx > 6) return { success: false, mensaje: 'Día inválido' };

    const hojaData = sheet.getDataRange().getValues();
    let seccionActual = 'general';
    let filaObjetivo = -1;
    let filaFallback = -1;

    const claveNorm = clave.toLowerCase();

    for (let i = 1; i < hojaData.length; i++) {
      const celda = String(hojaData[i][0] || '').trim();
      if (!celda) continue;
      if (celda.startsWith('──') || celda.startsWith('─') || celda.startsWith('📌') || celda.startsWith('—')) {
        const c = celda.toLowerCase();
        if (c.includes('relámpago') || c.includes('relampago')) seccionActual = 'relampago';
        else if (c.includes('destacada')) seccionActual = 'destacadas';
        else if (c.includes('especial')) seccionActual = 'especiales';
        else if (c.includes('horario')) seccionActual = 'horario';
        continue;
      }
      const celdaNorm = celda.toLowerCase();
      if (seccionActual === seccion && celdaNorm === claveNorm) {
        filaObjetivo = i + 1;
        break;
      }
      if (filaFallback === -1 && celdaNorm === claveNorm) {
        filaFallback = i + 1;
      }
    }

    if (filaObjetivo === -1) filaObjetivo = filaFallback;
    if (filaObjetivo === -1) return { success: false, mensaje: 'No se encontró: ' + seccion + '|' + clave };

    const columna = diaIdx + 2;
    sheet.getRange(filaObjetivo, columna).setValue(valor);

    return { success: true, mensaje: 'OK: ' + seccion + '|' + clave + '[' + diaIdx + '] = ' + valor };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ========== JUEVES CERVECERO ==========
function getJuevesCervecero() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetInv    = ss.getSheetByName(HOJA_INVENTARIO);
    const sheetVentas = ss.getSheetByName('Ventas');
    const sheetCfg    = ss.getSheetByName(HOJA_CONFIG);
    if (!sheetInv) return { success: false, mensaje: 'Sin inventario' };

    let horaCierre = 22;
    let horaInicio = 14; // fallback — fila JUEVES CERVECERO columna INICIO
    let modoActivo = 'A'; // 'A', '0', '1'
    if (sheetCfg) {
      const cfgData = sheetCfg.getDataRange().getValues();
      for (let i = 0; i < cfgData.length; i++) {
        const clave = String(cfgData[i][0] || '').trim().toLowerCase();
        if (clave.includes('jueves cervecero') || clave.includes('cervecero')) {
          Logger.log('Fila cervecero encontrada en fila ' + i + '. Columnas: ' + JSON.stringify(cfgData[i]));
          Logger.log('C[2]=' + cfgData[i][2] + ' tipo=' + typeof cfgData[i][2] + ' | G[6]=' + cfgData[i][6] + ' tipo=' + typeof cfgData[i][6]);
          // Fila 26: A=JUEVES CERVECERO, B=INICIO, C=hora inicio, D=HRS, E=—, F=FINAL, G=hora cierre, H=HRS
          const ini = cfgData[i][2]; // columna C = hora inicio
          const fin = cfgData[i][6]; // columna G = hora cierre
          const modo = cfgData[i][8]; // columna I = modo (si existe)
          if (ini !== null && ini !== '') {
            var pIni;
            if (typeof ini === 'number' && ini > 0 && ini < 1) {
              pIni = ini * 24;
            } else if (typeof ini === 'string' && ini.includes('T')) {
              // Google Sheets manda hora como ISO: usar UTC horas + ajuste ARG (-3)
              var dIni = new Date(ini);
              pIni = ((dIni.getUTCHours() - 3 + 24) % 24) + dIni.getUTCMinutes() / 60;
            } else {
              pIni = typeof ini === 'number' ? ini : parseFloat(String(ini).split(':')[0]);
            }
            if (!isNaN(pIni)) horaInicio = Math.round(pIni * 100) / 100;
          }
          if (fin !== null && fin !== '') {
            var pFin;
            if (typeof fin === 'number' && fin > 0 && fin < 1) {
              pFin = fin * 24;
            } else if (typeof fin === 'string' && fin.includes('T')) {
              var dFin = new Date(fin);
              pFin = ((dFin.getUTCHours() - 3 + 24) % 24) + dFin.getUTCMinutes() / 60;
            } else {
              pFin = typeof fin === 'number' ? fin : parseFloat(String(fin).split(':')[0]);
            }
            if (!isNaN(pFin)) horaCierre = Math.round(pFin * 100) / 100;
          }
          if (modo !== null && modo !== undefined && modo !== '') modoActivo = String(modo).trim().toUpperCase();
          break;
        }
      }
    }

    const inv = sheetInv.getDataRange().getValues();
    const cervezas = [];
    for (let i = 1; i < inv.length; i++) {
      const nombre = String(inv[i][0] || '').trim();
      const cat    = String(inv[i][2] || '').trim().toUpperCase();
      const precio = parseFloat(inv[i][1]) || 0;
      const stock  = parseInt(inv[i][5]) || 0;
      const costo  = parseFloat(inv[i][18]) || 0;
      if (!nombre || stock < 3 || precio <= 0) continue;
      const esCerveza = cat.includes('CERV') ||
                        nombre.toUpperCase().includes('CERVEZA') ||
                        nombre.toUpperCase().includes(' LATA') ||
                        nombre.toUpperCase().includes('BIRRA') ||
                        nombre.toUpperCase().includes(' IPA') ||
                        nombre.toUpperCase().includes(' STOUT') ||
                        nombre.toUpperCase().includes(' PORTER');
      if (!esCerveza) continue;
      const margen = costo > 0 ? (precio - costo) / precio : 0.35;
      cervezas.push({ i, nombre, precio, stock, costo, margen });
    }

    if (!cervezas.length) return { success: true, productos: [], horaCierre };

    const ventasProd = {};
    const ticketTotales = {};
    if (sheetVentas) {
      const hoy    = new Date();
      const hace30 = new Date(hoy - 30 * 86400000);
      const ventas = sheetVentas.getDataRange().getValues();
      for (let i = 1; i < ventas.length; i++) {
        const f = ventas[i];
        const nombre = String(f[3] || '').trim();
        if (nombre.includes('TOTAL TICKET')) {
          ticketTotales[String(f[0] || '').trim()] = parseFloat(f[7]) || 0;
        }
      }
      for (let i = 1; i < ventas.length; i++) {
        const f = ventas[i];
        const fecha = f[1];
        if (!(fecha instanceof Date) || fecha < hace30) continue;
        const nombre = String(f[3] || '').trim().toUpperCase();
        const qty    = parseFloat(f[4]) || 0;
        const tk     = String(f[0] || '').trim();
        if (!nombre || nombre.startsWith('─')) continue;
        if (!ventasProd[nombre]) ventasProd[nombre] = { qty: 0, tickets: 0, ticketTotal: 0 };
        ventasProd[nombre].qty += qty;
        ventasProd[nombre].tickets++;
        ventasProd[nombre].ticketTotal += ticketTotales[tk] || 0;
      }
    }

    const maxQty    = Math.max(...cervezas.map(c => (ventasProd[c.nombre.toUpperCase()] || {qty:0}).qty), 1);
    const maxStock  = Math.max(...cervezas.map(c => c.stock), 1);
    const maxTicket = Math.max(...cervezas.map(c => (ventasProd[c.nombre.toUpperCase()] || {ticketTotal:0}).ticketTotal), 1);

    const scored = cervezas.map(c => {
      const vd       = ventasProd[c.nombre.toUpperCase()] || { qty: 0, tickets: 0, ticketTotal: 0 };
      const rotScore  = vd.qty / maxQty;
      const stockScore = c.stock / maxStock;
      const margenScore = Math.min(c.margen, 1);
      const ticketScore = vd.ticketTotal / maxTicket;
      const promoScore = (rotScore * 0.35) + (stockScore * 0.25) + (margenScore * 0.20) + (ticketScore * 0.20);
      return { ...c, promoScore, ventas30: vd.qty };
    });

    const top6 = scored.sort((a, b) => b.promoScore - a.promoScore).slice(0, 6);
    const idxGancho = top6.reduce((bestIdx, c, i) => c.margen > top6[bestIdx].margen ? i : bestIdx, 0);

    const top6Final = top6.map((c, idx) => {
      let descuento = idx === idxGancho ? 20 : (c.margen > 0.30 ? 15 : 10);
      const precioBase = Math.max(
        Math.round(c.precio * (1 - descuento / 100)),
        c.costo > 0 ? Math.ceil(c.costo * 1.05) : 0
      );
      const precioPromo99 = Math.round((precioBase - 99) / 100) * 100 + 99;
      const precioPromo = Math.min(precioPromo99, c.precio - 1);
      return { ...c, descuento, precioPromo, esGancho: idx === idxGancho };
    });

    return { success: true, productos: top6Final, horaCierre, horaInicio, modoActivo };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function logEventoCerveza(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const HOJA_LOG = 'evento_cerveza_log';
    let sheet = ss.getSheetByName(HOJA_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(HOJA_LOG);
      sheet.appendRow(['FECHA', 'PRODUCTOS', 'UNIDADES VENDIDAS', 'FACTURACIÓN', 'GANANCIA', 'TICKET PROMEDIO', 'DESCUENTO APLICADO']);
      sheet.getRange(1, 1, 1, 7).setBackground('#4a148c').setFontColor('white').setFontWeight('bold');
      [120, 200, 80, 100, 100, 100, 100].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    }
    const tz = Session.getScriptTimeZone();
    const fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    sheet.appendRow([fecha, String(datos.productos || ''), datos.unidades || 0, datos.facturacion || 0, datos.ganancia || 0, datos.ticketPromedio || 0, String(datos.descuentos || '')]);
    return { success: true };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function logInterruptor(datos) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const HOJA_LOG = 'Interruptores_Log';
    let sheet = ss.getSheetByName(HOJA_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(HOJA_LOG);
      sheet.appendRow(['FECHA', 'HORA', 'OFERTA', 'ACCIÓN', 'DÍA SEMANA', 'VENDEDOR']);
      sheet.getRange(1, 1, 1, 6).setBackground('#1a237e').setFontColor('white').setFontWeight('bold');
      sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 120);
      sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 120);
    }
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, tz, 'dd/MM/yyyy');
    const hora  = Utilities.formatDate(ahora, tz, 'HH:mm:ss');
    const diasNombres = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    const diaSemana = diasNombres[ahora.getDay()];
    const oferta   = String(datos.tipo     || '').toUpperCase();
    const accion   = datos.estado === 1 ? '🟢 ACTIVADO' : '🔴 DESACTIVADO';
    const vendedor = String(datos.vendedor || 'Desconocido');
    sheet.appendRow([fecha, hora, oferta, accion, diaSemana, vendedor]);
    return { success: true };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ========== API PRINCIPAL doGet ==========
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'vender') {
      return procesarVenta(e.parameter.data);
    }
    if (e && e.parameter && e.parameter.action === 'agregar') {
      return agregarProducto(e.parameter.data);
    }
    if (e && e.parameter && e.parameter.action === 'ingresar') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(ingresarMercaderia(datos));
    }
    if (e && e.parameter && e.parameter.action === 'getInfoProducto') {
      return respuestaJSON(obtenerInfoProducto(e.parameter.nombre || ''));
    }
    if (e && e.parameter && e.parameter.action === 'salidaInterna') {
      var datosSalida = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(registrarSalidaInterna(datosSalida));
    }
    if (e && e.parameter && e.parameter.action === 'setPausadoListaCompra') {
      var datosPausa = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(setPausadoListaCompra(datosPausa));
    }
    if (e && e.parameter && e.parameter.action === 'ajusteRapido') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(ajustarProducto(datos));
    }
    if (e && e.parameter && e.parameter.action === 'setConfig') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(setConfig(datos));
    }
    if (e && e.parameter && e.parameter.action === 'getJuevesCervecero') {
      return respuestaJSON(getJuevesCervecero());
    }
    if (e && e.parameter && e.parameter.action === 'logEventoCerveza') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(logEventoCerveza(datos));
    }
    if (e && e.parameter && e.parameter.action === 'logInterruptor') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(logInterruptor(datos));
    }
    if (e && e.parameter && e.parameter.action === 'crearOrdenMP') {
      const monto = parseFloat(e.parameter.monto) || 0;
      if (monto <= 0) return respuestaJSON({ ok: false, error: 'Monto inválido' });
      return respuestaJSON(crearOrdenMP(monto));
    }
    if (e && e.parameter && e.parameter.action === 'getReportes') {
      try {
        return respuestaJSON(getReportes());
      } catch(errReportes) {
        return respuestaJSON({ error: true, mensaje: 'Error en reportes: ' + errReportes.toString() });
      }
    }
    if (e && e.parameter && e.parameter.action === 'getOfertas') {
      try {
        return respuestaJSON(calcularOfertas());
      } catch(errOfertas) {
        return respuestaJSON({ error: true, mensaje: 'Error en ofertas: ' + errOfertas.toString() });
      }
    }
    if (e && e.parameter && e.parameter.action === 'motorSugerencias') {
      return respuestaJSON(calcularMotorSugerencias());
    }
    if (e && e.parameter && e.parameter.action === 'relampago') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(activarRelampago(datos));
    }
    if (e && e.parameter && e.parameter.action === 'getConfig') {
      return respuestaJSON(getConfig());
    }
    if (e && e.parameter && e.parameter.action === 'actualizarEstadisticas') {
      return respuestaJSON(actualizarEstadisticas());
    }
    if (e && e.parameter && e.parameter.action === 'getProveedores') {
      return respuestaJSON(leerProveedores());
    }
    if (e && e.parameter && e.parameter.action === 'guardarProveedor') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(guardarProveedor(datos));
    }
    if (e && e.parameter && e.parameter.action === 'guardarCarritoTemp') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(guardarCarritoTemp(datos));
    }
    if (e && e.parameter && e.parameter.action === 'getCarritoTemp') {
      return respuestaJSON(getCarritoTemp());
    }
    if (e && e.parameter && e.parameter.action === 'getListaCompra') {
      // Primero regenera la hoja, luego devuelve los datos
      actualizarListaCompra();
      return respuestaJSON(getListaCompraJSON());
    }
    if (e && e.parameter && e.parameter.action === 'generarListaCompra') {
      return respuestaJSON(actualizarListaCompra());
    }
    if (e && e.parameter && e.parameter.action === 'actualizarMotor') {
      actualizarMotorInventario();
      actualizarListaCompra();
      return respuestaJSON({ ok: true, msg: 'Motor y lista actualizados' });
    }
    if (e && e.parameter && e.parameter.action === 'listarFiados') {
      return respuestaJSON(listarFiados());
    }
    if (e && e.parameter && e.parameter.action === 'getDetalleTicket') {
      var ticket = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)).ticket : '';
      return respuestaJSON(getDetalleTicket(ticket));
    }
    if (e && e.parameter && e.parameter.action === 'cobrarFiado') {
      var datosCobro = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(cobrarFiado(datosCobro));
    }
    if (e && e.parameter && e.parameter.action === 'abonarFiado') {
      var datosAbono = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(abonarFiado(datosAbono));
    }
    // ── NUEVAS ACCIONES v2.0 FIADOS ──
    if (e && e.parameter && e.parameter.action === 'listarFiadosCliente') {
      var datosFC = e.postData ? JSON.parse(e.postData.contents) : JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
      return respuestaJSON(listarFiadosCliente(datosFC));
    }
    if (e && e.parameter && e.parameter.action === 'pagarFiadosSeleccionados') {
      var datosPFS = e.postData ? JSON.parse(e.postData.contents) : JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
      return respuestaJSON(pagarFiadosSeleccionados(datosPFS));
    }
    if (e && e.parameter && e.parameter.action === 'abonoFiado') {
      var datosAbono = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(abonoFiado(datosAbono));
    }
    if (e && e.parameter && e.parameter.action === 'guardarFiado') {
      var datosFiado = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(guardarFiado(datosFiado));
    }
    if (e && e.parameter && e.parameter.action === 'consultarFiado') {
      return respuestaJSON(consultarFiado(e.parameter.telefono));
    }
    if (e && e.parameter && e.parameter.action === 'listarClientes') {
      return respuestaJSON(listarClientes());
    }
    if (e && e.parameter && e.parameter.action === 'setPausadoNombre') {
      var datosPN = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(setPausadoNombre(datosPN));
    }
    if (e && e.parameter && e.parameter.action === 'registrarMovimientoCaja') {
      var datosCaja = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(registrarMovimientoCaja(datosCaja));
    }
    if (e && e.parameter && e.parameter.action === 'ventasProducto') {
      var datosVP = JSON.parse(decodeURIComponent(e.parameter.data));
      return respuestaJSON(getVentasProducto(datosVP));
    }
    if (e && e.parameter && e.parameter.action === 'resetearPlanilla') {
      var pinReset = e.parameter.pin || '';
      return respuestaJSON(resetearPlanilla(pinReset));
    }
    if (e && e.parameter && e.parameter.action === 'getConfigUI') {
      return respuestaJSON(getConfigUI());
    }
    if (e && e.parameter && e.parameter.action === 'getCajaEgresos') {
      var datosEg = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(getCajaEgresos(datosEg));
    }
    if (e && e.parameter && e.parameter.action === 'getSalidasInternas') {
      var datosSal = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(getSalidasInternas(datosSal));
    }
    if (e && e.parameter && e.parameter.action === 'getPromoEspecial') {
      return respuestaJSON(getPromoEspecial());
    }
    if (e && e.parameter && e.parameter.action === 'getHorarioLocal') {
      var datosHor = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(getHorarioLocal(datosHor));
    }
    if (e && e.parameter && e.parameter.action === 'setHorarioLocal') {
      var datosSetHor = e.parameter.data ? JSON.parse(decodeURIComponent(e.parameter.data)) : {};
      return respuestaJSON(getHorarioLocal(datosSetHor));
    }

    // ── GET PRODUCTOS (carga la tienda) ──
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INVENTARIO);
    const datos = sheet.getDataRange().getValues();
    const rotacionAuto = calcularRotacionAuto();
    const productos = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila[0]) continue;
      const nombreKey = String(fila[0] || '').trim().toUpperCase();
      const rotacionCalculada = rotacionAuto[nombreKey] || (fila.length > 14 ? (parseInt(fila[14]) || 1) : 1);

      const producto = {
        id: i,
        name: String(fila[0] || '').trim(),
        price: parseInt(fila[1]) || 0,
        category: String(fila[2] || 'ALMACEN').trim().toUpperCase(),
        stock: parseInt(fila[5]) || 0,
        descripcion: fila[3] || '',
        image: '',
        relampago: parseInt(fila[6]) || 0,
        destacada: parseInt(fila[7]) || 0,
        especial: parseInt(fila[8]) || 0,
        normal: 0,
        proveedor: String(fila[4] || '').trim(),
        vencimiento: (() => {
          try {
            const v = fila.length > 15 ? fila[15] : '';
            if (!v) return '';
            if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            return String(v).trim();
          } catch(e) { return ''; }
        })(),
        rotacion: rotacionCalculada,
        costo: fila.length > 18 ? (parseFloat(fila[18]) || 0) : 0,
        diaCritico: fila.length > 16 ? String(fila[16] || '').trim().toLowerCase() : '',
        stockMin: parseInt(fila[13]) || 0,
        prioridad: fila.length > 31 ? String(fila[31] || '').trim().toUpperCase() : 'NORMAL',
        pausadoListaCompra: fila.length > 34 ? String(fila[34] || '').trim().toUpperCase() : '',
        ultimaPromo: '',
        categoriaOferta: String(fila[9] || '').trim()
      };

      if (fila.length > 11) {
        const columnaL = fila[11];
        if (columnaL !== undefined && columnaL !== '' && columnaL !== null) {
          const num = parseInt(columnaL);
          if (!isNaN(num) && num > 0) producto.normal = num;
        }
      }
      if (producto.normal === 0 && fila.length > 10) {
        const columnaK = fila[10];
        if (columnaK !== undefined && columnaK !== '' && columnaK !== null) {
          const num = parseInt(columnaK);
          if (!isNaN(num) && num > 0) producto.normal = num;
        }
      }

      productos.push(producto);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        productos: productos,
        total: productos.length,
        fecha: new Date().toISOString(),
        version: '7.0'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('❌ Error API:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ error: true, mensaje: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== PROCESAR VENTA ==========
function procesarVenta(dataStr) {
  try {
    const payload = JSON.parse(decodeURIComponent(dataStr));
    const items = Array.isArray(payload) ? payload : (payload.items || []);
    const metodoPago = Array.isArray(payload) ? 'efectivo' : (payload.metodoPago || 'efectivo');
    const vendedor = Array.isArray(payload) ? '' : (payload.vendedor || '');
    const ventaPayload = Array.isArray(payload) ? {} : payload;

    if (!items || !Array.isArray(items) || items.length === 0) return respuestaJSON({ success: false, mensaje: 'Sin items' });

    if (payload.idVenta) {
      var props = PropertiesService.getScriptProperties();
      var procesados2 = props.getProperty('ventas_procesadas') || '';
      if (procesados2.indexOf(payload.idVenta) !== -1) {
        return respuestaJSON({ success: true, duplicate: true });
      }
      var lista2 = procesados2 ? procesados2.split(',') : [];
      lista2.push(payload.idVenta);
      if (lista2.length > 100) lista2 = lista2.slice(-100);
      props.setProperty('ventas_procesadas', lista2.join(','));
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetInventario = ss.getSheetByName(HOJA_INVENTARIO);
    const datos = sheetInventario.getDataRange().getValues();
    const errores = [];
    const procesados = [];
    const fecha = new Date();
    let totalVenta = 0;

    for (const item of items) {
      const filaIndex = item.esPeso ? (item.idOriginal || item.id) : item.id;
      if (filaIndex < 1 || filaIndex >= datos.length) { errores.push('ID ' + item.id + ' no encontrado'); continue; }
      const fila = datos[filaIndex];
      const nombreEnSheet = String(fila[0] || '').trim();
      const stockActual = parseInt(fila[5]) || 0;
      const precioUnit = parseInt(fila[1]) || 0;
      const qty = parseInt(item.qty) || 0;
      const esPeso = item.esPeso === true;
      const gramos = parseInt(item.gramos) || 0;
      const precioVenta = item.precioVenta !== undefined && item.precioVenta !== null && item.precioVenta !== '' ? parseInt(item.precioVenta) : precioUnit;
      const precioOriginal = parseInt(item.precioOriginal) || 0;
      const precioModificado = item.precioModificado === true;
      if (esPeso && gramos <= 0) { errores.push(nombreEnSheet + ': gramos inválidos'); continue; }
      if (!esPeso && qty <= 0) continue;

      if (esPeso) {
        const kgVendidos = gramos / 1000;
        const stockEnKg  = parseFloat(fila[5]) || 0;
        const nuevoStockKg = Math.max(0, Math.round((stockEnKg - kgVendidos) * 1000) / 1000);
        sheetInventario.getRange(filaIndex + 1, 6).setValue(nuevoStockKg);
        const nombreConPeso = gramos > 0 ? nombreEnSheet + ' (' + gramos + 'gr)' : nombreEnSheet;
        const kgDisplay = Math.round(kgVendidos * 1000) / 1000;
        totalVenta += precioVenta;
        procesados.push({ name: nombreConPeso, qty: kgDisplay, nuevoStock: nuevoStockKg, precio: precioVenta, esPeso: true, gramos });
      } else {
        if (stockActual < qty) errores.push(nombreEnSheet + ': stock insuficiente (tiene ' + stockActual + ', vendió ' + qty + ')');
        const nuevoStock = Math.max(0, stockActual - qty);
        sheetInventario.getRange(filaIndex + 1, 6).setValue(nuevoStock);
        totalVenta += precioUnit * qty;
        procesados.push({ name: nombreEnSheet, qty, nuevoStock, precio: precioVenta, precioModificado, precioOriginal });
      }
    }

    var ticketStr = null;
    try {
      const sheetVentas = ss.getSheetByName('Ventas');
      if (sheetVentas) {
        const hora = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'HH:mm');
        const filasTotales = sheetVentas.getLastRow();
        let ultimoTicket = 0;
        if (filasTotales > 1) {
          const colA = sheetVentas.getRange(2, 1, filasTotales - 1, 1).getValues();
          for (const fila of colA) {
            const num = parseInt(fila[0]);
            if (!isNaN(num) && num > ultimoTicket) ultimoTicket = num;
          }
        }
        const nroTicket = ultimoTicket + 1;
        ticketStr = String(nroTicket).padStart(4, '0');

        for (const p of procesados) {
          const filaActual = sheetVentas.getLastRow() + 1;
          const celdaTicket = sheetVentas.getRange(filaActual, 1);
          celdaTicket.setNumberFormat('@STRING@');
          celdaTicket.setValue(ticketStr);
          const subtotal = p.esPeso ? p.precio : p.precio * p.qty;
          const cantDisplay = p.esPeso ? p.qty : p.qty;
          const nombreDisplay = p.precioModificado
            ? p.name + (p.precioOriginal ? ' [PRECIO ESP: $' + p.precioOriginal + '→$' + p.precio + ']' : ' [PRECIO ESPECIAL]')
            : p.name;
          sheetVentas.getRange(filaActual, 2, 1, 6).setValues([[fecha, hora, nombreDisplay, cantDisplay, p.precio, subtotal]]);
        }

        const totalUnidades = procesados.reduce((s, p) => s + (p.esPeso ? 1 : p.qty), 0);
        const metodoPagoEmoji = metodoPago === 'posnet' ? '💳' : metodoPago === 'transferencia' ? '📱' : metodoPago === 'FIADO' ? '🧾' : '💵';
        const clienteInfo = (metodoPago === 'FIADO' && ventaPayload.clienteFiado)
          ? ' | 👤 ' + ventaPayload.clienteFiado + (ventaPayload.telefonoFiado ? ' · ' + ventaPayload.telefonoFiado : '')
          : '';
        sheetVentas.appendRow(['', fecha, hora, '─── TOTAL TICKET N° ' + ticketStr + ' ───', totalUnidades, metodoPagoEmoji + ' ' + metodoPago.toUpperCase() + clienteInfo, vendedor, totalVenta]);
      }
    } catch (e) {
      console.warn('No se pudo registrar en hoja Ventas:', e);
    }

    return respuestaJSON({ success: true, procesados: procesados.length, total: totalVenta, ticket: ticketStr, errores: errores, fecha: fecha.toISOString() });
  } catch (error) {
    return respuestaJSON({ success: false, mensaje: error.toString() });
  }
}

// ========== AGREGAR PRODUCTO NUEVO ==========
function agregarProducto(dataStr) {
  try {
    const datos = JSON.parse(decodeURIComponent(dataStr));
    if (!datos.name || !datos.name.trim()) return respuestaJSON({ success: false, mensaje: 'El nombre es obligatorio' });
    if (!datos.price || parseInt(datos.price) <= 0) return respuestaJSON({ success: false, mensaje: 'El precio debe ser mayor a 0' });

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(HOJA_INVENTARIO);
    const datosActuales = sheet.getDataRange().getValues();
    const nombreNuevo = datos.name.trim().toUpperCase().replace(/\s+/g, ' ');

    for (let i = 1; i < datosActuales.length; i++) {
      const nombreExistente = String(datosActuales[i][0] || '').trim().toUpperCase().replace(/\s+/g, ' ');
      if (nombreExistente === nombreNuevo) return respuestaJSON({ success: false, mensaje: 'Ya existe: "' + datosActuales[i][0] + '" en fila ' + (i + 1) });
    }

    const totalFilas = datosActuales.length;
    const idDesc = 'ID' + String(totalFilas).padStart(3, '0');
    const nuevaFila = [
      datos.name.trim().toUpperCase(), parseInt(datos.price) || 0,
      String(datos.category || 'ALMACEN').trim().toUpperCase(), idDesc,
      String(datos.proveedor || '').trim().toUpperCase(), parseInt(datos.stock) || 0,
      0, 0, 0, 0, 0, 0, 0, 0, 0, datos.vencimiento || ''
    ];
    sheet.appendRow(nuevaFila);

    try {
      const hojaHistorial = ss.getSheetByName(HOJA_HISTORIAL);
      if (hojaHistorial) {
        hojaHistorial.appendRow([
          new Date(), datos.name.trim().toUpperCase(), parseInt(datos.stock) || 0,
          parseInt(datos.stock) || 0, 'ALTA',
          String(datos.proveedor || '').trim().toUpperCase(),
          parseInt(datos.precioCosto) || 0, parseInt(datos.price) || 0,
          datos.vencimiento ? new Date(datos.vencimiento) : ''
        ]);
      }
    } catch(eHist) { console.warn('⚠️ No se pudo escribir en Historial:', eHist.toString()); }

    return respuestaJSON({ success: true, mensaje: '✅ "' + datos.name.trim() + '" agregado correctamente', fila: totalFilas + 1, id: totalFilas });
  } catch (error) {
    return respuestaJSON({ success: false, mensaje: error.toString() });
  }
}

// ========== HELPERS ==========
function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pruebaRapida() {
  const resultado = doGet(null);
  const json = JSON.parse(resultado.getContent());
  console.log('Versión API:', json.version);
  console.log('Total productos:', json.total);
  return json;
}

// ========== REPORTES COMPLETOS ==========
function getReportes() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetVentas = ss.getSheetByName('Ventas');
    const sheetInv    = ss.getSheetByName(HOJA_INVENTARIO);
    if (!sheetVentas) return { error: 'No existe hoja Ventas' };

    const tz = Session.getScriptTimeZone();
    const filas = sheetVentas.getDataRange().getValues();
    const hoy = new Date();
    const costos = {};
    const sheetHist2 = ss.getSheetByName(HOJA_HISTORIAL);
    const stockInfo = {};

    if (sheetInv) {
      const inv = sheetInv.getDataRange().getValues();
      const _aliasEspecificos = {};
          // Agregar acá alias específicos de Patagonia Salvaje si hace falta
          // Ejemplo: _aliasEspecificos["HAMBU DOBLE"] = "HAMBURGUESA DOBLE CHEDDAR";
          _aliasEspecificos["PURE DE TOMATE NOEL 530GR"] = "PURE DE TOMATE NOEL 520GR";
          _aliasEspecificos["RAZA COMIDA DE GATO 1KG"] = "RAZA GATO CARNE PESCADO ARROZ 1KG";
          _aliasEspecificos["RAZA COMIDA DE GATO 500GR"] = "RAZA GATO CARNE PESCADO ARROZ 500GR";
          _aliasEspecificos["REBOSADOR 500GR"] = "REBOSADOR PREFERIDO 500GR";
          _aliasEspecificos["SAL DOS ESTRELLAS ENTREFINA 1KG"] = "SAL  ENTREFINA DOS ESTRELLAS  1KG";
          _aliasEspecificos["SAL FINA 500GR"] = "SAL FINA DOS ESTRELLAS 500GR";
          _aliasEspecificos["TOMATE PERITAS EN LATA ARCOR 400GR"] = "TOMATE PELADO PERITA ARCOR LATA 400GR";
          _aliasEspecificos["JUGO BAGGIO CAJA 1,5 LITROS NARANJA"] = "JUGO BAGGIO FRESH CAJA 1,5 LITROS NARANJA";
          _aliasEspecificos["NARANJAS KG"] = "NARANJA X KG";
          _aliasEspecificos["QUESO LA PAULINA BARRA SANDWICH 1KG"] = "QUESO LA PAULINA BARRA SANDWICH X KG";
          _aliasEspecificos["PAPELILLO"] = "PAPELILLO OCB SEDA";
          _aliasEspecificos["JUGO EN POLVO ARCOR SABOR A MULTIFRUTA 15G 1L"] = "JUGO EN POLVO MULTIFRUTA ARCOR 15GR";
          _aliasEspecificos["ARROZ 1 KG PRIMOR"] = "ARROZ 1KG PRIMOR";
          _aliasEspecificos["PROCENEX"] = "PROCENEX LIMPIADOR PISOS 900ML";
          _aliasEspecificos["MISTER TRAPO PISO 62X48"] = "MISTER TRAPO PISO 62X48 MR";
          _aliasEspecificos["ALFAJOR PEPITOS X3 57GR"] = "ALFAJOR TRIPLE PEPITOS 57GR";

      const _aliasOrigen = {};

      for (let i = 1; i < inv.length; i++) {
        const nom = String(inv[i][0]||'').trim().toUpperCase();
        const sto = parseInt(inv[i][5]) || 0;
        const cos = parseFloat(inv[i][18]) || 0;
        if (!nom) continue;

        stockInfo[nom] = sto;
        if (cos > 0) costos[nom] = cos;
        if (String(inv[i][34] || '').trim().toUpperCase() === 'SI') {
          stockInfo['__pausado__' + nom] = true;
        }

        Object.keys(_aliasEspecificos).forEach(function(viejo) {
          if (_aliasEspecificos[viejo] === nom && stockInfo[viejo] === undefined) {
            stockInfo[viejo] = sto;
            if (cos > 0 && !costos[viejo]) costos[viejo] = cos;
            Logger.log('⚠️ Alias específico: "' + viejo + '" → "' + nom + '"');
          }
        });

        function _normBase(s) {
          return s.replace(/\s+/g, ' ').replace(/,/g, '.').trim();
        }

        var base = _normBase(nom);
        var candidatos = new Set();

        candidatos.add(base.replace(/ X KG$/, ' KG').replace(/ X KG /, ' KG ').trim());
        candidatos.add(base.replace(/\bFRESH\b/g, '').replace(/\s+/g, ' ').trim());
        candidatos.add(base.replace(/\bCON FRUTAS?\b/gi, '').replace(/\s+/g, ' ').trim());
        candidatos.add(base.replace(/\bDOBLE CREMA\b/gi, '').replace(/\s+/g, ' ').trim());
        candidatos.add(base.replace(/\bX (\d+GR)\b/g, '$1').replace(/\s+/g, ' ').trim());
        candidatos.add(base.replace(/ X KG$/, ' KG').replace(/ X KG /, ' KG ').replace(/\bFRESH\b/g, '').replace(/\s+/g, ' ').trim());

        candidatos.forEach(function(alias) {
          if (!alias || alias === nom) return;
          if (stockInfo[alias] !== undefined) {
            if (_aliasOrigen[alias] && _aliasOrigen[alias] !== nom) {
              Logger.log('⚠️ Colisión de alias: "' + alias + '" entre "' + _aliasOrigen[alias] + '" y "' + nom + '"');
            }
            return;
          }
          stockInfo[alias] = sto;
          if (cos > 0 && !costos[alias]) costos[alias] = cos;
          _aliasOrigen[alias] = nom;
          Logger.log('⚠️ Alias usado para stockInfo: "' + alias + '" -> "' + nom + '"');
        });
      }
    }
    if (sheetHist2) {
      const hist = sheetHist2.getDataRange().getValues();
      for (let i = 1; i < hist.length; i++) {
        const nom = String(hist[i][1]||'').trim().toUpperCase();
        const cos = parseFloat(hist[i][6]) || 0;
        if (nom && cos > 0 && !costos[nom]) costos[nom] = cos;
      }
    }

    const vencProximos = [];
    if (sheetInv) {
      const inv = sheetInv.getDataRange().getValues();
      const limite = new Date(); limite.setDate(limite.getDate() + 30);
      for (let i = 1; i < inv.length; i++) {
        const nom = String(inv[i][0]||'').trim();
        const stock = parseInt(inv[i][5])||0;
        const venc = inv[i].length > 15 ? inv[i][15] : '';
        if (!venc || !nom) continue;
        let fechaVenc;
        try { fechaVenc = venc instanceof Date ? venc : new Date(venc); } catch(e) { continue; }
        if (isNaN(fechaVenc.getTime())) continue;
        const dias = Math.round((fechaVenc - hoy) / 86400000);
        if (dias <= 30) vencProximos.push({ nombre: nom, stock, fecha: Utilities.formatDate(fechaVenc, tz, 'yyyy-MM-dd'), dias, vencido: dias < 0 });
      }
      vencProximos.sort((a,b) => a.dias - b.dias);
    }

    const items = [];
    const totales = [];

    for (let i = 1; i < filas.length; i++) {
      const f = filas[i];
      const colA  = String(f[0]||'').trim();
      const fecha = f[1];
      if (!fecha || !(fecha instanceof Date)) continue;
      const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
      const rawHora = f[2];
      let hora = '';
      if (rawHora instanceof Date) {
        hora = Utilities.formatDate(rawHora, tz, 'HH:mm');
      } else {
        hora = String(rawHora||'').trim();
        const hm = hora.match(/\b(\d{1,2}):(\d{2}):/);
        if (hm) hora = hm[1].padStart(2,'0') + ':' + hm[2];
      }
      const nombre = String(f[3]||'').trim();

      if (nombre.includes('TOTAL TICKET')) {
        const ticketMatch = nombre.match(/N[°ºo]?\s*(\d+)/i);
        const ticket = ticketMatch ? String(ticketMatch[1]).padStart(4,'0') : colA;
        const metodo   = String(f[5]||'').replace(/[💵📱💳🔀🏦]/g,'').trim().toUpperCase();
        const vendedor = String(f[6]||'').trim();
        const total    = parseFloat(f[7]) || 0;
        totales.push({ fechaStr, hora, ticket, metodo, vendedor, total });
      } else if (nombre && !nombre.startsWith('─')) {
        const ticketRaw = String(colA||'').trim();
        const ticket    = ticketRaw.match(/^\d+$/) ? ticketRaw.padStart(4,'0') : ticketRaw;
        const qtyRaw   = f[4];
        const qtyStr   = String(qtyRaw||'').trim();
        let qty        = parseFloat(qtyStr.replace('gr','').replace('kg','')) || 0;
        const esQtyEnGramos = /\(\d+gr\)/i.test(nombre) && qty >= 50;
        if (esQtyEnGramos) qty = qty / 1000;
        const precio   = parseFloat(f[5]) || 0;
        const subtotal = parseFloat(f[6]) || 0;
        const nomLimpio = nombre.replace(/\[PRECIO ESP[^]]*\]/,'').replace(/\s*\(\d+gr\)/i,'').replace(/\s*\(\d+\.?\d*kg\)/i,'').trim().toUpperCase();
        const costo    = costos[nomLimpio] || 0;
        const tieneCosto = costo > 0;
        const gananciaReal = tieneCosto ? subtotal - (costo * qty) : null;
        const gananciaAprox = subtotal * 0.4;
        items.push({ fechaStr, hora, ticket, nombre: nomLimpio, qty, precio, subtotal, costo, gananciaReal, gananciaAprox, tieneCosto });
      }
    }

    const mapaGanancia = {};
    for (const it of items) {
      if (!mapaGanancia[it.nombre]) {
        mapaGanancia[it.nombre] = { nombre: it.nombre, qtdVendida: 0, ingresos: 0, gananciaReal: 0, gananciaAprox: 0, tieneCosto: it.tieneCosto, costo: it.costo };
      }
      const m = mapaGanancia[it.nombre];
      m.qtdVendida  += it.qty;
      m.ingresos    += it.subtotal;
      m.gananciaAprox += it.gananciaAprox;
      if (it.gananciaReal !== null) { m.gananciaReal += it.gananciaReal; m.tieneCosto = true; }
    }

    const topGanancia = Object.values(mapaGanancia)
      .map(p => ({ ...p, ganancia: p.tieneCosto ? p.gananciaReal : p.gananciaAprox, esReal: p.tieneCosto }))
      .filter(p => p.ganancia > 0)
      .sort((a, b) => b.ganancia - a.ganancia)
      .slice(0, 20);

    var recienLlegados = [];
    try {
      var hoy3 = new Date();
      hoy3.setHours(0, 0, 0, 0);
      var hace3 = new Date(hoy3.getTime() - 3 * 86400000);

      var nomIngresados = {};
      if (sheetHist2) {
        var histRL = sheetHist2.getDataRange().getValues();
        for (var hi = 1; hi < histRL.length; hi++) {
          var hFecha = histRL[hi][0];
          var hNom   = String(histRL[hi][1] || '').trim().toUpperCase();
          if (!hNom) continue;
          var hFechaD;
          if (hFecha instanceof Date) {
            hFechaD = hFecha;
          } else {
            var partes = String(hFecha).match(/(\d+)\/(\d+)\/(\d+)/);
            if (!partes) continue;
            hFechaD = new Date(parseInt(partes[3]), parseInt(partes[2])-1, parseInt(partes[1]));
          }
          if (isNaN(hFechaD.getTime())) continue;
          var dNorm = new Date(hFechaD); dNorm.setHours(0,0,0,0);
          if (dNorm < hace3) continue;
          if (!nomIngresados[hNom] || hFechaD > nomIngresados[hNom]) {
            nomIngresados[hNom] = hFechaD;
          }
        }
      }

      if (sheetInv) {
        var invRL = sheetInv.getDataRange().getValues();
        for (var ri = 1; ri < invRL.length; ri++) {
          var rNom = String(invRL[ri][0] || '').trim().toUpperCase();
          if (!rNom || !nomIngresados[rNom]) continue;
          var rStock = parseInt(invRL[ri][5]) || 0;
          if (rStock <= 0) continue;
          recienLlegados.push({
            nombre:       String(invRL[ri][0]).trim(),
            precio:       parseInt(invRL[ri][1]) || 0,
            categoria:    String(invRL[ri][2] || '').trim(),
            stock:        rStock,
            relampago:    parseInt(invRL[ri][6]) || 0,
            destacada:    parseInt(invRL[ri][7]) || 0,
            especial:     parseInt(invRL[ri][8]) || 0,
            fechaIngreso: Utilities.formatDate(nomIngresados[rNom], tz, 'dd/MM/yyyy')
          });
        }
      }
      recienLlegados.sort(function(a, b) {
        return b.fechaIngreso.localeCompare(a.fechaIngreso);
      });
    } catch(eRL) { Logger.log('Error recienLlegados: ' + eRL); }

    return { items, totales, vencProximos, stockInfo, topGanancia, recienLlegados };
  } catch(e) {
    return { error: e.toString() };
  }
}

// ========== MERCADO PAGO ==========
function crearOrdenMP(monto) {
  try {
    const payload = {
      items: [{ title: 'Compra Patagonia Salvaje', quantity: 1, unit_price: parseFloat(monto), currency_id: 'ARS' }],
      payment_methods: { excluded_payment_types: [], installments: 1 },
      statement_descriptor: 'PATAGONIA SALVAJE',
      expires: true,
      expiration_date_from: new Date().toISOString(),
      expiration_date_to: new Date(Date.now() + 30 * 60 * 1000).toISOString()
    };
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + MP_ACCESS_TOKEN, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    if (data.init_point) return { ok: true, url: data.init_point, id: data.id };
    return { ok: false, error: JSON.stringify(data) };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ========== CARRITO TEMPORAL ==========
function guardarCarritoTemp(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('carrito_temp');
    if (!sh) {
      sh = ss.insertSheet('carrito_temp');
      sh.getRange(1,1,1,3).setValues([['timestamp','vendedor','items_json']]);
    }
    const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
    if (!data.items || data.items.length === 0) {
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, 3).clearContent();
      return { success: true, action: 'cleared' };
    }
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, 3).clearContent();
    sh.getRange(2,1,1,3).setValues([[new Date().toISOString(), data.vendedor || '', JSON.stringify(data.items)]]);
    return { success: true, action: 'saved', count: data.items.length };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getCarritoTemp() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('carrito_temp');
    if (!sh || sh.getLastRow() < 2) return { success: true, items: [] };
    const row = sh.getRange(2,1,1,3).getValues()[0];
    const ts = row[0];
    if (ts && (Date.now() - new Date(ts).getTime()) > 8 * 60 * 60 * 1000) {
      sh.getRange(2, 1, 1, 3).clearContent();
      return { success: true, items: [] };
    }
    const items = row[2] ? JSON.parse(row[2]) : [];
    return { success: true, items, vendedor: row[1] || '', ts: ts ? new Date(ts).toISOString() : '' };
  } catch(e) {
    return { success: false, error: e.toString(), items: [] };
  }
}

// ========== LISTA DE COMPRA JSON ==========
function getListaCompraJSON() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    var shLista = ss.getSheetByName('Lista de compra');
    if (!shLista) return { ok: false, error: 'Hoja "Lista de compra" no encontrada.' };
    var listaData = shLista.getDataRange().getValues();
    if (listaData.length <= 1) return { ok: false, error: 'Lista vacía.' };

    var shInv = ss.getSheetByName('inventario');
    var invMap = {};
    if (shInv) {
      var invData = shInv.getDataRange().getValues();
      for (var i = 1; i < invData.length; i++) {
        var row = invData[i];
        var nombre = String(row[0] || '').trim();
        if (!nombre) continue;
        var nombreKey = nombre.toUpperCase().replace(/\s+/g,' ');
        invMap[nombreKey] = {
          categoria: String(row[2] || '').trim().toUpperCase(),
          v7:        parseFloat(row[21]) || 0,
          v30:       parseFloat(row[22]) || 0,
          diasStock: row[24] !== '' && row[24] !== null ? (parseFloat(row[24]) || 0) : null,
          riesgo:    String(row[30] || '').trim().toUpperCase(),
          prioridad: String(row[31] || '').trim().toUpperCase(),
          pisoManu:  row[32] !== '' && row[32] !== null ? parseFloat(row[32]) : null,
          alertaPiso: String(row[33] || '').trim().toUpperCase(),
          pausado:   String(row[34] || '').trim().toUpperCase() === 'SI'
        };
      }
    }

    var costoMap = {};
    var shHist = ss.getSheetByName(HOJA_HISTORIAL);
    if (shHist) {
      var histData = shHist.getDataRange().getValues();
      for (var h = histData.length - 1; h >= 1; h--) {
        var nomHist = String(histData[h][1] || '').trim().toUpperCase();
        var costoHist = parseFloat(histData[h][6]) || 0;
        if (nomHist && costoHist > 0 && !costoMap[nomHist]) {
          costoMap[nomHist] = costoHist;
        }
      }
    }

    var items = [];
    var totalPesos = 0;
    var urgentes = 0;

    for (var j = 1; j < listaData.length; j++) {
      var lr = listaData[j];
      var nombre = String(lr[0] || '').trim();
      if (!nombre || nombre.startsWith('──')) continue;

      var stock     = parseFloat(lr[1]) || 0;
      var minimo    = parseFloat(lr[2]) || 0;
      var cantidad  = parseFloat(lr[3]) || 0;
      var precio    = parseFloat(lr[4]) || 0;
      var total     = parseFloat(lr[5]) || 0;
      var proveedor = String(lr[6] || '').trim();

      var inv = invMap[nombre.toUpperCase().replace(/\s+/g,' ')] || {};

      var catInv = inv.categoria || '';
      if (catInv === 'COMPUTACION' || catInv === 'ELECTRONICA' || catInv === 'TECNOLOGIA') continue;

      var diasStock = inv.diasStock !== undefined ? inv.diasStock : null;
      var v7  = inv.v7  || 0;
      var v30 = inv.v30 || 0;
      var riesgo    = inv.riesgo    || '';
      var prioridad = inv.prioridad || '';
      var alertaPiso = inv.alertaPiso || '';
      var pausado   = inv.pausado   || false;

      if (pausado) {
        items.push({
          nombre, stock, minimo, cantidad, precio, costo: costoMap[nombre.trim().toUpperCase()] || 0,
          total, totalVenta: total, proveedor, diasStock, urgencia: 99,
          prioridad: 'PAUSADO', pausado: true
        });
        continue;
      }

      var urgencia;
      var altaRotacion = v30 > 15;

      if (alertaPiso === 'PEDIR YA') {
        urgencia = 0;
      } else if ((altaRotacion && stock === 0) || prioridad === 'EMPUJAR YA') {
        urgencia = 1;
      } else if (riesgo === 'ALERTA') {
        urgencia = 2;
      } else if (diasStock !== null && diasStock <= 3) {
        urgencia = 3;
      } else if (diasStock !== null && diasStock <= 7) {
        urgencia = 4;
      } else if (v7 >= 2 && stock === 0) {
        urgencia = 5;
      } else {
        urgencia = 7;
      }

      if (urgencia <= 2) urgentes++;
      var nomKey = nombre.trim().toUpperCase();
      var costo = costoMap[nomKey] || 0;
      var totalCosto = costo > 0 ? Math.round(costo * cantidad) : 0;
      totalPesos += totalCosto > 0 ? totalCosto : total;

      items.push({
        nombre, stock, minimo, cantidad,
        precio, costo,
        total:      totalCosto > 0 ? totalCosto : total,
        totalVenta: total,
        proveedor, diasStock, urgencia,
        prioridad: prioridad || (urgencia <= 2 ? 'URGENTE' : 'OK')
      });
    }

    items.sort(function(a, b) { return a.urgencia - b.urgencia; });

    return { ok: true, items, urgentes, totalPesos, generadoEn: new Date().toISOString() };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ========== FIADOS ==========
var HOJA_FIADOS = 'FIADOS';

function listarFiados() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { ok: false, error: 'Hoja FIADOS no encontrada' };
    var datos = sh.getDataRange().getValues();
    var hoy = new Date(); hoy.setHours(0,0,0,0);
    var fiados = [];
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue;
      var estado = String(fila[10] || '').toUpperCase();
      var vencFecha = fila[9] ? new Date(String(fila[9]).split('T')[0] + 'T12:00:00') : null;
      if (vencFecha && vencFecha < hoy && estado !== 'PAGADO') estado = 'VENCIDO';
      // Pagados: incluir solo los de los últimos 30 días
      if (estado === 'PAGADO') {
        var fechaPago = fila[11] ? new Date(String(fila[11]).split('T')[0] + 'T12:00:00') : null;
        var hace30 = new Date(hoy); hace30.setDate(hace30.getDate() - 30);
        if (!fechaPago || fechaPago < hace30) continue;
      }
      fiados.push({
        fila:            i + 1,
        idFiado:         String(fila[0] || ''),
        fecha:           String(fila[1] || ''),
        ticket:          String(fila[2] || ''),
        cliente:         String(fila[3] || ''),
        telefono:        String(fila[4] || ''),
        descripcion:     String(fila[5] || ''),
        cantItems:       fila[6] || 0,
        total:           parseFloat(fila[8]) || 0,
        fechaVencimiento: fila[9] ? String(fila[9]).split('T')[0] : '',
        fechaPago:       fila[11] ? String(fila[11]).split('T')[0] : '',
        metodoPago:      String(fila[12] || ''),
        estado:          estado
      });
    }
    fiados.sort(function(a, b) {
      if (a.estado === 'VENCIDO' && b.estado !== 'VENCIDO') return -1;
      if (b.estado === 'VENCIDO' && a.estado !== 'VENCIDO') return 1;
      return (a.fechaVencimiento || '').localeCompare(b.fechaVencimiento || '');
    });
    return { ok: true, fiados };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function abonarFiado(datos) {
  try {
    if (!datos.idFiado) return { success: false, mensaje: 'Falta ID del fiado' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { success: false, mensaje: 'Hoja FIADOS no encontrada' };
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(datos.idFiado)) continue;
      var totalActual = parseFloat(rows[i][8]) || 0;
      var abono = parseFloat(datos.abono) || 0;
      if (abono <= 0) return { success: false, mensaje: 'Monto inválido' };
      var nuevoTotal = Math.max(0, totalActual - abono);
      sh.getRange(i + 1, 9).setValue(nuevoTotal);
      var obsActual = String(rows[i][13] || '');
      var tz = Session.getScriptTimeZone();
      var fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
      var nuevaObs = (obsActual ? obsActual + ' | ' : '') +
        'Abono $' + abono + ' (' + (datos.metodoPago||'EFECTIVO') + ') ' + fecha;
      sh.getRange(i + 1, 14).setValue(nuevaObs);
      if (nuevoTotal === 0) {
        sh.getRange(i + 1, 11).setValue('PAGADO');
        sh.getRange(i + 1, 12).setValue(fecha);
        sh.getRange(i + 1, 13).setValue(datos.metodoPago || 'EFECTIVO');
      }
      return { success: true, nuevoTotal };
    }
    return { success: false, mensaje: 'Fiado no encontrado' };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function cobrarFiado(datos) {
  try {
    if (!datos.idFiado) return { success: false, mensaje: 'Falta ID del fiado' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { success: false, mensaje: 'Hoja FIADOS no encontrada' };
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(datos.idFiado)) {
        var tz = Session.getScriptTimeZone();
      var fechaPago = datos.fechaPago || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
        if (datos.esAbonoParcial) {
          var obsActual = String(rows[i][13] || '');
          var nuevaObs = obsActual ? obsActual + ' | ' + (datos.obs||'') : (datos.obs||'');
          sh.getRange(i + 1, 14).setValue(nuevaObs);
        } else {
          sh.getRange(i + 1, 11).setValue('PAGADO');
          sh.getRange(i + 1, 12).setValue(fechaPago);
          sh.getRange(i + 1, 13).setValue(datos.metodoPago || 'EFECTIVO');
          if (datos.obs) {
            var obsAct = String(rows[i][13] || '');
            sh.getRange(i + 1, 14).setValue(obsAct ? obsAct + ' | ' + datos.obs : datos.obs);
          }
        }
        return { success: true };
      }
    }
    return { success: false, mensaje: 'Fiado no encontrado: ' + datos.idFiado };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

var COL_PAUSADO_LC = 35;

function setPausadoNombre(datos) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var sh  = ss.getSheetByName(HOJA_INVENTARIO);
    if (!sh) return { success: false, mensaje: 'Hoja inventario no encontrada' };
    var rows = sh.getDataRange().getValues();
    var nom = String(datos.nombre || '').trim().toUpperCase();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toUpperCase() === nom) {
        sh.getRange(i + 1, COL_PAUSADO_LC).setValue(datos.pausado || '');
        return { success: true };
      }
    }
    return { success: false, mensaje: 'Producto no encontrado: ' + datos.nombre };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function setPausadoListaCompra(datos) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var sh  = ss.getSheetByName(HOJA_INVENTARIO);
    if (!sh) return { success: false, mensaje: 'Hoja inventario no encontrada' };
    var filaIdx = parseInt(datos.id);
    if (!filaIdx || filaIdx < 1) return { success: false, mensaje: 'ID inválido' };
    sh.getRange(filaIdx + 1, COL_PAUSADO_LC).setValue(datos.pausado || '');
    return { success: true };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function registrarSalidaInterna(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var shInv  = ss.getSheetByName(HOJA_INVENTARIO);
    var shSal  = ss.getSheetByName('SALIDAS');

    if (!shInv) return { success: false, mensaje: 'Hoja inventario no encontrada' };

    if (!shSal) {
      shSal = ss.insertSheet('SALIDAS');
      shSal.getRange(1, 1, 1, 10).setValues([[
        'Fecha', 'Hora', 'Producto', 'ID', 'Cantidad', 'Costo Unit.', 'Precio Venta', 'Motivo', 'Observación', 'Vendedor'
      ]]);
      shSal.getRange(1, 1, 1, 10).setFontWeight('bold');
    }

    var invData = shInv.getDataRange().getValues();
    var filaIdx = datos.id;
    if (!filaIdx || filaIdx < 1 || filaIdx >= invData.length) {
      return { success: false, mensaje: 'Producto no encontrado' };
    }

    var stockActual = parseFloat(invData[filaIdx][5]) || 0;
    var cantidad    = parseFloat(datos.cantidad) || 0;

    if (cantidad <= 0) return { success: false, mensaje: 'Cantidad inválida' };
    if (cantidad > stockActual) return { success: false, mensaje: 'Stock insuficiente (hay ' + stockActual + ')' };

    var nuevoStock = Math.max(0, stockActual - cantidad);
    shInv.getRange(filaIdx + 1, 6).setValue(nuevoStock);

    var tz    = Session.getScriptTimeZone();
    var now   = new Date();
    var fecha = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var hora  = Utilities.formatDate(now, tz, 'HH:mm');

    shSal.appendRow([
      fecha, hora,
      datos.nombre || '', 'ID' + filaIdx,
      cantidad, datos.costo || 0, datos.precio || 0,
      datos.motivo || '', datos.obs || '', datos.vendedor || ''
    ]);

    return { success: true, nuevoStock };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function abonoFiado(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { success: false, mensaje: 'Hoja FIADOS no encontrada' };

    var rows = sh.getDataRange().getValues();
    var tz   = Session.getScriptTimeZone();
    var hoy  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var metodo   = datos.metodoPago || 'EFECTIVO';
    var restante = parseFloat(datos.monto) || 0;
    var cancelados = 0;

    if (restante <= 0) return { success: false, mensaje: 'Monto inválido' };

    var fiados = [];
    for (var i = 1; i < rows.length; i++) {
      var tel = String(rows[i][4] || '').replace(/\D/g,'');
      var cli = String(rows[i][3] || '').trim().toLowerCase();
      var est = String(rows[i][10] || '').toUpperCase();
      if (est === 'PAGADO') continue;
      var telBuscar = String(datos.telefono || '').replace(/\D/g,'');
      var cliBuscar = String(datos.cliente || '').trim().toLowerCase();
      if ((telBuscar && tel === telBuscar) || (!telBuscar && cli === cliBuscar)) {
        fiados.push({ fila: i + 1, total: parseFloat(rows[i][8]) || 0, fecha: String(rows[i][1] || '') });
      }
    }

    fiados.sort(function(a,b){ return a.fecha.localeCompare(b.fecha); });

    for (var j = 0; j < fiados.length; j++) {
      if (restante <= 0) break;
      var f = fiados[j];
      if (restante >= f.total) {
        sh.getRange(f.fila, 11).setValue('PAGADO');
        sh.getRange(f.fila, 12).setValue(hoy);
        sh.getRange(f.fila, 13).setValue(metodo);
        restante = Math.round((restante - f.total) * 100) / 100;
        cancelados++;
      } else {
        var nuevoTotal = Math.round((f.total - restante) * 100) / 100;
        sh.getRange(f.fila, 9).setValue(nuevoTotal);
        sh.getRange(f.fila, 8).setValue(nuevoTotal);
        var obsActual = String(sh.getRange(f.fila, 14).getValue() || '');
        var nuevaObs = (obsActual ? obsActual + ' | ' : '') +
          'Abono $' + datos.monto + ' (' + metodo + ') ' + hoy + ' — queda $' + nuevoTotal;
        sh.getRange(f.fila, 14).setValue(nuevaObs);
        restante = 0;
      }
    }

    return { success: true, cancelados, restante };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

function registrarMovimientoCaja(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var HOJA_CAJA = 'CAJA_MOVIMIENTOS';
    var sh = ss.getSheetByName(HOJA_CAJA);
    if (!sh) {
      sh = ss.insertSheet(HOJA_CAJA);
      sh.getRange(1, 1, 1, 8).setValues([['FECHA', 'HORA', 'TIPO', 'MOTIVO', 'MONTO', 'MEDIO', 'CATEGORIA', 'OBSERVACION']]);
      sh.getRange(1, 1, 1, 8).setBackground('#37474f').setFontColor('white').setFontWeight('bold');
    }
    sh.appendRow([
      datos.fecha       || '',
      datos.hora        || '',
      datos.tipo        || 'EGRESO',
      datos.motivo      || '',
      datos.monto       || 0,
      datos.medio       || 'EFECTIVO',
      datos.categoria   || '',
      datos.observacion || ''
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ========== DETALLE DE TICKET DESDE HOJA VENTAS ==========
function getDetalleTicket(ticket) {
  try {
    if (!ticket) return { ok: false, error: 'Ticket requerido' };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('Ventas');
    if (!sh) return { ok: false, error: 'Hoja Ventas no encontrada' };

    var ticketStr = String(ticket).trim();
    // Normalizar: puede venir como "0123" o "123"
    var ticketNum  = parseInt(ticketStr, 10);
    var ticketPad  = String(ticketNum).padStart(4, '0');

    var filas = sh.getDataRange().getValues();
    var items = [];
    var totalTicket = 0;
    var fechaTicket = '';
    var horaTicket  = '';
    var metodoPago  = '';
    var vendedor    = '';
    var clienteInfo = '';

    for (var i = 1; i < filas.length; i++) {
      var f = filas[i];
      var tFila = String(f[0] || '').trim();
      // Fila TOTAL del ticket
      var desc = String(f[3] || '');
      if (desc.includes('TOTAL TICKET') && (desc.includes('N° ' + ticketPad) || desc.includes('N° ' + ticketNum))) {
        fechaTicket = f[1] instanceof Date
          ? Utilities.formatDate(f[1], Session.getScriptTimeZone(), 'dd/MM/yyyy')
          : String(f[1] || '');
        horaTicket  = String(f[2] || '');
        metodoPago  = String(f[5] || '');
        vendedor    = String(f[6] || '');
        totalTicket = parseFloat(f[7]) || 0;
        // clienteInfo está dentro del metodoPago después del pipe
        if (metodoPago.includes('|')) {
          var partes = metodoPago.split('|');
          metodoPago  = partes[0].trim();
          clienteInfo = partes[1].trim();
        }
        continue;
      }
      // Filas de productos del ticket
      if (tFila !== ticketPad && tFila !== String(ticketNum)) continue;
      items.push({
        nombre:    String(f[3] || '').trim(),
        cantidad:  f[4],
        precio:    parseFloat(f[5]) || 0,
        subtotal:  parseFloat(f[6]) || 0
      });
    }

    if (!items.length && !totalTicket) return { ok: false, error: 'Ticket ' + ticketStr + ' no encontrado' };

    return {
      ok: true,
      ticket:     ticketPad,
      fecha:      fechaTicket,
      hora:       horaTicket,
      metodoPago: metodoPago,
      vendedor:   vendedor,
      cliente:    clienteInfo,
      total:      totalTicket,
      items:      items
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

function guardarFiado(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { success: false, mensaje: 'Hoja FIADOS no encontrada' };

    var now = new Date();
    var tz  = Session.getScriptTimeZone();
    var fechaStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');

    var ultimaFila = sh.getLastRow();
    var numFiado = ultimaFila;
    var idFiado = datos.idFiado || ('FIADO-' + String(numFiado).padStart(4, '0'));

    sh.appendRow([
      idFiado, fechaStr,
      datos.ticket || '', datos.cliente || '', datos.telefono || '',
      datos.descripcion || 'TICKET COMPLETO',
      datos.cantItems || 1, datos.total || 0, datos.total || 0,
      datos.vencimiento || '',
      '', '', '',
      datos.obs || ''
    ]);

    var nuevaFila = sh.getLastRow();
    sh.getRange(nuevaFila, 11).setFormula(
      '=IF(D'+nuevaFila+'="";"";IF(L'+nuevaFila+'<>"";"PAGADO";IF(AND(J'+nuevaFila+'<>"";TODAY()>J'+nuevaFila+');"VENCIDO";"PENDIENTE")))'
    );

    // Registrar / actualizar cliente automáticamente
    if (datos.cliente) {
      upsertCliente(ss, datos.cliente, datos.telefono || '');
    }

    return { success: true, idFiado };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// Inserta el cliente si no existe; si existe actualiza el teléfono si estaba vacío
function upsertCliente(ss, nombre, telefono) {
  try {
    var sh = ss.getSheetByName('Clientes');
    if (!sh) {
      // Crear hoja con encabezados si no existe
      sh = ss.insertSheet('Clientes');
      sh.getRange(1, 1, 1, 2).setValues([['Nombre', 'Telefono']]);
      sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f5e9');
      sh.setFrozenRows(1);
    }

    var nombreNorm = nombre.trim().toUpperCase();
    var telLimpio  = String(telefono).replace(/\D/g, '');
    var datos = sh.getDataRange().getValues();

    // Buscar fila existente por nombre (insensible a mayúsculas)
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0] || '').trim().toUpperCase() === nombreNorm) {
        // Ya existe — completar teléfono solo si la celda está vacía
        if (!String(datos[i][1] || '').trim() && telLimpio) {
          sh.getRange(i + 1, 2).setValue(telLimpio);
        }
        return; // nada más que hacer
      }
    }

    // No existe — agregar nueva fila
    sh.appendRow([nombre.trim(), telLimpio]);
  } catch(e) {
    console.error('upsertCliente error (no crítico):', e);
    // Fallo silencioso — no interrumpe el guardado del fiado
  }
}

function consultarFiado(telefono) {
  try {
    if (!telefono) return { ok: true, deuda: null };
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName(HOJA_FIADOS);
    if (!sh) return { ok: true, deuda: null };

    var datos = sh.getDataRange().getValues();
    var telBuscar = String(telefono).replace(/\D/g, '');

    var totalDeuda = 0;
    var pendientes = 0;
    var vencidos = 0;
    var nombreCliente = '';
    var hoy = new Date(); hoy.setHours(0,0,0,0);

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var telFila = String(fila[4] || '').replace(/\D/g, '');
      if (telFila !== telBuscar) continue;

      var estado = String(fila[10] || '').toUpperCase();
      if (estado === 'PAGADO') continue;

      var fechaVenc = fila[9] ? new Date(fila[9]) : null;
      if (fechaVenc && fechaVenc < hoy && estado !== 'PAGADO') estado = 'VENCIDO';

      var total = parseFloat(fila[8]) || 0;
      totalDeuda += total;
      pendientes++;
      if (estado === 'VENCIDO') vencidos++;
      if (!nombreCliente && fila[3]) nombreCliente = String(fila[3]);
    }

    if (pendientes === 0) return { ok: true, deuda: null };
    return { ok: true, deuda: { nombre: nombreCliente, total: totalDeuda, pendientes, vencidos } };
  } catch(e) {
    return { ok: true, deuda: null };
  }
}

// ========== LISTAR CLIENTES v2 (autocomplete fiados) ==========
// Lee primero hoja Clientes; si está vacía, extrae clientes únicos de FIADOS.
// Así siempre hay datos aunque la hoja Clientes no esté poblada aún.
function listarClientes() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var clientes = [];
    var vistos   = {}; // para deduplicar por nombre normalizado

    // ── 1. Leer hoja Clientes (fuente principal) ──
    var shCli = ss.getSheetByName('Clientes');
    if (shCli) {
      var datosCli = shCli.getDataRange().getValues();
      if (datosCli.length >= 2) {
        var hdr     = datosCli[0].map(function(h){ return String(h).toLowerCase().trim(); });
        var iNombre = hdr.findIndex(function(h){ return h.includes('nombre'); });
        var iTel    = hdr.findIndex(function(h){ return h.includes('tel') || h.includes('fono') || h.includes('phone'); });
        if (iNombre >= 0) {
          for (var i = 1; i < datosCli.length; i++) {
            var nom = String(datosCli[i][iNombre] || '').trim();
            if (!nom) continue;
            var key = nom.toUpperCase();
            if (vistos[key]) continue;
            vistos[key] = true;
            var tel = iTel >= 0 ? String(datosCli[i][iTel] || '').trim() : '';
            clientes.push({ nombre: nom, telefono: tel });
          }
        }
      }
    }

    // ── 2. Fallback: extraer clientes únicos de hoja FIADOS ──
    // (necesario si los fiados se crearon antes de la v7.5 que puebla Clientes)
    var shFia = ss.getSheetByName('FIADOS');
    if (shFia) {
      var datosFia = shFia.getDataRange().getValues();
      // Col D (idx 3) = Cliente, Col E (idx 4) = Teléfono
      for (var j = 1; j < datosFia.length; j++) {
        var fila = datosFia[j];
        if (!fila[0]) continue; // fila vacía
        var nomF = String(fila[3] || '').trim();
        if (!nomF) continue;
        var keyF = nomF.toUpperCase();
        if (vistos[keyF]) continue; // ya está desde hoja Clientes
        vistos[keyF] = true;
        var telF = String(fila[4] || '').trim();
        clientes.push({ nombre: nomF, telefono: telF });

        // Aprovechar y sincronizar a hoja Clientes para próximas consultas
        try { upsertCliente(nomF, telF); } catch(ex) {}
      }
    }

    // Ordenar alfabéticamente
    clientes.sort(function(a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });

    return { ok: true, clientes: clientes };
  } catch(e) {
    console.error('Error listarClientes v2:', e);
    return { ok: true, clientes: [] };
  }
}

// ========== VENTAS POR PRODUCTO (para panel de info POS) ==========
function getVentasProducto(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('Ventas');
    if (!sh) return { success: false, mensaje: 'Hoja Ventas no encontrada' };

    var nombre = String(datos.nombre || '').trim().toUpperCase();
    var dias   = parseInt(datos.dias) || 60;
    if (!nombre) return { success: false, mensaje: 'Nombre requerido' };

    var tz    = Session.getScriptTimeZone();
    var hoy   = new Date(); hoy.setHours(0,0,0,0);
    var desde = new Date(hoy); desde.setDate(desde.getDate() - dias);

    var filas = sh.getDataRange().getValues();
    var porDia = {}; // 'YYYY-MM-DD' -> qty

    for (var i = 1; i < filas.length; i++) {
      var f = filas[i];
      // Fila de producto: col A tiene ticket, col D tiene nombre
      if (!f[0] || String(f[3] || '').includes('TOTAL TICKET')) continue;
      var nomFila = String(f[3] || '').trim().toUpperCase();
      // Quitar sufijos de precio especial para comparar
      var nomLimpio = nomFila.replace(/\s*\[PRECIO.*?\]$/,'').replace(/\s*\(\d+\s*gr\)$/i,'').replace(/\s*\(\d+\.?\d*\s*kg\)$/i,'').trim();
      if (nomLimpio !== nombre) continue;

      var fechaFila = f[1] instanceof Date ? f[1] : new Date(f[1]);
      if (isNaN(fechaFila) || fechaFila < desde) continue;

      var fechaKey = Utilities.formatDate(fechaFila, tz, 'yyyy-MM-dd');
      var qty = parseFloat(f[4]) || 1;
      porDia[fechaKey] = (porDia[fechaKey] || 0) + qty;
    }

    // Convertir a array ordenado desc
    var resultado = Object.keys(porDia).sort().reverse().map(function(d) {
      return { fecha: d, qty: Math.round(porDia[d] * 1000) / 1000 };
    });

    var totalQty = resultado.reduce(function(s, r) { return s + r.qty; }, 0);

    return { success: true, dias: resultado, total: Math.round(totalQty * 1000) / 1000 };
  } catch(e) {
    return { success: false, mensaje: e.toString() };
  }
}

// ======================
// FIADOS v2.0 — Detalle por ticket + Pago selectivo
// ======================

// ========== LISTAR FIADOS DE UN CLIENTE CON DETALLE DE TICKETS ==========
// Busca por teléfono (o nombre) todos los fiados pendientes de un cliente
// y enriquece cada uno con el detalle real de productos desde hoja Ventas.
function listarFiadosCliente(datos) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var shFiados = ss.getSheetByName(HOJA_FIADOS);
    if (!shFiados) return { ok: false, error: 'Hoja FIADOS no encontrada' };

    var telefono = String(datos.telefono  || '').replace(/\D/g, '');
    var cliente  = String(datos.cliente   || '').trim().toUpperCase();
    if (!telefono && !cliente) return { ok: false, error: 'Se requiere teléfono o cliente' };

    var rows = shFiados.getDataRange().getValues();
    var hoy  = new Date(); hoy.setHours(0, 0, 0, 0);
    var fiados = [];

    for (var i = 1; i < rows.length; i++) {
      var fila = rows[i];
      if (!fila[0]) continue;

      var telFila = String(fila[4] || '').replace(/\D/g, '');
      var cliFila = String(fila[3] || '').trim().toUpperCase();

      var coincide = (telefono && telFila === telefono) ||
                     (!telefono && cliFila === cliente);
      if (!coincide) continue;

      var estado = String(fila[10] || '').toUpperCase();
      var vencFecha = fila[9]
        ? new Date(String(fila[9]).split('T')[0] + 'T12:00:00')
        : null;
      if (vencFecha && vencFecha < hoy && estado !== 'PAGADO') estado = 'VENCIDO';

      // Solo pendientes + vencidos
      if (estado === 'PAGADO') continue;

      var ticketNum    = String(fila[2] || '').trim();
      var detalle      = [];
      var totalDetalle = 0;

      // Enriquecer con detalle real de hoja Ventas si hay ticket
      if (ticketNum) {
        var res = getDetalleTicket(ticketNum);
        if (res.ok && res.items && res.items.length) {
          detalle      = res.items;
          totalDetalle = res.total;
        }
      }

      var totalPendiente = parseFloat(fila[8]) || 0;
      var totalOriginal  = parseFloat(fila[7]) || totalPendiente;

      fiados.push({
        idFiado:          String(fila[0]),
        fecha:            String(fila[1] || ''),
        ticket:           ticketNum,
        cliente:          String(fila[3] || ''),
        telefono:         String(fila[4] || ''),
        descripcion:      String(fila[5] || ''),
        cantItems:        fila[6] || 0,
        totalOriginal:    totalOriginal,
        totalPendiente:   totalPendiente,
        abonadoParcial:   totalOriginal > totalPendiente,
        montoAbonado:     Math.round((totalOriginal - totalPendiente) * 100) / 100,
        fechaVencimiento: fila[9] ? String(fila[9]).split('T')[0] : '',
        estado:           estado,
        observaciones:    String(fila[13] || ''),
        detalle:          detalle,
        totalDetalle:     totalDetalle
      });
    }

    // Ordenar: VENCIDO primero, luego por fecha ascendente
    fiados.sort(function(a, b) {
      if (a.estado === 'VENCIDO' && b.estado !== 'VENCIDO') return -1;
      if (b.estado === 'VENCIDO' && a.estado !== 'VENCIDO') return 1;
      return a.fecha.localeCompare(b.fecha);
    });

    var totalDeuda  = fiados.reduce(function(s, f) { return s + f.totalPendiente; }, 0);
    var hayVencidos = fiados.some(function(f) { return f.estado === 'VENCIDO'; });

    return {
      ok:          true,
      cliente:     fiados.length ? fiados[0].cliente : (datos.cliente || ''),
      telefono:    telefono,
      totalDeuda:  Math.round(totalDeuda * 100) / 100,
      cantFiados:  fiados.length,
      hayVencidos: hayVencidos,
      fiados:      fiados
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ========== PAGAR FIADOS SELECCIONADOS (total o parcial) ==========
// Recibe array de pagos: [{ idFiado, monto, parcial }]
// Actualiza FIADOS y registra ingreso en CAJA_MOVIMIENTOS.
function pagarFiadosSeleccionados(datos) {
  try {
    if (!datos.pagos || !datos.pagos.length) {
      return { ok: false, error: 'No se indicaron fiados a cobrar' };
    }

    var ss      = SpreadsheetApp.openById(SS_ID);
    var shFiado = ss.getSheetByName(HOJA_FIADOS);
    if (!shFiado) return { ok: false, error: 'Hoja FIADOS no encontrada' };

    var tz       = Session.getScriptTimeZone();
    var ahora    = new Date();
    var fechaStr = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
    var horaStr  = Utilities.formatDate(ahora, tz, 'HH:mm');
    var metodo   = String(datos.metodoPago || 'EFECTIVO').toUpperCase();

    var rows      = shFiado.getDataRange().getValues();
    var resultado = [];
    var totalCobrado = 0;

    for (var p = 0; p < datos.pagos.length; p++) {
      var pago      = datos.pagos[p];
      var idFiado   = String(pago.idFiado || '').trim();
      var monto     = parseFloat(pago.monto) || 0;
      var esParcial = !!pago.parcial;

      if (!idFiado || monto <= 0) continue;

      // Buscar fila en sheet
      var filaEncontrada = -1;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === idFiado) {
          filaEncontrada = i;
          break;
        }
      }
      if (filaEncontrada < 0) {
        resultado.push({ idFiado: idFiado, ok: false, error: 'No encontrado' });
        continue;
      }

      var totalActual = parseFloat(rows[filaEncontrada][8]) || 0;
      var montoReal   = Math.min(monto, totalActual); // nunca cobrar más de lo que debe
      var nuevo       = Math.round((totalActual - montoReal) * 100) / 100;
      var filaSh      = filaEncontrada + 1;

      // Actualizar total pendiente (col I = col 9)
      shFiado.getRange(filaSh, 9).setValue(nuevo);

      // Actualizar observaciones (col N = col 14)
      var obsActual = String(rows[filaEncontrada][13] || '');
      var tipoOp    = esParcial ? 'Abono' : 'Pago total';
      var nuevaObs  = (obsActual ? obsActual + ' | ' : '') +
        tipoOp + ' $' + montoReal + ' (' + metodo + ') ' + fechaStr + ' ' + horaStr;
      shFiado.getRange(filaSh, 14).setValue(nuevaObs);

      // Si quedó en cero: marcar PAGADO (col K=11, L=12, M=13)
      if (nuevo === 0) {
        shFiado.getRange(filaSh, 11).setValue('PAGADO');
        shFiado.getRange(filaSh, 12).setValue(fechaStr + ' ' + horaStr);
        shFiado.getRange(filaSh, 13).setValue(metodo);
      }

      totalCobrado += montoReal;
      resultado.push({
        idFiado:      idFiado,
        ok:           true,
        montoCobrado: montoReal,
        nuevoTotal:   nuevo,
        pagadoTotal:  nuevo === 0
      });
    }

    // Registrar en CAJA_MOVIMIENTOS
    if (totalCobrado > 0) {
      _registrarIngresoFiadoCaja_(ss, {
        fecha:    fechaStr,
        hora:     horaStr,
        monto:    totalCobrado,
        medio:    metodo,
        cliente:  datos.cliente || '',
        obs:      'Cobro fiado — ' + resultado.filter(function(r){ return r.ok; }).length + ' ticket(s)'
      });
    }

    return {
      ok:           true,
      totalCobrado: Math.round(totalCobrado * 100) / 100,
      detalle:      resultado
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ——— Helper: registra ingreso en CAJA_MOVIMIENTOS ———
function _registrarIngresoFiadoCaja_(ss, d) {
  try {
    var HOJA_CAJA = 'CAJA_MOVIMIENTOS';
    var sh = ss.getSheetByName(HOJA_CAJA);
    if (!sh) {
      sh = ss.insertSheet(HOJA_CAJA);
      sh.getRange(1, 1, 1, 8).setValues([[
        'FECHA','HORA','TIPO','MOTIVO','MONTO','MEDIO','CATEGORIA','OBSERVACION'
      ]]);
      sh.getRange(1, 1, 1, 8).setBackground('#37474f').setFontColor('white').setFontWeight('bold');
    }
    sh.appendRow([
      d.fecha, d.hora,
      'INGRESO',
      'COBRO FIADO — ' + (d.cliente || '').toUpperCase(),
      d.monto,
      d.medio,
      'FIADOS',
      d.obs || ''
    ]);
  } catch(e) {
    console.error('_registrarIngresoFiadoCaja_ error (no crítico):', e);
  }
}

// ========== RESET TOTAL — SÓLO PARA PRUEBAS ==========
// Borra todos los datos transaccionales y deja el inventario con stock=0
// Se activa sólo si el PIN recibido coincide con RESET_PIN
function resetearPlanilla(pin) {
  // PIN desde CONFIG_UI — si no está configurado usa 7749 como fallback
  var RESET_PIN = '7749';
  try {
    var ssCfg  = SpreadsheetApp.openById(SS_ID);
    var shCfgU = ssCfg.getSheetByName('CONFIG_UI');
    if (shCfgU) {
      var cfgUIData = shCfgU.getDataRange().getValues();
      for (var ci = 1; ci < cfgUIData.length; ci++) {
        if (String(cfgUIData[ci][0]).trim() === 'RESET_PIN') {
          var pinVal = String(cfgUIData[ci][1] || '').trim();
          if (pinVal) RESET_PIN = pinVal;
          break;
        }
      }
    }
  } catch(ePin) { /* usa fallback */ }
  if (String(pin) !== RESET_PIN) {
    return { ok: false, msg: 'PIN incorrecto' };
  }

  try {
    var ss = SpreadsheetApp.openById(SS_ID);

    // 1. Hojas de transacciones → borrar todo menos la fila de encabezado
    var hojasTrans = ['Ventas', 'FIADOS', 'SALIDAS', 'Caja',
                      'CAJA_MOVIMIENTOS', 'Historial', 'Ajuste_Rapido',
                      'Interruptores_Log', 'evento_cerveza_log',
                      'Clientes', 'Lista de compra', 'carrito_temp'];

    hojasTrans.forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (!sh) return;
      var lastRow = sh.getLastRow();
      if (lastRow > 1) {
        sh.deleteRows(2, lastRow - 1);
      }
    });

    // 2. Inventario → poner STOCK = 0 en columna F (col 6), mantener todo lo demás
    var shInv = ss.getSheetByName(HOJA_INVENTARIO);
    var lastRowInv = shInv.getLastRow();
    if (lastRowInv > 1) {
      // Columna F = stock (índice 6)
      var stockRange = shInv.getRange(2, 6, lastRowInv - 1, 1);
      var zeros = Array(lastRowInv - 1).fill([0]);
      stockRange.setValues(zeros);
    }

    // 3. config_sistema → limpiar solo filas de estado (ventas del día, caja, etc.)
    //    Dejamos las filas de configuración del sistema intactas
    var shCfg = ss.getSheetByName(HOJA_CONFIG);
    if (shCfg) {
      var cfgData = shCfg.getDataRange().getValues();
      for (var i = 1; i < cfgData.length; i++) {
        var clave = String(cfgData[i][0] || '').toUpperCase();
        // Solo resetear claves de estado, no de configuración
        if (['VENTAS_HOY', 'CAJA_HOY', 'ULTIMO_TICKET', 'TICKET_COUNTER',
             'TOTAL_HOY', 'EFECTIVO_HOY', 'TRANSFERENCIA_HOY', 'MP_HOY',
             'FIADO_HOY', 'GASTOS_HOY'].indexOf(clave) !== -1) {
          shCfg.getRange(i + 1, 2).setValue(0);
        }
      }
    }

    return { ok: true, msg: '✅ Planilla reseteada correctamente. Stock en 0, transacciones borradas.' };

  } catch(e) {
    return { ok: false, msg: '❌ Error: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// PATAGONIA SALVAJE — FUNCIONES NUEVAS
// ═══════════════════════════════════════════════════════════════

// ── 1. getCajaEgresos ────────────────────────────────────────
// Devuelve registros de EGRESO/RETIRO de CAJA_MOVIMIENTOS
// Params: { fecha: 'yyyy-MM-dd' (opcional), limite: N }
function getCajaEgresos(datos) {
  try {
    var ss   = SpreadsheetApp.openById(SS_ID);
    var sh   = ss.getSheetByName('CAJA_MOVIMIENTOS');
    if (!sh) return { ok: true, registros: [] };

    var rows  = sh.getDataRange().getValues();
    var limite = (datos && datos.limite) ? parseInt(datos.limite) : 200;
    var filtFecha = datos && datos.fecha ? String(datos.fecha).trim() : null;

    // Headers esperados: FECHA, HORA, TIPO, MOTIVO, MONTO, MEDIO, CATEGORIA, OBSERVACION
    var registros = [];
    for (var i = 1; i < rows.length; i++) {
      var r     = rows[i];
      var tipo  = String(r[2] || '').toUpperCase();
      var fecha = r[0] instanceof Date
        ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(r[0] || '').trim();

      // Solo EGRESO / RETIRO (excluye INGRESO de abono fiado)
      if (tipo !== 'EGRESO' && tipo !== 'RETIRO' && tipo !== 'GASTO') continue;

      // Filtro por fecha si vino
      if (filtFecha && fecha !== filtFecha) continue;

      registros.push({
        fecha    : fecha,
        hora     : String(r[1] || ''),
        tipo     : tipo,
        motivo   : String(r[3] || ''),
        monto    : parseFloat(r[4]) || 0,
        medio    : String(r[5] || 'EFECTIVO'),
        categoria: String(r[6] || ''),
        obs      : String(r[7] || '')
      });
    }

    // Más reciente primero
    registros.sort(function(a, b) {
      return (b.fecha + b.hora).localeCompare(a.fecha + a.hora);
    });

    if (registros.length > limite) registros = registros.slice(0, limite);

    return { ok: true, registros: registros };
  } catch(e) {
    return { ok: false, error: e.toString(), registros: [] };
  }
}

// ── 2. getSalidasInternas ────────────────────────────────────
// Devuelve registros de la hoja SALIDAS (merma, consumo, etc.)
// Params: { fecha: 'yyyy-MM-dd' (opcional), limite: N }
function getSalidasInternas(datos) {
  try {
    var ss  = SpreadsheetApp.openById(SS_ID);
    var sh  = ss.getSheetByName('SALIDAS');
    if (!sh) return { ok: true, registros: [] };

    var rows     = sh.getDataRange().getValues();
    var limite   = (datos && datos.limite) ? parseInt(datos.limite) : 200;
    var filtFecha = datos && datos.fecha ? String(datos.fecha).trim() : null;

    // Headers: Fecha, Hora, Producto, ID, Cantidad, Costo Unit., Precio Venta, Motivo, Observación, Vendedor
    var registros = [];
    for (var i = 1; i < rows.length; i++) {
      var r     = rows[i];
      var fecha = r[0] instanceof Date
        ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(r[0] || '').trim();

      if (filtFecha && fecha !== filtFecha) continue;

      registros.push({
        fecha   : fecha,
        hora    : String(r[1] || ''),
        producto: String(r[2] || ''),
        id      : String(r[3] || ''),
        cantidad: parseFloat(r[4]) || 0,
        costo   : parseFloat(r[5]) || 0,
        precio  : parseFloat(r[6]) || 0,
        motivo  : String(r[7] || '').toUpperCase(),
        obs     : String(r[8] || ''),
        vendedor: String(r[9] || '')
      });
    }

    // Más reciente primero
    registros.sort(function(a, b) {
      return (b.fecha + b.hora).localeCompare(a.fecha + a.hora);
    });

    if (registros.length > limite) registros = registros.slice(0, limite);

    return { ok: true, registros: registros };
  } catch(e) {
    return { ok: false, error: e.toString(), registros: [] };
  }
}

// ── 3. getPromoEspecial ──────────────────────────────────────
// Devuelve productos de categoría PROMOS para el flyer
// Igual a getJuevesCervecero pero filtra PROMOS en lugar de cervezas
function getPromoEspecial() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var shInv = ss.getSheetByName(HOJA_INVENTARIO);
    var shCfg = ss.getSheetByName(HOJA_CONFIG);

    if (!shInv) return { success: false, productos: [] };

    var invRows = shInv.getDataRange().getValues();
    var tz      = Session.getScriptTimeZone();
    var hoy     = new Date();
    var diaSem  = hoy.getDay(); // 0=DOM..6=SÁB → mapeamos a col LUN=1
    var colDia  = diaSem === 0 ? 7 : diaSem; // domingo → col 7

    // Hora de cierre desde config_sistema
    var horaCierre = 22;
    if (shCfg) {
      var cfgRows = shCfg.getDataRange().getValues();
      for (var c = 1; c < cfgRows.length; c++) {
        var lk = String(cfgRows[c][0] || '').toLowerCase().trim();
        if (lk === 'horario cierre local') {
          var v = cfgRows[c][colDia];
          if (v !== '' && !isNaN(parseFloat(v))) { horaCierre = parseFloat(v); break; }
        }
      }
    }

    var productos = [];
    for (var i = 1; i < invRows.length; i++) {
      var row  = invRows[i];
      var cat  = String(row[2] || '').toUpperCase().trim();
      var est  = String(row[19] || '').toUpperCase();
      var stock = parseFloat(row[5]) || 0;

      if (cat !== 'PROMOS') continue;
      if (est === '❌' || est === 'INACTIVO') continue;
      if (stock <= 0) continue;

      var precio      = parseFloat(row[1]) || 0;
      var precioFinal = parseFloat(row[10]) || precio;
      var precioAntes = parseFloat(row[11]) || 0;

      productos.push({
        nombre      : String(row[0] || ''),
        precio      : precioAntes > 0 ? precioAntes : precio,
        precioPromo  : precioFinal,
        stock       : stock,
        categoria   : cat
      });
    }

    // Ordenar por mayor ahorro absoluto
    productos.sort(function(a, b) {
      var ahorroA = (a.precio - a.precioPromo);
      var ahorroB = (b.precio - b.precioPromo);
      return ahorroB - ahorroA;
    });

    return { success: true, productos: productos, horaCierre: horaCierre };
  } catch(e) {
    return { success: false, productos: [], error: e.toString() };
  }
}

// ── 4. getHorarioLocal ───────────────────────────────────────
// Lee/escribe los horarios de apertura y cierre por día
// Params para leer: {} → devuelve { ok, horarios: [{dia, apertura, cierre}] }
// Params para escribir: { dia: 'LUN', apertura: 9, cierre: 22 }
function getHorarioLocal(datos) {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var shCfg = ss.getSheetByName(HOJA_CONFIG);
    if (!shCfg) return { ok: false, error: 'Sin hoja config_sistema' };

    var dias   = ['LUN','MAR','MIÉ','JUE','VIE','SÁB','DOM'];
    var cfgRows = shCfg.getDataRange().getValues();

    // ── ESCRITURA ──
    if (datos && datos.dia) {
      var idxDia = dias.indexOf(String(datos.dia).toUpperCase().trim());
      if (idxDia < 0) return { ok: false, error: 'Día inválido' };
      var col = idxDia + 2; // col B=2 → LUN

      // Buscar/crear filas de apertura y cierre
      var filaApertura = -1, filaCierre = -1;
      for (var i = 1; i < cfgRows.length; i++) {
        var lk = String(cfgRows[i][0] || '').toLowerCase().trim();
        if (lk === 'horario apertura local') filaApertura = i + 1;
        if (lk === 'horario cierre local')   filaCierre   = i + 1;
      }
      if (filaApertura > 0 && datos.apertura !== undefined)
        shCfg.getRange(filaApertura, col).setValue(datos.apertura);
      if (filaCierre > 0 && datos.cierre !== undefined)
        shCfg.getRange(filaCierre, col).setValue(datos.cierre);
      return { ok: true };
    }

    // ── LECTURA ──
    var horApertura = [9,9,9,9,9,9,9];
    var horCierre   = [22,22,22,22,22,22,22];

    for (var j = 1; j < cfgRows.length; j++) {
      var lk2 = String(cfgRows[j][0] || '').toLowerCase().trim();
      if (lk2 === 'horario apertura local') {
        for (var d = 0; d < 7; d++) {
          var v = cfgRows[j][d + 1];
          if (v !== '' && !isNaN(parseFloat(v))) horApertura[d] = parseFloat(v);
        }
      }
      if (lk2 === 'horario cierre local') {
        for (var d2 = 0; d2 < 7; d2++) {
          var v2 = cfgRows[j][d2 + 1];
          if (v2 !== '' && !isNaN(parseFloat(v2))) horCierre[d2] = parseFloat(v2);
        }
      }
    }

    var horarios = dias.map(function(dia, i) {
      return { dia: dia, apertura: horApertura[i], cierre: horCierre[i] };
    });

    return { ok: true, horarios: horarios };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════
// GENERAR LISTA DE COMPRA — lee inventario y llena hoja
// ══════════════════════════════════════════════════════
function actualizarListaCompra() {
  try {
    var ss    = SpreadsheetApp.openById(SS_ID);
    var shInv = ss.getSheetByName(HOJA_INVENTARIO);
    var shList = ss.getSheetByName('Lista de compra');

    if (!shInv)  return { ok: false, error: 'Sin hoja inventario' };
    if (!shList) {
      shList = ss.insertSheet('Lista de compra');
    }

    // Limpiar todo menos encabezado
    var lastRow = shList.getLastRow();
    if (lastRow > 1) shList.deleteRows(2, lastRow - 1);

    // Asegurar encabezado
    shList.getRange(1,1,1,7).setValues([['PRODUCTO','STOCK','MÍNIMO','FALTANTE','PRECIO','TOTAL','PROVEEDOR']]);
    shList.getRange(1,1,1,7).setFontWeight('bold')
      .setBackground('#1B4332').setFontColor('white');

    var invData = shInv.getDataRange().getValues();
    var tz  = Session.getScriptTimeZone();
    var hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var filas = [];

    for (var i = 1; i < invData.length; i++) {
      var row      = invData[i];
      var nombre   = String(row[0]  || '').trim();
      var precio   = parseFloat(row[1])  || 0;
      var cat      = String(row[2]  || '').trim().toUpperCase();
      var prov     = String(row[4]  || '').trim();
      var stock    = parseFloat(row[5])  || 0;
      var estado   = String(row[19] || '').toUpperCase();
      var stockMin = parseFloat(row[13]) || 0; // col N = stock_min_popup

      if (!nombre) continue;
      if (estado === '❌' || estado === 'INACTIVO') continue;
      if (stockMin <= 0) continue;          // sin mínimo definido → no incluir
      if (stock >= stockMin) continue;      // stock ok → no incluir

      var faltante = Math.max(0, stockMin - stock);
      var total    = Math.round(faltante * precio);

      filas.push([nombre, stock, stockMin, faltante, precio, total, prov]);
    }

    // Ordenar: más urgentes primero (mayor faltante relativo)
    filas.sort(function(a,b) {
      var urgA = a[1] === 0 ? 999 : (a[3]/a[2]);
      var urgB = b[1] === 0 ? 999 : (b[3]/b[2]);
      return urgB - urgA;
    });

    if (filas.length > 0) {
      shList.getRange(2, 1, filas.length, 7).setValues(filas);
      // Formato columnas de precio
      shList.getRange(2, 5, filas.length, 2).setNumberFormat('$#,##0');
      // Color rojo para stock 0
      for (var j = 0; j < filas.length; j++) {
        if (filas[j][1] === 0) {
          shList.getRange(j+2, 1, 1, 7).setBackground('#FFEBEE');
        }
      }
    }

    return { ok: true, total: filas.length };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════
// CONFIG_UI — devuelve todas las claves de la hoja CONFIG_UI
// ══════════════════════════════════════════════════════════
function getConfigUI() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sh = ss.getSheetByName('CONFIG_UI');
    if (!sh) return { ok: false, error: 'Sin hoja CONFIG_UI', config: {} };

    var data = sh.getDataRange().getValues();
    var config = {};

    for (var i = 1; i < data.length; i++) {
      var clave = String(data[i][0] || '').trim();
      var valor = data[i][1];
      if (!clave || clave.startsWith('──')) continue;
      // Convertir a string salvo números puros
      config[clave] = (valor === null || valor === undefined || valor === '')
        ? '' : valor;
    }
    // Procesar VENDEDORES_EXCLUIR como array
    if (config['VENDEDORES_EXCLUIR']) {
      config['VENDEDORES_EXCLUIR_ARR'] = String(config['VENDEDORES_EXCLUIR'])
        .split(',').map(function(v){ return v.trim().toUpperCase(); }).filter(Boolean);
    } else {
      config['VENDEDORES_EXCLUIR_ARR'] = [];
    }
    return { ok: true, config: config };
  } catch(e) {
    return { ok: false, error: e.toString(), config: {} };
  }
}

// ══════════════════════════════════════════════════════════════
// CONFIG_UI — colorear celdas con preview visual del color hex
// ══════════════════════════════════════════════════════════════
function colorearConfigUI() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('CONFIG_UI');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach((row, i) => {
    const color = row[1];
    if (typeof color === 'string' && color.match(/^#([0-9A-F]{3}){1,2}$/i)) {
      const cell = sh.getRange(i + 2, 2);
      cell.setBackground(color);
      const rgb = hexToRgb(color);
      const luminancia = (0.299 * rgb.r) + (0.587 * rgb.g) + (0.114 * rgb.b);
      cell.setFontColor(luminancia > 186 ? '#000000' : '#ffffff');
    }
  });
}

function hexToRgb(hex) {
  hex = hex.replace('#', '');
  if (hex.length === 3) {
    hex = hex.split('').map(x => x + x).join('');
  }
  const num = parseInt(hex, 16);
  return { r: (num >> 16) & 255, g: (num >> 8) & 255, b: num & 255 };
}

// Trigger automático: cada vez que editás CONFIG_UI colorea al instante
function onEdit(e) {
  try {
    const sh = e.source.getActiveSheet();
    if (sh.getName() !== 'CONFIG_UI') return;
    colorearConfigUI();
  } catch(err) {}
}
