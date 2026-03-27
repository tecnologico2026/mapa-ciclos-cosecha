// ===== MAPA POLINIZACION - GOOGLE APPS SCRIPT API =====
// Hojas: POLINIZACION, CIERRES
// CIERRES columns: LOTE | TIPO | FECHA | SUPERVISOR | LABOR | REGISTRADO

function doGet(e) {
  var action = e.parameter.action || 'all';
  var result = {};

  if (action === 'polinizacion') {
    result = getLaborData('POLINIZACION');
  } else if (action === 'cierres') {
    result = getCierres();
  } else if (action === 'all') {
    result = {
      polinizacion: getLaborData('POLINIZACION'),
      cierres: getCierres()
    };
  } else if (action === 'registro') {
    var lote = e.parameter.lote || '';
    var tipo = e.parameter.tipo || 'CIERRE';
    var fecha = e.parameter.fecha || '';
    var labor = e.parameter.labor || 'POLINIZACION';
    var supervisor = e.parameter.supervisor || 'SUPERVISOR';
    result = addRegistro(lote, tipo, fecha, supervisor, labor);
  } else if (action === 'cierre') {
    // Backward compatible
    var lote = e.parameter.lote || '';
    var fecha = e.parameter.fecha || '';
    var labor = e.parameter.labor || 'POLINIZACION';
    var supervisor = e.parameter.supervisor || 'SUPERVISOR';
    result = addRegistro(lote, 'CIERRE', fecha, supervisor, labor);
  } else if (action === 'delete_registro') {
    var lote = e.parameter.lote || '';
    var tipo = e.parameter.tipo || '';
    var fecha = e.parameter.fecha || '';
    result = deleteRegistro(lote, tipo, fecha);
  } else if (action === 'delete_cierre') {
    // Backward compatible
    var lote = e.parameter.lote || '';
    var fecha = e.parameter.fecha || '';
    result = deleteRegistro(lote, 'CIERRE', fecha);
  } else if (action === 'rendimientos') {
    var desde = e.parameter.desde || '';
    var hasta = e.parameter.hasta || '';
    result = getRendimientosPolinizacion(desde, hasta);
  } else if (action === 'rendimientos_detalle') {
    var fecha = e.parameter.fecha || '';
    result = getRendimientosDetalle(fecha);
  } else if (action === 'db_explore') {
    var tabla = e.parameter.tabla || 'Ejecucion_Polinizacion';
    result = dbExplore(tabla);
  } else if (action === 'consumo_historial') {
    var desde = e.parameter.desde || '';
    var hasta = e.parameter.hasta || '';
    var ruta = e.parameter.ruta || '';
    result = getConsumoHistorial(desde, hasta, ruta);
  } else if (action === 'consumo_por_ruta') {
    var desde = e.parameter.desde || '';
    var hasta = e.parameter.hasta || '';
    result = getConsumoPorRuta(desde, hasta);
  } else if (action === 'consumo_detalle') {
    var fecha = e.parameter.fecha || '';
    result = getConsumoDetalle(fecha);
  } else if (action === 'carga_operarios') {
    var desde = e.parameter.desde || '';
    var hasta = e.parameter.hasta || '';
    var ruta = e.parameter.ruta || '';
    result = getCargaOperarios(desde, hasta, ruta);
  } else if (action === 'sheet_info') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    result = { name: ss.getName(), url: ss.getUrl(), id: ss.getId() };
  } else if (action === 'list_tables') {
    result = dbQuery("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'mty-palmas-db'");
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action || '';
  var result = {};

  if (action === 'registro') {
    result = addRegistro(data.lote, data.tipo || 'CIERRE', data.fecha, data.supervisor || 'SUPERVISOR', data.labor || 'POLINIZACION');
  } else if (action === 'cierre') {
    result = addRegistro(data.lote, 'CIERRE', data.fecha, data.supervisor || 'SUPERVISOR', data.labor || 'POLINIZACION');
  } else if (action === 'delete_registro') {
    result = deleteRegistro(data.lote, data.tipo || '', data.fecha);
  } else if (action === 'delete_cierre') {
    result = deleteRegistro(data.lote, 'CIERRE', data.fecha);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLaborData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { lots: [], error: 'Hoja ' + sheetName + ' no encontrada' };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var lots = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var lot = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j].toString().toLowerCase().replace(/ /g, '_');
      var val = row[j];
      if (val === '') val = null;
      lot[key] = val;
    }
    // Ensure numeric types
    if (lot.dias_ciclo !== null) lot.dias_ciclo = Number(lot.dias_ciclo);
    if (lot.has !== null) lot.has = Number(lot.has);
    if (lot.palmas !== null) lot.palmas = Number(lot.palmas);
    if (lot.siembra !== null) lot.siembra = Number(lot.siembra);
    lots.push(lot);
  }

  return { lots: lots, updated: new Date().toISOString() };
}

function getCierres() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CIERRES');
  if (!sheet) return { cierres: [] };

  var data = sheet.getDataRange().getValues();
  var cierres = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    // New format: LOTE | TIPO | FECHA | SUPERVISOR | LABOR | REGISTRADO
    var tipo = row[1] ? row[1].toString() : '';
    var isNewFormat = (tipo === 'INGRESO' || tipo === 'CIERRE');

    var fecha = '';
    var supervisor = '';
    var labor = '';
    var registrado = '';

    if (isNewFormat) {
      // New format
      try {
        fecha = Utilities.formatDate(new Date(row[2]), 'America/Bogota', 'yyyy-MM-dd');
      } catch(e) {
        fecha = row[2] ? row[2].toString() : '';
      }
      supervisor = row[3] ? row[3].toString() : '';
      labor = row[4] ? row[4].toString() : '';
      if (row[5]) {
        try {
          registrado = Utilities.formatDate(new Date(row[5]), 'America/Bogota', 'yyyy-MM-dd\'T\'HH:mm:ss');
        } catch(e2) {
          registrado = row[5].toString();
        }
      }
    } else {
      // Old format: LOTE | FECHA | SUPERVISOR | LABOR | REGISTRADO
      tipo = 'CIERRE';
      try {
        fecha = Utilities.formatDate(new Date(row[1]), 'America/Bogota', 'yyyy-MM-dd');
      } catch(e) {
        fecha = row[1] ? row[1].toString() : '';
      }
      supervisor = row[2] ? row[2].toString() : '';
      labor = row[3] ? row[3].toString() : '';
      if (row[4]) {
        try {
          registrado = Utilities.formatDate(new Date(row[4]), 'America/Bogota', 'yyyy-MM-dd\'T\'HH:mm:ss');
        } catch(e2) {
          registrado = row[4].toString();
        }
      }
    }

    cierres.push({
      lote: row[0].toString(),
      tipo: tipo,
      fecha: fecha,
      supervisor: supervisor,
      labor: labor,
      registrado: registrado
    });
  }

  return { cierres: cierres };
}

function addRegistro(lote, tipo, fecha, supervisor, labor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CIERRES');
  var registrado = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
  var newRow = sheet.getLastRow() + 1;
  // New format: LOTE | TIPO | FECHA | SUPERVISOR | LABOR | REGISTRADO
  sheet.getRange(newRow, 1).setNumberFormat('@').setValue(lote);
  sheet.getRange(newRow, 2).setValue(tipo);
  sheet.getRange(newRow, 3).setValue(fecha);
  sheet.getRange(newRow, 4).setValue(supervisor);
  sheet.getRange(newRow, 5).setValue(labor);
  sheet.getRange(newRow, 6).setValue(registrado);

  // Update labor sheet - recalculate dias_ciclo based on tipo
  var laborSheet = ss.getSheetByName(labor);
  if (laborSheet) {
    var data = laborSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var regDate = new Date(fecha + 'T00:00:00');
    var diffDays = Math.round((today - regDate) / 86400000 * 10) / 10;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === lote) {
        // dias_ciclo = today - fecha_cierre (only update on CIERRE)
        if (tipo === 'CIERRE') {
          laborSheet.getRange(i + 1, 3).setValue(diffDays);
        }
        break;
      }
    }
  }

  return { success: true, lote: lote, tipo: tipo, fecha: fecha };
}

function deleteRegistro(lote, tipo, fecha) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CIERRES');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var rowLote = data[i][0] ? data[i][0].toString() : '';
    var rowTipo = data[i][1] ? data[i][1].toString() : '';
    var rowFecha = '';

    // Detect format
    var isNewFmt = (rowTipo === 'INGRESO' || rowTipo === 'CIERRE');
    var fechaCol = isNewFmt ? 2 : 1;

    try {
      rowFecha = Utilities.formatDate(new Date(data[i][fechaCol]), 'America/Bogota', 'yyyy-MM-dd');
    } catch(e) {
      rowFecha = data[i][fechaCol] ? data[i][fechaCol].toString() : '';
    }

    if (!isNewFmt) rowTipo = 'CIERRE';

    if (rowLote === lote && rowFecha === fecha && (tipo === '' || rowTipo === tipo)) {
      sheet.deleteRow(i + 1);
      return { success: true, deleted: lote + ' ' + rowTipo + ' ' + fecha };
    }
  }

  return { success: false, error: 'Registro no encontrado' };
}

// ===== CLOUD SQL - SOLO LECTURA =====
// CREDENCIALES EN APPS SCRIPT SOLAMENTE - NO SUBIR A GITHUB
var DB_URL = 'jdbc:google:mysql://INSTANCE/DATABASE';
var DB_USER = 'USER';
var DB_PASS = 'PASSWORD';

function dbQuery(sql) {
  var conn = Jdbc.getCloudSqlConnection(DB_URL, DB_USER, DB_PASS);
  var stmt = conn.createStatement();
  var rs = stmt.executeQuery(sql);
  var meta = rs.getMetaData();
  var cols = meta.getColumnCount();
  var headers = [];
  for (var i = 1; i <= cols; i++) headers.push(meta.getColumnName(i));
  var rows = [];
  while (rs.next()) {
    var row = {};
    for (var i = 1; i <= cols; i++) {
      row[headers[i-1]] = rs.getString(i);
    }
    rows.push(row);
  }
  rs.close();
  stmt.close();
  conn.close();
  return { headers: headers, rows: rows };
}

function dbExplore(tabla) {
  try {
    // Whitelist allowed table names to prevent SQL injection
    var allowed = ['Ejecucion_Polinizacion', 'Ejecucion_Cosecha', 'Ejecucion_Plateo', 'Ejecucion_Poda', 'Consumo_Hormona', 'Carga_CT', 'Empleado'];
    if (allowed.indexOf(tabla) === -1) {
      return { error: 'Tabla no permitida: ' + tabla };
    }
    var result = dbQuery('SELECT * FROM ' + tabla + ' LIMIT 5');
    return { tabla: tabla, headers: result.headers, sample: result.rows };
  } catch(e) {
    return { error: e.message };
  }
}

function getRendimientosDetalle(fecha) {
  try {
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!fecha || !dateRegex.test(fecha)) {
      return { error: 'Fecha requerida en formato yyyy-MM-dd' };
    }
    var result = dbQuery(
      'SELECT ep.Ruta as ruta, ep.Lote as lote, ep.id_empleado, ' +
      'ep.`Flores Totales` as flores_totales, ' +
      'CAST(ep.area_total AS DECIMAL(10,2)) as area_total, ' +
      'ep.fecha ' +
      'FROM Ejecucion_Polinizacion ep ' +
      'WHERE ep.fecha = "' + fecha + '" ' +
      'AND ep.Ruta IS NOT NULL AND ep.Ruta != "" ' +
      'ORDER BY ep.Ruta, ep.Lote'
    );
    return { detalle: result.rows, fecha: fecha };
  } catch(e) {
    return { error: e.message };
  }
}

function getRendimientosPolinizacion(desde, hasta) {
  try {
    var where = 'WHERE ep.Ruta IS NOT NULL AND ep.Ruta != ""';
    // Sanitize date parameters - only allow yyyy-MM-dd format
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (desde && dateRegex.test(desde)) where += ' AND ep.fecha >= "' + desde + '"';
    if (hasta && dateRegex.test(hasta)) where += ' AND ep.fecha <= "' + hasta + '"';
    var result = dbQuery(
      'SELECT ep.Ruta as ruta, ' +
      'COUNT(*) as registros, ' +
      'SUM(CAST(ep.`Flores Totales` AS SIGNED)) as flores_totales, ' +
      'SUM(CAST(ep.area_total AS DECIMAL(10,2))) as area_total, ' +
      'COUNT(DISTINCT ep.id_empleado) as empleados, ' +
      'AVG(CAST(ep.`Flores Totales` AS SIGNED)) as prom_flores, ' +
      'AVG(CAST(ep.area_total AS DECIMAL(10,2))) as prom_area ' +
      'FROM Ejecucion_Polinizacion ep ' +
      where + ' ' +
      'GROUP BY ep.Ruta ' +
      'ORDER BY ep.Ruta'
    );
    return { rendimientos: result.rows };
  } catch(e) {
    return { error: e.message };
  }
}

// ===== CONSUMO DE INSUMOS =====

function getConsumoHistorial(desde, hasta, ruta) {
  try {
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    var where = 'WHERE 1=1';
    if (desde && dateRegex.test(desde)) where += ' AND fecha >= "' + desde + '"';
    if (hasta && dateRegex.test(hasta)) where += ' AND fecha <= "' + hasta + '"';
    if (ruta) {
      var cleanRuta = ruta.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ %._-]/g, '');
      where += ' AND ruta = "' + cleanRuta + '"';
    }
    var result = dbQuery(
      'SELECT fecha, ruta, ' +
      'SUM(hormona_entregada) as entregada, ' +
      'SUM(hormona_devuelta) as devuelta, ' +
      'SUM(hormona_entregada - hormona_devuelta) as consumo_neto, ' +
      'SUM(flores_totales) as flores_totales, ' +
      'COUNT(*) as registros ' +
      'FROM Consumo_Hormona ' +
      where + ' ' +
      'GROUP BY fecha, ruta ' +
      'ORDER BY fecha DESC, ruta'
    );
    var rutas = dbQuery('SELECT DISTINCT ruta FROM Consumo_Hormona WHERE ruta IS NOT NULL AND ruta != "" ORDER BY ruta');
    return { historial: result.rows, rutas: rutas.rows.map(function(r) { return r.ruta; }) };
  } catch(e) {
    return { error: e.message };
  }
}

function getConsumoPorRuta(desde, hasta) {
  try {
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    var where = 'WHERE 1=1';
    if (desde && dateRegex.test(desde)) where += ' AND fecha >= "' + desde + '"';
    if (hasta && dateRegex.test(hasta)) where += ' AND fecha <= "' + hasta + '"';
    var result = dbQuery(
      'SELECT ruta, ' +
      'SUM(hormona_entregada) as entregado, ' +
      'SUM(hormona_devuelta) as devuelto, ' +
      'SUM(hormona_entregada - hormona_devuelta) as gastado, ' +
      'SUM(flores_totales) as flores_totales, ' +
      'ROUND(SUM(hormona_devuelta) / NULLIF(SUM(hormona_entregada), 0) * 100, 1) as pct_devolucion, ' +
      'COUNT(DISTINCT fecha) as dias, ' +
      'COUNT(DISTINCT id_empleado) as empleados ' +
      'FROM Consumo_Hormona ' +
      where + ' ' +
      'GROUP BY ruta ' +
      'ORDER BY ruta'
    );
    return { consumo_ruta: result.rows };
  } catch(e) {
    return { error: e.message };
  }
}

function getConsumoDetalle(fecha) {
  try {
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!fecha || !dateRegex.test(fecha)) {
      return { error: 'Fecha requerida en formato yyyy-MM-dd' };
    }
    var result = dbQuery(
      'SELECT ruta, id_empleado, auxiliar_responsable, ' +
      'hormona_entregada, hormona_devuelta, ' +
      '(hormona_entregada - hormona_devuelta) as consumo_neto, ' +
      'flores_totales, PROMEDIO_HORMONA ' +
      'FROM Consumo_Hormona ' +
      'WHERE fecha = "' + fecha + '" ' +
      'ORDER BY ruta, id_empleado'
    );
    return { detalle: result.rows, fecha: fecha };
  } catch(e) {
    return { error: e.message };
  }
}

// ===== CARGA OPERARIOS (desde Carga_CT) =====

function getCargaOperarios(desde, hasta, ruta) {
  try {
    var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    var whereCT = 'WHERE 1=1';
    var whereEP = 'WHERE ep.Ruta IS NOT NULL AND ep.Ruta != ""';
    if (desde && dateRegex.test(desde)) {
      whereCT += ' AND c.fecha >= "' + desde + '"';
      whereEP += ' AND ep.fecha >= "' + desde + '"';
    }
    if (hasta && dateRegex.test(hasta)) {
      whereCT += ' AND c.fecha <= "' + hasta + '"';
      whereEP += ' AND ep.fecha <= "' + hasta + '"';
    }
    if (ruta) {
      var cleanRuta = ruta.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ %._-]/g, '');
      whereCT += ' AND c.ruta = "' + cleanRuta + '"';
      whereEP += ' AND ep.Ruta = "' + cleanRuta + '"';
    }

    // Flores from Carga_CT with JOIN to Empleado for name
    var flores = dbQuery(
      'SELECT c.id_empleado, e.nombre as nombre, c.ruta, c.fecha, ' +
      'SUM(CAST(c.flores_1ra_visita AS SIGNED) + CAST(c.flores_2da_visita AS SIGNED) + CAST(c.flores_3ra_visita AS SIGNED)) as flores_totales ' +
      'FROM Carga_CT c ' +
      'LEFT JOIN Empleado e ON c.id_empleado = e.id_empleado ' +
      whereCT + ' ' +
      'GROUP BY c.id_empleado, e.nombre, c.ruta, c.fecha ' +
      'ORDER BY c.id_empleado, c.fecha'
    );

    // Area from Ejecucion_Polinizacion
    var areas = dbQuery(
      'SELECT ep.id_empleado, ep.Ruta, ep.fecha, ' +
      'SUM(CAST(ep.area_total AS DECIMAL(10,2))) as area_total ' +
      'FROM Ejecucion_Polinizacion ep ' +
      whereEP + ' ' +
      'GROUP BY ep.id_empleado, ep.Ruta, ep.fecha ' +
      'ORDER BY ep.id_empleado, ep.fecha'
    );

    // Get available rutas for the filter
    var rutas = dbQuery(
      'SELECT DISTINCT ruta FROM Carga_CT ' +
      'WHERE ruta IS NOT NULL AND ruta != "" ORDER BY ruta'
    );

    return {
      flores: flores.rows,
      areas: areas.rows,
      rutas: rutas.rows.map(function(r) { return r.ruta; })
    };
  } catch(e) {
    return { error: e.message };
  }
}
