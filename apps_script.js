// ===== MAPA CICLOS - GOOGLE APPS SCRIPT API =====
// Hojas: POLINIZACION, COSECHA, CIERRES

function doGet(e) {
  var action = e.parameter.action || 'all';
  var result = {};

  if (action === 'polinizacion') {
    result = getLaborData('POLINIZACION');
  } else if (action === 'cosecha') {
    result = getLaborData('COSECHA');
  } else if (action === 'cierres') {
    result = getCierres();
  } else if (action === 'all') {
    result = {
      polinizacion: getLaborData('POLINIZACION'),
      cosecha: getLaborData('COSECHA'),
      cierres: getCierres()
    };
  } else if (action === 'cierre') {
    var lote = e.parameter.lote || '';
    var fecha = e.parameter.fecha || '';
    var labor = e.parameter.labor || 'POLINIZACION';
    var supervisor = e.parameter.supervisor || 'SUPERVISOR';
    result = addCierre(lote, fecha, supervisor, labor);
  } else if (action === 'delete_cierre') {
    var lote = e.parameter.lote || '';
    var fecha = e.parameter.fecha || '';
    result = deleteCierre(lote, fecha);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action || '';
  var result = {};

  if (action === 'cierre') {
    result = addCierre(data.lote, data.fecha, data.supervisor || 'SUPERVISOR', data.labor || 'POLINIZACION');
  } else if (action === 'delete_cierre') {
    result = deleteCierre(data.lote, data.fecha);
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
    var fecha = '';
    try {
      fecha = Utilities.formatDate(new Date(row[1]), 'America/Bogota', 'yyyy-MM-dd');
    } catch(e) {
      fecha = row[1].toString();
    }
    var registrado = '';
    if (row[4]) {
      try {
        registrado = Utilities.formatDate(new Date(row[4]), 'America/Bogota', 'yyyy-MM-dd\'T\'HH:mm:ss');
      } catch(e2) {
        registrado = row[4].toString();
      }
    }
    cierres.push({
      lote: row[0].toString(),
      fecha: fecha,
      supervisor: row[2] ? row[2].toString() : '',
      labor: row[3] ? row[3].toString() : '',
      registrado: registrado
    });
  }

  return { cierres: cierres };
}

function addCierre(lote, fecha, supervisor, labor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CIERRES');
  var registrado = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setNumberFormat('@').setValue(lote);
  sheet.getRange(newRow, 2).setValue(fecha);
  sheet.getRange(newRow, 3).setValue(supervisor);
  sheet.getRange(newRow, 4).setValue(labor);
  sheet.getRange(newRow, 5).setValue(registrado);

  // Update labor sheet - recalculate dias_ciclo
  var laborSheet = ss.getSheetByName(labor);
  if (laborSheet) {
    var data = laborSheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var cierreDate = new Date(fecha + 'T00:00:00');
    var diffDays = Math.round((today - cierreDate) / 86400000 * 10) / 10;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === lote) {
        laborSheet.getRange(i + 1, 3).setValue(diffDays);
        break;
      }
    }
  }

  return { success: true, lote: lote, fecha: fecha };
}

function deleteCierre(lote, fecha) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CIERRES');
  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var rowLote = data[i][0] ? data[i][0].toString() : '';
    var rowFecha = '';
    try {
      rowFecha = Utilities.formatDate(new Date(data[i][1]), 'America/Bogota', 'yyyy-MM-dd');
    } catch(e) {
      rowFecha = data[i][1] ? data[i][1].toString() : '';
    }
    if (rowLote === lote && rowFecha === fecha) {
      sheet.deleteRow(i + 1);
      return { success: true, deleted: lote + ' ' + fecha };
    }
  }

  return { success: false, error: 'Cierre no encontrado' };
}
