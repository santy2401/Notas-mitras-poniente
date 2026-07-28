// ============================================
// GOOGLE APPS SCRIPT - FOLIO AUTOMÁTICO
// Sheet ID: 1UBCTp258RzFhDSqTrETWyE4o_QDJMgQ-Mydvk8SNQKk
// ============================================

const SHEET_ID = "1UBCTp258RzFhDSqTrETWyE4o_QDJMgQ-Mydvk8SNQKk";
const SHEET_NAME = "Entrada"; // Tu pestaña donde guardas los datos

function doPost(e) {
  try {
    // Obtener la hoja específica
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`La pestaña "${SHEET_NAME}" no existe. Verifica el nombre.`);
    }
    
    // Parsear los datos JSON del formulario
    const data = JSON.parse(e.postData.contents);
    
    // GENERAR FOLIO AUTOMÁTICAMENTE
    // Obtener el último número de folio y sumarle 1
    const allRows = sheet.getDataRange().getValues();
    let nextFolio = 1;
    
    if (allRows.length > 1) {
      // Buscar el folio más alto en la columna B (Folio)
      for (let i = 1; i < allRows.length; i++) {
        const folioValue = allRows[i][1]; // Columna B (índice 1)
        if (folioValue && !isNaN(folioValue)) {
          nextFolio = Math.max(nextFolio, parseInt(folioValue) + 1);
        }
      }
    }
    
    // Crear la fila con los datos MAPEADOS A TUS COLUMNAS:
    // A: ID (secuencial)
    // B: Folio (número manual o automático)
    // C: Domicilio
    // D: Cantidad
    // E: Hora Inicio
    // F: Hora Final
    // G: Tiempo Entrega
    // H: Nota
    // I: Longitud (GPS)
    // J: Latitud (GPS)
    // + Columnas adicionales para referencia
    
    const row = [
      allRows.length, // Columna A: ID (número de fila)
      nextFolio, // Columna B: Folio (automático)
      (data.direccion || '') + (data.ciudad ? ', ' + data.ciudad : ''), // Columna C: Domicilio
      (data.saco1 ? 1 : 0) + (data.saco2 ? 1 : 0), // Columna D: Cantidad de trajes
      data.horarioEntrega ? data.horarioEntrega.split('-')[0].trim() : '', // Columna E: Hora inicio
      data.horarioEntrega ? data.horarioEntrega.split('-')[1].trim() : '', // Columna F: Hora final
      5, // Columna G: Tiempo estimado (5 horas)
      'Renta de ' + ((data.saco1 ? 1 : 0) + (data.saco2 ? 1 : 0)) + ' traje(s) - ' + (data.nombre || ''), // Columna H: Nota
      data.tipoEntrega === 'domicilio' ? (data.coordenadas ? data.coordenadas.split(',')[1].trim() : '') : '', // Columna I: Longitud
      data.tipoEntrega === 'domicilio' ? (data.coordenadas ? data.coordenadas.split(',')[0].trim() : '') : '', // Columna J: Latitud
      // Columnas adicionales para referencia
      data.nombre || '', // Nombre del cliente
      data.telefono || '', // Teléfono
      data.tipoEntrega || '', // Tipo de entrega
      data.fechaEvento || '', // Fecha del evento
      data.fechaEntrega || '', // Fecha de entrega
      data.fechaDevolucion || '', // Fecha de devolución
      data.importe || '', // Importe
      data.anticipo || '', // Anticipo
      data.saldo || '', // Saldo
      new Date().toLocaleString('es-MX') // Timestamp
    ];
    
    // Insertar la fila en la hoja
    sheet.appendRow(row);
    
    // Responder con éxito
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true, 
        message: 'Datos guardados correctamente',
        folio: nextFolio
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Responder con error detallado
    Logger.log('Error: ' + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false, 
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// FUNCIÓN AUXILIAR: Ver el último folio
// Ejecuta esto desde Apps Script para probar
// ============================================
/*
function getLastFolio() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  const allRows = sheet.getDataRange().getValues();
  
  Logger.log('Total de filas: ' + allRows.length);
  Logger.log('Última fila: ' + JSON.stringify(allRows[allRows.length - 1]));
}
*/