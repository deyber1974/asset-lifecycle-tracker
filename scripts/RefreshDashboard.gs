// ============================================================
// Asset Lifecycle Tracker — Refresh Button Script
//
// SETUP (una sola vez, ~90 segundos):
//   1. Abrir el spreadsheet
//   2. Click: Extensiones → Apps Script
//   3. Borrar el código existente y pegar este archivo completo
//   4. Click Guardar (ícono diskette)
//   5. Cerrar la pestaña y reabrir el spreadsheet
//   → Aparecerá el menú "Asset Tracker" arriba
//
// OPCIONAL — Botón visual:
//   Insertar → Dibujo → dibujá un rectángulo, escribí "🔄 REFRESH"
//   Click derecho en el dibujo → Asignar secuencia de comandos → refreshDashboard
// ============================================================

var DASHBOARD_SHEET_NAME = "Dashboard";
var TIMESTAMP_CELL       = "M3";
var WEBHOOK_URL          = "";   // Opcional: pegar URL de webhook n8n aquí

function refreshDashboard() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Dashboard' no encontrado.");
    return;
  }

  var now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  );

  // Actualizar timestamp
  var cell = sheet.getRange(TIMESTAMP_CELL);
  cell.setValue("🔄  Actualizado: " + now);

  // Llamar webhook opcional (n8n)
  if (WEBHOOK_URL !== "") {
    try {
      var response = UrlFetchApp.fetch(WEBHOOK_URL, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ event: "refresh", timestamp: now }),
        muteHttpExceptions: true
      });
      var code = response.getResponseCode();
      cell.setNote("Webhook status: " + code + " — " + now);
    } catch (e) {
      cell.setNote("Webhook error: " + e.message);
    }
  }

  SpreadsheetApp.getUi().alert("✅ Dashboard actualizado\n" + now);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Asset Tracker")
    .addItem("🔄 Refresh Dashboard", "refreshDashboard")
    .addSeparator()
    .addItem("ℹ️ Acerca de", "showAbout")
    .addToUi();
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    "Asset Lifecycle Tracker\n\n" +
    "Detecta inconsistencias en activos tecnológicos\n" +
    "usando análisis automatizado + IA (Claude).\n\n" +
    "Para re-ejecutar el análisis completo,\n" +
    "correr: python3 scripts/diagnosis.py"
  );
}
