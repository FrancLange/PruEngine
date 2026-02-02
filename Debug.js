function testSpamManuale() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LOG_IN");
  
  // Leggi riga 2 (la tua email test)
  var row = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var emailTest = {
    mittente: row[2],  // Colonna C
    oggetto: row[3],   // Colonna D
    corpo: row[4]      // Colonna E
  };
  
  Logger.log("Test email: " + emailTest.oggetto);
  
  // Chiama Layer 0
  var result = analisiLayer0SpamFilter(emailTest);
  
  Logger.log("Risultato: " + JSON.stringify(result));
  
  // Alert risultato
  var msg = "🧪 TEST LAYER 0\n\n" +
    "Email: " + emailTest.oggetto.substring(0, 50) + "...\n\n" +
    "━━━ RISULTATO ━━━\n" +
    "IsSpam: " + (result.isSpam ? "🚫 SI" : "✅ NO") + "\n" +
    "Confidence: " + result.confidence + "%\n" +
    "Reason: " + (result.reason || "-") + "\n" +
    "Modello: " + result.modello + "\n\n" +
    (result.isSpam ? "✅ SPAM RILEVATO!" : "⚠️ Non rilevato come spam");
  
  SpreadsheetApp.getUi().alert("Test Completato", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
