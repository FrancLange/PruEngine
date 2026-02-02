/**
 * ==========================================================================================
 * TESTS.gs - Test e Debug v1.0.0
 * ==========================================================================================
 * Funzioni per testare il sistema e creare dati di esempio
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// TEST LAYER 0
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test Layer 0 Spam Filter
 */
function testLayer0SpamFilter() {
  var ui = SpreadsheetApp.getUi();
  
  logSistema("🧪 Test L0: Inizio");
  
  // Test email legit
  var emailLegit = {
    mittente: "ordini@fornitoretest.it",
    oggetto: "Conferma Ordine #12345",
    corpo: "Gentile cliente, confermiamo ricezione ordine. Totale: 1250 euro."
  };
  
  var resultLegit = analisiLayer0SpamFilter(emailLegit);
  
  // Test email spam
  var emailSpam = {
    mittente: "noreply@marketing-platform.com",
    oggetto: "🎉 Super Offerta! Sconto 90% Solo Oggi!!!",
    corpo: "Clicca qui per vincere un iPhone! Offerta limitata!"
  };
  
  var resultSpam = analisiLayer0SpamFilter(emailSpam);
  
  var report = "🧪 TEST LAYER 0 SPAM FILTER\n\n" +
    "━━━ EMAIL LEGIT ━━━\n" +
    "IsSpam: " + resultLegit.isSpam + "\n" +
    "Confidence: " + resultLegit.confidence + "%\n\n" +
    "━━━ EMAIL SPAM ━━━\n" +
    "IsSpam: " + resultSpam.isSpam + "\n" +
    "Confidence: " + resultSpam.confidence + "%\n" +
    "Reason: " + (resultSpam.reason || "-") + "\n\n" +
    (resultLegit.isSpam === false && resultSpam.isSpam === true ? 
     "✅ TEST SUPERATO!" : 
     "⚠️ Verifica risultati");
  
  ui.alert("Test L0", report, ui.ButtonSet.OK);
  logSistema("🧪 Test L0: Completato");
}

/**
 * Test Layer 1
 */
function testLayer1() {
  var ui = SpreadsheetApp.getUi();
  
  logSistema("🧪 Test L1: Inizio");
  
  var emailTest = {
    mittente: "ordini@fornitoretest.it",
    oggetto: "Conferma partecipazione promo + nuovo listino",
    corpo: "Confermiamo partecipazione alla promo del 15/02. Sconto 20% linea corpo. Allegato listino primavera."
  };
  
  var result = analisiLayer1(emailTest);
  
  var report = "🧪 TEST LAYER 1 (GPT)\n\n" +
    "Success: " + result.success + "\n" +
    "Tags: " + (result.tags || []).join(", ") + "\n" +
    "Sintesi: " + (result.sintesi || "").substring(0, 100) + "\n\n" +
    (result.success ? "✅ TEST SUPERATO!" : "❌ ERRORE: " + result.error);
  
  ui.alert("Test L1", report, ui.ButtonSet.OK);
  logSistema("🧪 Test L1: Completato");
}

// ═══════════════════════════════════════════════════════════════════════
// TEST ANALISI COMPLETA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test analisi singola email
 */
function testAnalisiSingola() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🧪 Test Analisi Singola',
    'Eseguirà analisi completa L1→L2→L3 sulla prima email DA_ANALIZZARE.\n\n' +
    'Tempo stimato: 10-20 secondi\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    logSistema("🧪 TEST: Avvio analisi singola");
    
    var l1Result = analizzaEmailL1({ max_email: 1 });
    var l2Result = analizzaEmailL2({ max_email: 1 });
    var l3Result = mergeAnalisiL3({});
    
    var report = '✅ Test Completato!\n\n' +
      '📊 L1: ' + l1Result.ok + ' OK, ' + l1Result.errori + ' errori\n' +
      '📊 L2: ' + l2Result.ok + ' OK, ' + l2Result.errori + ' errori\n' +
      '📊 L3: ' + l3Result.ok + ' OK, ' + l3Result.needsReview + ' review\n\n' +
      'Controlla LOG_IN per i risultati.';
    
    ui.alert('Test Completato', report, ui.ButtonSet.OK);
    logSistema("🧪 TEST: Completato con successo");
    
  } catch (error) {
    ui.alert('❌ Errore', error.toString(), ui.ButtonSet.OK);
    logSistema("❌ TEST FALLITO: " + error.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE DATI TEST
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea email test base
 */
function creaEmailTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Foglio LOG_IN non trovato");
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  var emailTest = [
    new Date(),
    "TEST-001",
    "ordini@fornitoretest.it",
    "Re: Promo Shopping Night - Conferma partecipazione",
    "Gentile Team,\n\nConfermiamo partecipazione alla Shopping Night del 15 Febbraio.\n" +
    "Sconto 20% su linea corpo.\n\nIn allegato listino primavera 2025.\n\nCordiali saluti,\nMario Rossi"
  ];
  
  // Riempi fino a STATUS
  while (emailTest.length < colStatus) {
    emailTest.push("");
  }
  emailTest.push(CONFIG.STATUS_EMAIL.DA_FILTRARE);
  
  sheet.appendRow(emailTest);
  
  SpreadsheetApp.getUi().alert(
    "✅ Email Test Creata",
    "Email TEST-001 aggiunta con STATUS = DA_FILTRARE",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  logSistema("Email test TEST-001 creata");
}

/**
 * Crea 6 email test per spam filter (3 spam + 3 legit)
 */
function creaEmailTestSpam() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Foglio LOG_IN non trovato");
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  var now = new Date();
  
  var emailsTest = [
    // SPAM
    { id: "SPAM-001", mittente: "newsletter@marketing-pro.com", 
      oggetto: "🎉 Aumenta vendite 300%!", corpo: "Software CRM rivoluzionario! GRATIS 30 giorni!" },
    { id: "SPAM-002", mittente: "info@webagency-deluxe.it", 
      oggetto: "Sito web 299€ - Offerta Flash!", corpo: "SITO COMPLETO a 299€ invece di 2000€!" },
    { id: "SPAM-003", mittente: "mario.rossi@supplier.com", 
      oggetto: "Automatic reply: Fuori ufficio", corpo: "Sono fuori ufficio dal 25/01 al 10/02." },
    // LEGIT
    { id: "LEGIT-001", mittente: "ordini@biocosmetics.it", 
      oggetto: "Conferma Ordine #2025-0142", corpo: "Confermiamo ordine. Totale: €1.245,00. Spedizione: 05/02." },
    { id: "LEGIT-002", mittente: "commerciale@naturalbeauty.it", 
      oggetto: "Re: Richiesta Listino Primavera", corpo: "In allegato listino aggiornato. Sconto 15% ordini >€500." },
    { id: "LEGIT-003", mittente: "logistica@greenlab.it", 
      oggetto: "URGENTE: Ritardo consegna #8821", corpo: "Causa sciopero, consegna posticipata a lunedì 03/02." }
  ];
  
  emailsTest.forEach(function(email) {
    var row = [now, email.id, email.mittente, email.oggetto, email.corpo];
    while (row.length < colStatus) { row.push(""); }
    row.push(CONFIG.STATUS_EMAIL.DA_FILTRARE);
    sheet.appendRow(row);
  });
  
  SpreadsheetApp.getUi().alert(
    "✅ Email Test Create",
    "Aggiunte 6 email:\n• 3 SPAM\n• 3 LEGIT\n\nTutte con STATUS = DA_FILTRARE",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  logSistema("Create 6 email test spam filter");
}

// ═══════════════════════════════════════════════════════════════════════
// DEBUG
// ═══════════════════════════════════════════════════════════════════════

/**
 * Debug: mostra stato email in LOG_IN
 */
function debugEmailDaAnalizzare() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("LOG_IN vuoto");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  var colL0 = headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_TIMESTAMP);
  var colL1 = headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_TIMESTAMP);
  var colL2 = headers.indexOf(CONFIG.COLONNE_LOG_IN.L2_TIMESTAMP);
  var colL3 = headers.indexOf(CONFIG.COLONNE_LOG_IN.L3_TIMESTAMP);
  
  var report = "📊 DEBUG EMAIL\n\n";
  report += "Totale righe: " + (data.length - 1) + "\n\n";
  report += "ID | STATUS | L0 | L1 | L2 | L3\n";
  report += "─────────────────────────────────\n";
  
  for (var i = 1; i < Math.min(data.length, 11); i++) {
    var row = data[i];
    var id = (row[colId] || "").toString().substring(0, 10);
    var status = (row[colStatus] || "").toString().substring(0, 12);
    var l0 = row[colL0] ? "✓" : "-";
    var l1 = row[colL1] ? "✓" : "-";
    var l2 = row[colL2] ? "✓" : "-";
    var l3 = row[colL3] ? "✓" : "-";
    
    report += id + " | " + status + " | " + l0 + " | " + l1 + " | " + l2 + " | " + l3 + "\n";
  }
  
  SpreadsheetApp.getUi().alert("Debug", report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Test manuale spam su riga 2
 */
function testSpamManuale() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colMittente = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colOggetto = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  var colCorpo = headers.indexOf(CONFIG.COLONNE_LOG_IN.CORPO);
  
  var emailTest = {
    mittente: row[colMittente],
    oggetto: row[colOggetto],
    corpo: row[colCorpo]
  };
  
  var result = analisiLayer0SpamFilter(emailTest);
  
  var msg = "🧪 TEST LAYER 0\n\n" +
    "Email: " + (emailTest.oggetto || "").substring(0, 40) + "...\n\n" +
    "IsSpam: " + (result.isSpam ? "🚫 SI" : "✅ NO") + "\n" +
    "Confidence: " + result.confidence + "%\n" +
    "Reason: " + (result.reason || "-");
  
  SpreadsheetApp.getUi().alert("Test", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}