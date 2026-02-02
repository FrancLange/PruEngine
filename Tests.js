/**
 * ==========================================================================================
 * TESTS.gs - Test e Debug v1.1.0
 * ==========================================================================================
 * Funzioni per testare il sistema e creare dati di esempio
 * 
 * CHANGELOG v1.1.0:
 * - AGGIUNTO: Test L0 da AiLayer.js
 * - AGGIUNTO: testLayer0SuFoglio()
 * - Riorganizzazione test per layer
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// TEST LAYER 0 - SPAM FILTER
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test Layer 0 Spam Filter con email mock
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
    "IsSpam: " + (resultLegit.isSpam ? "🚫 SI" : "✅ NO") + "\n" +
    "Confidence: " + resultLegit.confidence + "%\n" +
    "Modello: " + resultLegit.modello + "\n\n" +
    "━━━ EMAIL SPAM ━━━\n" +
    "IsSpam: " + (resultSpam.isSpam ? "🚫 SI" : "✅ NO") + "\n" +
    "Confidence: " + resultSpam.confidence + "%\n" +
    "Reason: " + (resultSpam.reason || "-") + "\n" +
    "Modello: " + resultSpam.modello + "\n\n" +
    (resultLegit.isSpam === false && resultSpam.isSpam === true ? 
     "✅ TEST SUPERATO!" : 
     "⚠️ Verifica risultati manualmente");
  
  ui.alert("Test L0", report, ui.ButtonSet.OK);
  logSistema("🧪 Test L0: Completato");
}

/**
 * Test Layer 0 su email reali nel foglio
 */
function testLayer0SuFoglio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Foglio LOG_IN non trovato");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  data.shift();
  
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colMittente = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colOggetto = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  var colCorpo = headers.indexOf(CONFIG.COLONNE_LOG_IN.CORPO);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  var emailDaFiltrare = [];
  
  // Carica email DA_FILTRARE
  for (var i = 0; i < data.length; i++) {
    if (data[i][colStatus] === CONFIG.STATUS_EMAIL.DA_FILTRARE) {
      emailDaFiltrare.push({
        rowNum: i + 2,
        id: data[i][colId],
        mittente: data[i][colMittente],
        oggetto: data[i][colOggetto],
        corpo: data[i][colCorpo]
      });
    }
  }
  
  if (emailDaFiltrare.length === 0) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Nessuna Email da Filtrare",
      "Non ci sono email con STATUS = DA_FILTRARE.\n\n" +
      "Esegui prima: Menu > Test > Crea Email Test Spam",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  logSistema("🛡️ TEST L0: Inizio filtro su " + emailDaFiltrare.length + " email");
  
  var risultati = [];
  
  emailDaFiltrare.forEach(function(email) {
    logSistema("Filtro " + email.id + "...");
    
    var l0Result = analisiLayer0SpamFilter(email);
    
    risultati.push({
      id: email.id,
      isSpam: l0Result.isSpam,
      confidence: l0Result.confidence,
      reason: l0Result.reason || "-"
    });
    
    // Scrivi risultati nel foglio
    scriviRisultatoL0(sheet, email.rowNum, l0Result);
    
    // Aggiorna status
    var nuovoStatus = l0Result.isSpam ? 
      CONFIG.STATUS_EMAIL.SPAM : 
      CONFIG.STATUS_EMAIL.DA_ANALIZZARE;
    
    aggiornaStatus(sheet, email.rowNum, nuovoStatus);
    
    logSistema(email.id + " → " + (l0Result.isSpam ? "SPAM" : "LEGIT") + 
               " (conf: " + l0Result.confidence + "%)");
  });
  
  // Report finale
  var spamCount = risultati.filter(function(r) { return r.isSpam; }).length;
  var legitCount = risultati.length - spamCount;
  
  var report = "🛡️ TEST LAYER 0 COMPLETATO\n\n" +
    "Email analizzate: " + risultati.length + "\n" +
    "🚫 SPAM: " + spamCount + "\n" +
    "✅ LEGIT: " + legitCount + "\n\n" +
    "━━━ DETTAGLIO ━━━\n";
  
  risultati.forEach(function(r) {
    var icon = r.isSpam ? "🚫" : "✅";
    report += icon + " " + r.id + ": " + 
              (r.isSpam ? "SPAM" : "LEGIT") + 
              " (" + r.confidence + "%) - " + r.reason + "\n";
  });
  
  report += "\nVerifica il foglio LOG_IN per dettagli.";
  
  SpreadsheetApp.getUi().alert("Test L0 Completato", report, SpreadsheetApp.getUi().ButtonSet.OK);
  logSistema("🛡️ TEST L0 COMPLETATO: " + spamCount + " spam, " + legitCount + " legit");
}


// ═══════════════════════════════════════════════════════════════════════
// TEST LAYER 1 - GPT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test Layer 1 con email mock
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
    "Sintesi: " + (result.sintesi || "").substring(0, 100) + "...\n" +
    "Modello: " + (result.modello || "-") + "\n\n" +
    (result.success ? "✅ TEST SUPERATO!" : "❌ ERRORE: " + result.error);
  
  ui.alert("Test L1", report, ui.ButtonSet.OK);
  logSistema("🧪 Test L1: Completato");
}


// ═══════════════════════════════════════════════════════════════════════
// TEST ANALISI COMPLETA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test analisi singola email (flusso completo)
 */
function testAnalisiSingola() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🧪 Test Analisi Singola',
    'Eseguirà analisi completa L0→L1→L2→L3 sulla prima email disponibile.\n\n' +
    'Tempo stimato: 10-20 secondi\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    logSistema("🧪 TEST: Avvio analisi singola completa");
    
    // Esegui flusso completo con max 1 email
    var risultati = analizzaEmailInCoda(1, false);
    
    var report = '✅ Test Completato!\n\n' +
      '📊 RISULTATI:\n' +
      '━━━━━━━━━━━━━━━━━━\n' +
      '🛡️ L0: ' + risultati.l0.spam + ' spam, ' + risultati.l0.legit + ' legit\n' +
      '🔵 L1: ' + risultati.l1.ok + ' OK, ' + risultati.l1.errori + ' errori\n';
    
    if (risultati.l2.skipped) {
      report += '🟣 L2: SKIPPED (no Claude)\n';
      report += '🟢 L3: ' + risultati.l3.ok + ' finalizzate da L1\n';
    } else {
      report += '🟣 L2: ' + risultati.l2.ok + ' OK, ' + risultati.l2.errori + ' errori\n';
      report += '🟢 L3: ' + risultati.l3.ok + ' OK, ' + risultati.l3.needsReview + ' review\n';
    }
    
    report += '\nModalità: ' + risultati.modalita + '\n';
    report += '\nControlla LOG_IN per i dettagli.';
    
    ui.alert('Test Completato', report, ui.ButtonSet.OK);
    logSistema("🧪 TEST: Completato con successo");
    
  } catch (error) {
    ui.alert('❌ Errore', error.toString(), ui.ButtonSet.OK);
    logSistema("❌ TEST FALLITO: " + error.toString());
  }
}

/**
 * Test flusso completo su tutte le email in coda
 */
function testFlussoCompleto() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🧪 Test Flusso Completo',
    'Eseguirà il flusso completo su TUTTE le email in coda (max 50).\n\n' +
    'Tempo stimato: 5-10 secondi per email\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var risultati = analizzaEmailInCoda(50, false);
    
    var report = '✅ Flusso Completato!\n\n' +
      '📊 TOTALE: ' + risultati.totali + ' email\n\n' +
      '🛡️ L0: ' + risultati.l0.spam + ' spam, ' + risultati.l0.legit + ' legit\n' +
      '🔵 L1: ' + risultati.l1.ok + ' OK\n' +
      '🟣 L2: ' + (risultati.l2.skipped ? 'SKIPPED' : risultati.l2.ok + ' OK') + '\n' +
      '🟢 L3: ' + risultati.l3.ok + ' OK, ' + (risultati.l3.needsReview || 0) + ' review\n\n' +
      'Modalità: ' + risultati.modalita;
    
    ui.alert('Flusso Completato', report, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ Errore', error.toString(), ui.ButtonSet.OK);
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
    "Aggiunte 6 email:\n• 3 SPAM (SPAM-001, SPAM-002, SPAM-003)\n• 3 LEGIT (LEGIT-001, LEGIT-002, LEGIT-003)\n\nTutte con STATUS = DA_FILTRARE\n\nOra esegui: Menu > Test > Test L0 su Foglio",
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
  
  // Conta per status
  var statusCount = {};
  
  for (var i = 1; i < Math.min(data.length, 21); i++) {
    var row = data[i];
    var id = (row[colId] || "").toString().substring(0, 10);
    var status = (row[colStatus] || "-").toString();
    var l0 = row[colL0] ? "✓" : "-";
    var l1 = row[colL1] ? "✓" : "-";
    var l2 = row[colL2] ? "✓" : "-";
    var l3 = row[colL3] ? "✓" : "-";
    
    // Padding
    while (id.length < 10) id += " ";
    while (status.length < 14) status += " ";
    
    report += id + " | " + status + " | " + l0 + "  | " + l1 + "  | " + l2 + "  | " + l3 + "\n";
    
    // Count
    statusCount[row[colStatus] || "VUOTO"] = (statusCount[row[colStatus] || "VUOTO"] || 0) + 1;
  }
  
  if (data.length > 21) {
    report += "... e altre " + (data.length - 21) + " righe\n";
  }
  
  report += "\n━━━ RIEPILOGO STATUS ━━━\n";
  Object.keys(statusCount).forEach(function(s) {
    report += s + ": " + statusCount[s] + "\n";
  });
  
  SpreadsheetApp.getUi().alert("Debug", report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Test manuale spam su riga 2
 */
function testSpamManuale() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("❌ LOG_IN vuoto o non trovato");
    return;
  }
  
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
  
  var msg = "🧪 TEST LAYER 0 (Riga 2)\n\n" +
    "Email: " + (emailTest.oggetto || "").substring(0, 40) + "...\n\n" +
    "━━━ RISULTATO ━━━\n" +
    "IsSpam: " + (result.isSpam ? "🚫 SI" : "✅ NO") + "\n" +
    "Confidence: " + result.confidence + "%\n" +
    "Reason: " + (result.reason || "-") + "\n" +
    "Modello: " + result.modello;
  
  SpreadsheetApp.getUi().alert("Test", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Reset email test per riprocessarle
 */
function resetEmailTest() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "⚠️ Reset Email Test",
    "Questa operazione resetterà lo STATUS di tutte le email che iniziano con 'TEST-', 'SPAM-' o 'LEGIT-' a DA_FILTRARE.\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  // Colonne da pulire (L0, L1, L2, L3)
  var colsToReset = [
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_TIMESTAMP),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_IS_SPAM),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_CONFIDENCE),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_REASON),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L0_MODELLO),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_TIMESTAMP),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_TAGS),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_SINTESI),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_SCORES_JSON),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L1_MODELLO),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L2_TIMESTAMP),
    headers.indexOf(CONFIG.COLONNE_LOG_IN.L3_TIMESTAMP)
  ];
  
  var count = 0;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = (row[colId] || "").toString();
    
    if (id.indexOf("TEST-") === 0 || id.indexOf("SPAM-") === 0 || id.indexOf("LEGIT-") === 0) {
      var rowNum = i + 1;
      
      // Reset status
      sheet.getRange(rowNum, colStatus + 1).setValue(CONFIG.STATUS_EMAIL.DA_FILTRARE);
      
      // Pulisci colonne analisi
      colsToReset.forEach(function(col) {
        if (col >= 0) {
          sheet.getRange(rowNum, col + 1).setValue("");
        }
      });
      
      count++;
    }
  }
  
  ui.alert("✅ Reset Completato", "Resettate " + count + " email test.", ui.ButtonSet.OK);
  logSistema("Reset " + count + " email test");
}