/**
 * ==========================================================================================
 * INIZIALIZZAZIONE - MOTORE FORNITORI v1.1.0
 * ==========================================================================================
 * Setup fogli, menu e dati iniziali
 * 
 * CHANGELOG v1.1.0:
 * - Aggiunte 5 colonne per integrazione Questioni Engine
 * - Aggiunto menu Upgrade per sistemi esistenti
 * - Setup automatico include colonne Questioni
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👥 Fornitori Engine v1.1')
    .addItem('🛠️ Setup Completo', 'setupCompleto')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Fornitori')
      .addItem('➕ Aggiungi Fornitore', 'menuAggiungiFornitore')
      .addItem('🔍 Cerca Fornitore', 'menuCercaFornitore')
      .addItem('📊 Report Fornitori', 'menuReportFornitori'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔗 Integrazioni')
      .addItem('🔄 Sync con Email Engine', 'menuSyncEmailEngine')
      .addItem('📧 Test Lookup Email', 'menuTestLookupEmail'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Upgrade')
      .addItem('➕ Aggiungi Colonne Questioni', 'aggiungiColonneQuestioni')
      .addItem('✅ Verifica Colonne', 'verificaColonneQuestioni'))
    .addSeparator()
    .addItem('📜 Storico Azioni', 'menuStoricoAzioni')
    .addItem('ℹ️ Info Sistema', 'menuInfoSistema')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP COMPLETO
// ═══════════════════════════════════════════════════════════════════════

function setupCompleto() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🛠️ Setup Motore Fornitori v1.1.0',
    'Questa operazione creerà:\n\n' +
    '✅ 6 Fogli preconfigurati\n' +
    '✅ Struttura 37 colonne FORNITORI\n' +
    '   (include 5 colonne Questioni Engine)\n' +
    '✅ Storico Azioni\n' +
    '✅ Foglio Prompts (vuoto)\n' +
    '✅ Dashboard HOME\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Crea Fogli
    creaFogli(ss);
    
    // 2. Popola SETUP
    popolaSetup(ss);
    
    // 3. Crea fornitore test
    creaFornitoreTest(ss);
    
    logSistemaFornitori("Setup completo v1.1.0 eseguito con successo");
    
    ui.alert(
      '✅ Sistema Pronto!',
      'Setup completato:\n\n' +
      '📁 6 Fogli creati\n' +
      '📋 37 colonne FORNITORI\n' +
      '   (include supporto Questioni Engine)\n' +
      '📜 Storico Azioni pronto\n' +
      '👤 1 Fornitore test\n\n' +
      'Prossimi step:\n' +
      '1. Importa fornitori esistenti\n' +
      '2. Configura integrazione Email Engine\n' +
      '3. Configura integrazione Questioni Engine',
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore Setup', e.toString(), ui.ButtonSet.OK);
    logSistemaFornitori("ERRORE Setup: " + e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE FOGLI
// ═══════════════════════════════════════════════════════════════════════

function creaFogli(ss) {
  var fogli = [
    { nome: CONFIG_FORNITORI.SHEETS.HOME, colore: "#134F5C" },
    { nome: CONFIG_FORNITORI.SHEETS.FORNITORI, colore: "#6AA84F" },
    { nome: CONFIG_FORNITORI.SHEETS.SETUP, colore: "#666666" },
    { nome: CONFIG_FORNITORI.SHEETS.PROMPTS, colore: "#F1C232" },
    { nome: CONFIG_FORNITORI.SHEETS.STORICO_AZIONI, colore: "#9B59B6" },
    { nome: CONFIG_FORNITORI.SHEETS.LOG_SISTEMA, colore: "#000000" }
  ];
  
  fogli.forEach(function(foglio) {
    var sheet = ss.getSheetByName(foglio.nome);
    if (!sheet) {
      sheet = ss.insertSheet(foglio.nome);
      sheet.setTabColor(foglio.colore);
      Logger.log("✅ Creato: " + foglio.nome);
    }
    
    // Headers specifici
    if (foglio.nome === CONFIG_FORNITORI.SHEETS.FORNITORI) {
      setupFoglioFornitori(sheet);
    }
    
    if (foglio.nome === CONFIG_FORNITORI.SHEETS.STORICO_AZIONI) {
      setupFoglioStorico(sheet);
    }
    
    if (foglio.nome === CONFIG_FORNITORI.SHEETS.PROMPTS) {
      setupFoglioPrompts(sheet);
    }
    
    if (foglio.nome === CONFIG_FORNITORI.SHEETS.LOG_SISTEMA) {
      setupFoglioLog(sheet);
    }
    
    if (foglio.nome === CONFIG_FORNITORI.SHEETS.HOME) {
      setupFoglioHome(sheet);
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP SINGOLI FOGLI
// ═══════════════════════════════════════════════════════════════════════

function setupFoglioFornitori(sheet) {
  if (sheet.getLastRow() > 0) return; // Già configurato
  
  var headers = Object.values(CONFIG_FORNITORI.COLONNE_FORNITORI);
  
  // Scrivi headers
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#E8F5E9")
    .setHorizontalAlignment("center");
  
  // Colora headers Questioni in modo diverso (ultime 5)
  var colQuestioni = headers.length - 4; // Inizio colonne Questioni
  sheet.getRange(1, colQuestioni, 1, 5)
    .setBackground("#E3F2FD"); // Blu chiaro per Questioni
  
  // Freeze
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  
  // Larghezza colonne principali
  sheet.setColumnWidth(1, 100);  // ID
  sheet.setColumnWidth(2, 50);   // Priorità
  sheet.setColumnWidth(3, 200);  // Nome Azienda
  sheet.setColumnWidth(4, 200);  // Email Ordini
  
  // Formattazione condizionale per Priorità Urgente
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("X")
    .setBackground("#FFCDD2")
    .setRanges([sheet.getRange("B:B")])
    .build();
  sheet.setConditionalFormatRules([rule]);
  
  Logger.log("✅ Setup FORNITORI completato (37 colonne)");
}

function setupFoglioStorico(sheet) {
  if (sheet.getLastRow() > 0) return;
  
  var headers = Object.values(CONFIG_FORNITORI.COLONNE_STORICO);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#E1BEE7");
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(5, 300); // Descrizione
  
  Logger.log("✅ Setup STORICO_AZIONI completato");
}

function setupFoglioPrompts(sheet) {
  if (sheet.getLastRow() > 0) return;
  
  var headers = Object.values(CONFIG_FORNITORI.COLONNE_PROMPTS);
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#FFF9C4");
  
  sheet.setFrozenRows(1);
  
  Logger.log("✅ Setup PROMPTS completato (vuoto)");
}

function setupFoglioLog(sheet) {
  if (sheet.getLastRow() > 0) return;
  
  sheet.getRange("A1:B1")
    .setValues([["TIMESTAMP", "MESSAGGIO"]])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF");
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 600);
  
  Logger.log("✅ Setup LOG_SISTEMA completato");
}

function setupFoglioHome(sheet) {
  if (sheet.getLastRow() > 0) return;
  
  var dashboard = [
    ["📊 DASHBOARD FORNITORI", ""],
    ["", ""],
    ["Totale Fornitori", "=COUNTA(FORNITORI!A:A)-1"],
    ["Fornitori Attivi", '=COUNTIF(FORNITORI!AF:AF,"ATTIVO")'],
    ["Priorità Urgente", '=COUNTIF(FORNITORI!B:B,"X")'],
    ["", ""],
    ["📅 PROSSIME PROMO", ""],
    ["Promo questa settimana", '=COUNTIFS(FORNITORI!Q:Q,">="&TODAY(),FORNITORI!Q:Q,"<="&TODAY()+7)'],
    ["", ""],
    ["🎫 QUESTIONI", ""],
    ["Fornitori con questioni aperte", '=COUNTIF(FORNITORI!AG:AG,">0")'],
    ["Totale questioni aperte", "=SUM(FORNITORI!AG:AG)"],
    ["", ""],
    ["📈 PERFORMANCE", ""],
    ["Score Medio", "=AVERAGE(FORNITORI!AE:AE)"],
    ["", ""],
    ["🕐 Ultimo aggiornamento", "=NOW()"]
  ];
  
  sheet.getRange(1, 1, dashboard.length, 2).setValues(dashboard);
  
  // Formattazione
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A7").setFontWeight("bold");
  sheet.getRange("A10").setFontWeight("bold");
  sheet.getRange("A14").setFontWeight("bold");
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  
  Logger.log("✅ Setup HOME completato");
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP
// ═══════════════════════════════════════════════════════════════════════

function popolaSetup(ss) {
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.SETUP);
  if (sheet.getLastRow() > 1) return;
  
  sheet.getRange("A1:B1")
    .setValues([["CHIAVE", "VALORE"]])
    .setFontWeight("bold")
    .setBackground("#EEEEEE");
  
  var dati = [
    ["=== SISTEMA ===", ""],
    ["VERSIONE", CONFIG_FORNITORI.VERSION],
    ["DATA_RELEASE", CONFIG_FORNITORI.DATA_RELEASE],
    ["", ""],
    ["=== INTEGRAZIONE EMAIL ENGINE ===", ""],
    ["EMAIL_ENGINE_SHEET_ID", ""],
    ["SYNC_AUTOMATICO", "SI"],
    ["LOOKUP_CACHE_TTL", "300"],
    ["", ""],
    ["=== INTEGRAZIONE QUESTIONI ENGINE ===", ""],
    ["QUESTIONI_ENGINE_SHEET_ID", ""],
    ["STATS_UPDATE_AUTOMATICO", "SI"],
    ["", ""],
    ["=== DEFAULTS ===", ""],
    ["DEFAULT_STATUS", CONFIG_FORNITORI.DEFAULTS.STATUS_INIZIALE],
    ["DEFAULT_PERFORMANCE", CONFIG_FORNITORI.DEFAULTS.PERFORMANCE_SCORE],
    ["DEFAULT_SCONTO_TARGET", CONFIG_FORNITORI.DEFAULTS.SCONTO_TARGET]
  ];
  
  sheet.getRange(2, 1, dati.length, 2).setValues(dati);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);
}

// ═══════════════════════════════════════════════════════════════════════
// FORNITORE TEST
// ═══════════════════════════════════════════════════════════════════════

function creaFornitoreTest(ss) {
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  if (sheet.getLastRow() > 1) return;
  
  var fornitoreTest = [
    "FOR-001",           // ID_FORNITORE
    "",                  // PRIORITA_URGENTE
    "BioCosmetics Test", // NOME_AZIENDA
    "ordini@biocosmetics-test.it", // EMAIL_ORDINI
    "info@biocosmetics-test.it, commerciale@biocosmetics-test.it", // EMAIL_ALTRI
    "Mario Rossi",       // CONTATTO
    "+39 02 1234567",    // TELEFONO
    "IT12345678901",     // PARTITA_IVA
    "Via Roma 1, Milano",// INDIRIZZO
    "Fornitore di test per sviluppo sistema", // NOTE
    "EMAIL_DIRETTA",     // METODO_INVIO
    "EXCEL",             // TIPO_LISTINO
    "",                  // LINK_LISTINO
    100,                 // MIN_ORDINE
    5,                   // LEAD_TIME_GIORNI
    "SI",                // PROMO_SU_RICHIESTA
    "",                  // DATA_PROSSIMA_PROMO
    4,                   // PROMO_ANNUALI_NUM
    2,                   // PROMO_ANNUALI_RICHIESTE
    15,                  // SCONTO_TARGET
    "",                  // RICHIESTA_SCONTO
    "NON_RICHIESTO",     // ESITO_RICHIESTA_SCONTO
    "",                  // TIPO_SCONTO_CONCESSO
    10,                  // SCONTO_PERCENTUALE
    "",                  // DATA_ULTIMO_ORDINE
    "",                  // DATA_ULTIMA_EMAIL
    "",                  // DATA_ULTIMA_ANALISI
    "",                  // STATUS_ULTIMA_AZIONE
    0,                   // EMAIL_ANALIZZATE_COUNT
    3,                   // PERFORMANCE_SCORE
    "ATTIVO",            // STATUS_FORNITORE
    new Date(),          // DATA_CREAZIONE
    // ═══ COLONNE QUESTIONI (33-37) ═══
    0,                   // QUESTIONI_APERTE
    0,                   // QUESTIONI_TOTALI
    "",                  // ULTIMA_QUESTIONE
    "",                  // TIPO_QUESTIONI
    ""                   // NOTE_QUESTIONI
  ];
  
  sheet.appendRow(fornitoreTest);
  
  // Registra in storico
  registraAzione("FOR-001", "BioCosmetics Test", "CREAZIONE", "Fornitore test creato da setup");
  
  Logger.log("✅ Fornitore test creato (37 campi)");
}

// ═══════════════════════════════════════════════════════════════════════
// UPGRADE v1.1.0 - COLONNE QUESTIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiunge le 5 colonne per Questioni Engine a un sistema esistente
 * Usa questa funzione se hai già un Fornitori Engine v1.0.0
 */
function aggiungiColonneQuestioni() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "🔧 Upgrade a v1.1.0",
    "Questa operazione aggiungerà 5 nuove colonne:\n\n" +
    "• QUESTIONI_APERTE (Numero)\n" +
    "• QUESTIONI_TOTALI (Numero)\n" +
    "• ULTIMA_QUESTIONE (Data)\n" +
    "• TIPO_QUESTIONI (Testo)\n" +
    "• NOTE_QUESTIONI (Testo)\n\n" +
    "Queste colonne permettono l'integrazione con Questioni Engine.\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
    
    if (!sheet) {
      ui.alert("❌ Errore", "Foglio FORNITORI non trovato!", ui.ButtonSet.OK);
      return;
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Colonne da aggiungere
    var colonneNuove = [
      "QUESTIONI_APERTE",
      "QUESTIONI_TOTALI",
      "ULTIMA_QUESTIONE",
      "TIPO_QUESTIONI",
      "NOTE_QUESTIONI"
    ];
    
    var esistenti = [];
    var daAggiungere = [];
    
    colonneNuove.forEach(function(col) {
      if (headers.indexOf(col) >= 0) {
        esistenti.push(col);
      } else {
        daAggiungere.push(col);
      }
    });
    
    if (esistenti.length === colonneNuove.length) {
      ui.alert(
        "✅ Già Aggiornato",
        "Tutte le colonne Questioni sono già presenti!\n\n" +
        "Il sistema è già alla versione 1.1.0",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Aggiungi colonne mancanti
    var ultimaColonna = sheet.getLastColumn();
    var righe = sheet.getLastRow();
    
    daAggiungere.forEach(function(nomeColonna, index) {
      var nuovaCol = ultimaColonna + index + 1;
      
      // Header
      sheet.getRange(1, nuovaCol).setValue(nomeColonna);
      
      // Default values per righe esistenti
      if (righe > 1) {
        if (nomeColonna === "QUESTIONI_APERTE" || nomeColonna === "QUESTIONI_TOTALI") {
          sheet.getRange(2, nuovaCol, righe - 1, 1).setValue(0);
        }
      }
    });
    
    // Formattazione header nuove colonne
    var nuoveColonne = sheet.getRange(1, ultimaColonna + 1, 1, daAggiungere.length);
    nuoveColonne.setFontWeight("bold");
    nuoveColonne.setBackground("#E3F2FD"); // Blu chiaro
    
    logSistemaFornitori("✅ Upgrade v1.1.0: Aggiunte colonne " + daAggiungere.join(", "));
    
    ui.alert(
      "✅ Upgrade Completato!",
      "Aggiunte " + daAggiungere.length + " colonne:\n\n" +
      "• " + daAggiungere.join("\n• ") + "\n\n" +
      "Il Fornitori Engine è ora alla versione 1.1.0\n\n" +
      "📝 PROSSIMO STEP:\n" +
      "In Questioni Engine: Menu > Fornitori > Sync da Master",
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
    logSistemaFornitori("❌ Errore upgrade: " + e.toString());
  }
}

/**
 * Verifica stato colonne Questioni
 */
function verificaColonneQuestioni() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.FORNITORI);
  
  if (!sheet) {
    ui.alert("❌ Foglio FORNITORI non trovato");
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colonneQuestioni = [
    "QUESTIONI_APERTE",
    "QUESTIONI_TOTALI",
    "ULTIMA_QUESTIONE",
    "TIPO_QUESTIONI",
    "NOTE_QUESTIONI"
  ];
  
  var report = "📊 VERIFICA COLONNE QUESTIONI\n";
  report += "══════════════════════════════\n\n";
  report += "Totale colonne: " + headers.length + "\n\n";
  
  var presenti = 0;
  colonneQuestioni.forEach(function(col) {
    var idx = headers.indexOf(col);
    if (idx >= 0) {
      report += "✅ " + col + " (colonna " + (idx + 1) + ")\n";
      presenti++;
    } else {
      report += "❌ " + col + " (MANCANTE)\n";
    }
  });
  
  report += "\n";
  if (presenti === colonneQuestioni.length) {
    report += "✅ Sistema aggiornato a v1.1.0\n";
    report += "Pronto per integrazione Questioni Engine!";
  } else {
    report += "⚠️ " + (colonneQuestioni.length - presenti) + " colonne mancanti.\n";
    report += "Esegui: Menu > Upgrade > Aggiungi Colonne Questioni";
  }
  
  ui.alert("Verifica Colonne", report, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// MENU HANDLERS
// ═══════════════════════════════════════════════════════════════════════

function menuAggiungiFornitore() {
  var ui = SpreadsheetApp.getUi();
  
  var nomeRisposta = ui.prompt('Nuovo Fornitore', 'Nome Azienda:', ui.ButtonSet.OK_CANCEL);
  if (nomeRisposta.getSelectedButton() !== ui.Button.OK) return;
  var nome = nomeRisposta.getResponseText().trim();
  
  var emailRisposta = ui.prompt('Nuovo Fornitore', 'Email Ordini:', ui.ButtonSet.OK_CANCEL);
  if (emailRisposta.getSelectedButton() !== ui.Button.OK) return;
  var email = emailRisposta.getResponseText().trim();
  
  if (!nome || !email) {
    ui.alert('❌ Errore', 'Nome e Email sono obbligatori', ui.ButtonSet.OK);
    return;
  }
  
  var id = creaFornitore(nome, email);
  
  ui.alert('✅ Fornitore Creato', 'ID: ' + id + '\nNome: ' + nome, ui.ButtonSet.OK);
}

function menuCercaFornitore() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.prompt('Cerca Fornitore', 'Email o Nome:', ui.ButtonSet.OK_CANCEL);
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var query = risposta.getResponseText().trim();
  var risultato = cercaFornitore(query);
  
  if (risultato) {
    ui.alert('✅ Trovato', 
      'ID: ' + risultato.id + '\n' +
      'Nome: ' + risultato.nome + '\n' +
      'Email: ' + risultato.email + '\n' +
      'Status: ' + risultato.status,
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Non Trovato', 'Nessun fornitore trovato per: ' + query, ui.ButtonSet.OK);
  }
}

function menuTestLookupEmail() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.prompt('Test Lookup', 'Inserisci email mittente:', ui.ButtonSet.OK_CANCEL);
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var email = risposta.getResponseText().trim();
  var fornitore = lookupFornitoreByEmail(email);
  
  if (fornitore) {
    var contesto = generaContestoFornitore(fornitore);
    ui.alert('✅ Fornitore Trovato', 
      'Nome: ' + fornitore.nome + '\n\n' +
      'Contesto per Email Engine:\n' + contesto,
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Non Trovato', 'Nessun fornitore associato a: ' + email, ui.ButtonSet.OK);
  }
}

function menuReportFornitori() {
  var report = generaReportFornitori();
  SpreadsheetApp.getUi().alert('📊 Report Fornitori', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuStoricoAzioni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.STORICO_AZIONI);
  ss.setActiveSheet(sheet);
}

function menuSyncEmailEngine() {
  SpreadsheetApp.getUi().alert('🔄 Sync', 'Funzione in sviluppo.\n\nPermette di aggiornare automaticamente DATA_ULTIMA_EMAIL quando Email Engine analizza email.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuInfoSistema() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '👥 Fornitori Engine',
    'Versione: ' + CONFIG_FORNITORI.VERSION + '\n' +
    'Release: ' + CONFIG_FORNITORI.DATA_RELEASE + '\n\n' +
    '✅ Funzionalità:\n' +
    '• Anagrafica fornitori (37 campi)\n' +
    '• Lookup email per Email Engine\n' +
    '• Storico azioni automatico\n' +
    '• Dashboard KPI\n' +
    '• Integrazione Questioni Engine\n\n' +
    '🔗 Integrazione Email:\n' +
    '• Skip L0 spam per fornitori noti\n' +
    '• Contesto fornitore in L1 analysis\n\n' +
    '🔗 Integrazione Questioni:\n' +
    '• Stats questioni per fornitore\n' +
    '• Tracking questioni aperte/totali',
    ui.ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════════════════

function logSistemaFornitori(messaggio) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.LOG_SISTEMA);
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}

function registraAzione(idFornitore, nomeFornitore, tipoAzione, descrizione, emailCollegata) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_FORNITORI.SHEETS.STORICO_AZIONI);
    
    sheet.appendRow([
      new Date(),
      idFornitore,
      nomeFornitore,
      tipoAzione,
      descrizione,
      emailCollegata || "",
      Session.getActiveUser().getEmail() || "SISTEMA",
      ""
    ]);
  } catch(e) {
    Logger.log("Errore registrazione azione: " + e.toString());
  }
}