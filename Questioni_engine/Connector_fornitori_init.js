/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - INIT v1.0.0 (per Questioni Engine)
 * ==========================================================================================
 * Inizializzazione, setup e menu per il connettore Fornitori
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// INIZIALIZZAZIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Setup iniziale del connettore Fornitori per Questioni Engine
 * Aggiunge configurazione in SETUP e prepara tab sync
 */
function initConnectorFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "🔗 Inizializza Connector Fornitori (per Questioni)",
    "Questa operazione:\n\n" +
    "1. Aggiunge chiavi configurazione in SETUP\n" +
    "2. Crea tab " + CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_SYNC + "\n" +
    "3. Prepara il connettore per sync e stats\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Verifica SETUP esiste
    var setupSheet = ss.getSheetByName("SETUP");
    if (!setupSheet) {
      ui.alert("❌ Errore", "Foglio SETUP non trovato. Esegui prima il setup di Questioni Engine.", ui.ButtonSet.OK);
      return;
    }
    
    // 2. Aggiungi chiavi in SETUP
    aggiungiChiaviSetupFornitoriQE(setupSheet);
    
    // 3. Crea tab FORNITORI_SYNC
    var syncSheet = ss.getSheetByName(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_SYNC);
    if (!syncSheet) {
      syncSheet = ss.insertSheet(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_SYNC);
      syncSheet.setTabColor("#A5D6A7");
      syncSheet.getRange("A1").setValue("In attesa di sync dal master Fornitori Engine...");
      syncSheet.getRange("A2").setValue("Configura ID master in SETUP, poi esegui Sync.");
      
      logConnectorFornitoriQE("Creata tab FORNITORI_SYNC");
    }
    
    logConnectorFornitoriQE("Connettore inizializzato");
    
    ui.alert(
      "✅ Connettore Inizializzato",
      "Prossimi step:\n\n" +
      "1. Vai nel foglio SETUP\n" +
      "2. Trova: CONNECTOR_FORNITORI_MASTER_ID\n" +
      "3. Incolla l'ID del foglio Fornitori Engine\n" +
      "4. Menu > Connettori > Fornitori > Sync\n\n" +
      "💡 Tip: L'ID del foglio è nella URL dopo /d/",
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
    logConnectorFornitoriQE("Errore init: " + e.toString());
  }
}

/**
 * Aggiunge chiavi configurazione in SETUP
 */
function aggiungiChiaviSetupFornitoriQE(setupSheet) {
  var data = setupSheet.getDataRange().getValues();
  var chiavi = data.map(function(row) { return row[0]; });
  
  var nuoveChiavi = [
    ["", ""],
    ["═══ CONNECTOR FORNITORI ═══", ""],
    [CONNECTOR_FORNITORI_QE.SETUP_KEYS.ENABLED, "SI"],
    [CONNECTOR_FORNITORI_QE.SETUP_KEYS.MASTER_SHEET_ID, "Incolla qui ID foglio Fornitori Engine"],
    [CONNECTOR_FORNITORI_QE.SETUP_KEYS.SYNC_INTERVAL, CONNECTOR_FORNITORI_QE.DEFAULTS.SYNC_INTERVAL_MIN],
    [CONNECTOR_FORNITORI_QE.SETUP_KEYS.LAST_SYNC, ""],
    [CONNECTOR_FORNITORI_QE.SETUP_KEYS.LAST_STATS_UPDATE, ""]
  ];
  
  var hasConnectorKeys = chiavi.indexOf(CONNECTOR_FORNITORI_QE.SETUP_KEYS.MASTER_SHEET_ID) >= 0;
  
  if (!hasConnectorKeys) {
    nuoveChiavi.forEach(function(row) {
      setupSheet.appendRow(row);
    });
    
    logConnectorFornitoriQE("Aggiunte chiavi configurazione in SETUP");
  }
}

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea sottomenu Connettore Fornitori (da chiamare in onOpen)
 * @returns {Object} Submenu object
 */
function creaSubmenuConnettoreFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  return ui.createMenu("📦 Fornitori")
    .addItem("🔗 Inizializza Connettore", "initConnectorFornitoriQE")
    .addSeparator()
    .addItem("🔄 Sync da Master", "syncFornitoriQEMenu")
    .addItem("⏰ Configura Trigger Sync", "configuraTriggerSyncFornitoriQE")
    .addSeparator()
    .addSubMenu(ui.createMenu("📊 Stats → Fornitori Engine")
      .addItem("Aggiorna Singolo Fornitore", "menuAggiornaStatsFornitoreQE")
      .addItem("Aggiorna TUTTI i Fornitori", "menuAggiornaStatsTuttiFornitoriQE")
      .addItem("📋 Report Questioni x Fornitore", "menuReportQuestoniFornitori"))
    .addSeparator()
    .addItem("🔍 Test Lookup Fornitore", "testLookupFornitoreQE")
    .addItem("🧪 Test Connessione Master", "testConnessioneMasterFornitoriQE")
    .addSeparator()
    .addItem("📊 Stato Connettore", "mostraStatoConnettoreFornitoriQE");
}

/**
 * Aggiungi menu Connettori al menu principale di Questioni Engine
 * Da chiamare in onOpen()
 */
function aggiungiMenuConnettoreFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("🔗 Connettori")
    .addSubMenu(creaSubmenuConnettoreFornitoriQE())
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Menu: Mostra stato connettore
 */
function mostraStatoConnettoreFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  var attivo = isConnectorFornitoriQEAttivo();
  var check = checkConnectorFornitoriQEReady();
  var stats = getStatsSyncFornitoriQE();
  
  var report = "🔗 CONNECTOR FORNITORI (Questioni Engine)\n";
  report += "══════════════════════════════════════════\n\n";
  
  report += "Stato: " + (attivo ? "✅ ATTIVO" : "❌ NON ATTIVO") + "\n\n";
  
  report += "━━━ CONFIGURAZIONE ━━━\n";
  report += "Master configurato: " + (stats.masterConfigurato ? "✅" : "❌") + "\n";
  report += "Tab sync presente: " + (stats.tabEsiste ? "✅" : "❌") + "\n";
  report += "Fornitori locali: " + stats.righe + "\n\n";
  
  report += "━━━ SYNC ━━━\n";
  report += "Ultimo sync: " + (stats.ultimoSync ? stats.ultimoSync.toLocaleString('it-IT') : "Mai") + "\n";
  report += "Intervallo: " + stats.intervallo + " minuti\n";
  
  // Stats update
  var lastStats = getConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.LAST_STATS_UPDATE, null);
  report += "\n━━━ STATS UPDATE ━━━\n";
  report += "Ultimo update: " + (lastStats ? new Date(lastStats).toLocaleString('it-IT') : "Mai") + "\n";
  
  if (!attivo && !check.canActivate) {
    report += "\n━━━ AZIONI RICHIESTE ━━━\n";
    report += "⚠️ " + check.reason + "\n";
  }
  
  ui.alert("Stato Connettore Fornitori", report, ui.ButtonSet.OK);
}

/**
 * Menu: Test lookup fornitore
 */
function testLookupFornitoreQE() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isConnectorFornitoriQEAttivo()) {
    ui.alert(
      "❌ Connettore Non Attivo",
      "Inizializza e configura il connettore prima di eseguire il test.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var risposta = ui.prompt(
    "🔍 Test Lookup Fornitore",
    "Inserisci email o dominio da cercare:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var query = risposta.getResponseText().trim();
  if (!query) {
    ui.alert("Query non valida");
    return;
  }
  
  // Prova lookup
  var risultato = associaFornitoreAQuestione(query, query);
  
  var report = "🔍 RISULTATO LOOKUP\n";
  report += "══════════════════════════════\n\n";
  report += "Query: " + query + "\n\n";
  
  if (risultato.trovato) {
    report += "✅ FORNITORE TROVATO!\n\n";
    report += "ID: " + risultato.idFornitore + "\n";
    report += "Nome: " + risultato.nomeFornitore + "\n";
    
    if (risultato.fornitore) {
      report += "Email: " + risultato.fornitore.email + "\n";
      report += "Status: " + risultato.fornitore.status + "\n";
      
      if (risultato.fornitore.prioritaUrgente) {
        report += "⚠️ PRIORITA' URGENTE\n";
      }
    }
  } else {
    report += "❌ Nessun fornitore trovato.\n";
  }
  
  ui.alert("Risultato Lookup", report, ui.ButtonSet.OK);
}

/**
 * Menu: Test connessione master
 */
function testConnessioneMasterFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  var check = checkConnectorFornitoriQEReady();
  
  if (!check.canActivate) {
    ui.alert("❌ Test Fallito", check.reason, ui.ButtonSet.OK);
    return;
  }
  
  try {
    var masterId = getConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.MASTER_SHEET_ID, "");
    var masterSs = SpreadsheetApp.openById(masterId);
    var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_MASTER);
    
    var righe = masterSheet.getLastRow() - 1;
    var colonne = masterSheet.getLastColumn();
    
    // Verifica colonne stats
    var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    var colonneStats = [
      "QUESTIONI_APERTE",
      "QUESTIONI_TOTALI", 
      "ULTIMA_QUESTIONE",
      "TIPO_QUESTIONI",
      "NOTE_QUESTIONI"
    ];
    
    var colonneMancanti = [];
    colonneStats.forEach(function(col) {
      if (headers.indexOf(col) < 0) {
        colonneMancanti.push(col);
      }
    });
    
    var report = "✅ Connessione OK!\n\n";
    report += "Master: " + masterSs.getName() + "\n";
    report += "Fornitori: " + righe + "\n";
    report += "Colonne: " + colonne + "\n\n";
    
    if (colonneMancanti.length > 0) {
      report += "⚠️ COLONNE STATS MANCANTI:\n";
      report += colonneMancanti.join(", ") + "\n\n";
      report += "Aggiungi queste colonne al Fornitori Engine\n";
      report += "per abilitare update stats automatico.";
    } else {
      report += "✅ Tutte le colonne stats presenti!";
    }
    
    ui.alert("Test Connessione", report, ui.ButtonSet.OK);
    
    logConnectorFornitoriQE("Test connessione OK");
    
  } catch(e) {
    ui.alert("❌ Test Fallito", e.toString(), ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// DISATTIVAZIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Disattiva connettore Fornitori
 */
function disattivaConnectorFornitoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "⚠️ Disattiva Connettore",
    "Vuoi disattivare il Connector Fornitori?\n\n" +
    "Il connettore non verrà più usato durante l'analisi questioni.\n" +
    "I dati locali rimarranno, potrai riattivarlo in qualsiasi momento.",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  setConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.ENABLED, "NO");
  rimuoviTriggerSyncFornitoriQE();
  
  logConnectorFornitoriQE("Connettore disattivato");
  
  ui.alert("✅ Connettore Disattivato", "Puoi riattivarlo modificando ENABLED = SI in SETUP.", ui.ButtonSet.OK);
}

/**
 * Riattiva connettore Fornitori
 */
function riattivaConnectorFornitoriQE() {
  setConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.ENABLED, "SI");
  logConnectorFornitoriQE("Connettore riattivato");
  
  SpreadsheetApp.getUi().alert("✅ Connettore Riattivato");
}

// ═══════════════════════════════════════════════════════════════════════
// PATCH PER QUESTIONI ENGINE - INTEGRAZIONE COMPLETA
// ═══════════════════════════════════════════════════════════════════════

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║                    PATCH 1: CREAZIONE QUESTIONE                       ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Aggiunge lookup fornitore automatico durante la creazione di una questione.
 * 
 * DOVE: Nella funzione creaQuestione() o creaQuestioneDaEmail()
 * QUANDO: Dopo aver estratto il dominio/email mittente
 * 
 * SNIPPET DA COPIARE:
 * ─────────────────────────────────────────────────────────────────────────
 */
function _PATCH_SNIPPET_CREAZIONE_QUESTIONE() {
  // ╔═══════════════════════════════════════════════════════════════════╗
  // ║ COPIA QUESTO BLOCCO nella tua funzione creaQuestione()            ║
  // ╚═══════════════════════════════════════════════════════════════════╝
  
  var emailMittente = ""; // <-- Sostituisci con la variabile reale
  var dominioEmail = "";  // <-- Sostituisci con la variabile reale (opzionale)
  
  // === PATCH CONNECTOR FORNITORI ===
  if (typeof associaFornitoreAQuestione === 'function') {
    var lookupFornitore = associaFornitoreAQuestione(dominioEmail, emailMittente);
    if (lookupFornitore.trovato) {
      // Usa queste variabili per popolare i campi della questione
      var idFornitore = lookupFornitore.idFornitore;
      var nomeFornitore = lookupFornitore.nomeFornitore;
      var contestoFornitore = lookupFornitore.contesto; // Per arricchire analisi AI
      
      // Esempio: questione.idFornitore = idFornitore;
      // Esempio: questione.nomeFornitore = nomeFornitore;
      
      Logger.log("🔗 Fornitore associato: " + nomeFornitore);
    }
  }
  // === FINE PATCH ===
}

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║                    PATCH 2: MODIFICA/CHIUSURA QUESTIONE               ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Aggiorna le statistiche sul Fornitori Engine dopo ogni modifica.
 * 
 * DOVE: Nella funzione salvaQuestione(), chiudiQuestione(), etc.
 * QUANDO: Dopo il salvataggio riuscito
 * 
 * SNIPPET DA COPIARE:
 * ─────────────────────────────────────────────────────────────────────────
 */
function _PATCH_SNIPPET_MODIFICA_QUESTIONE() {
  // ╔═══════════════════════════════════════════════════════════════════╗
  // ║ COPIA QUESTO BLOCCO dopo il salvataggio della questione           ║
  // ╚═══════════════════════════════════════════════════════════════════╝
  
  var idFornitore = ""; // <-- Sostituisci con la variabile reale
  var tipoEvento = "MODIFICATA"; // Può essere: "CREATA", "MODIFICATA", "CHIUSA"
  
  // === PATCH CONNECTOR FORNITORI - UPDATE STATS ===
  if (idFornitore && typeof onQuestioneModificata === 'function') {
    try {
      onQuestioneModificata(idFornitore, tipoEvento);
      Logger.log("🔗 Stats fornitore aggiornate: " + idFornitore);
    } catch(e) {
      Logger.log("⚠️ Errore update stats fornitore: " + e.message);
    }
  }
  // === FINE PATCH ===
}

// ═══════════════════════════════════════════════════════════════════════
// PATCH 3: MENU ONOPEN
// ═══════════════════════════════════════════════════════════════════════

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║  PATCH ONOPEN - COPIA QUESTA FUNZIONE NEL TUO QUESTIONI ENGINE       ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Questa funzione sostituisce o integra la tua onOpen() esistente.
 * Include automaticamente il menu Connettori se i file sono presenti.
 * 
 * ISTRUZIONI:
 * 1. Trova la tua funzione onOpen() in Questioni Engine
 * 2. Sostituiscila con questa OPPURE aggiungi il blocco Connettori
 */
function onOpen_PATCHED_QE() {
  var ui = SpreadsheetApp.getUi();
  
  // ═══ MENU PRINCIPALE QUESTIONI ENGINE ═══
  var menu = ui.createMenu('🎯 Questioni Engine')
    .addItem('🛠️ Setup Completo', 'setupCompleto')
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 Questioni')
      .addItem('➕ Nuova Questione', 'creaQuestioneManuale')
      .addItem('📋 Lista Questioni Aperte', 'mostraQuestioniAperte')
      .addItem('🔍 Cerca Questione', 'cercaQuestione'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔑 API Keys')
      .addItem('OpenAI', 'impostaApiKeyOpenAI')
      .addItem('Claude', 'impostaApiKeyClaude')
      .addItem('🧪 Test Connessioni', 'testConnessioniAI'))
    .addSeparator()
    .addItem('⏰ Configura Trigger', 'configuraTrigger')
    .addItem('ℹ️ Info & Changelog', 'showChangelog');
  
  // ═══ PATCH: AGGIUNGI MENU CONNETTORI SE PRESENTE ═══
  if (typeof creaSubmenuConnettoreFornitoriQE === 'function') {
    menu.addSeparator()
        .addSubMenu(ui.createMenu('🔗 Connettori')
          .addSubMenu(creaSubmenuConnettoreFornitoriQE()));
  }
  
  menu.addToUi();
}

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║  SNIPPET ALTERNATIVO - SE HAI GIA' UN ONOPEN                         ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Se hai già una onOpen() funzionante, aggiungi SOLO questo blocco
 * PRIMA della riga finale .addToUi();
 * 
 * // ═══ CONNECTOR FORNITORI ═══
 * if (typeof creaSubmenuConnettoreFornitoriQE === 'function') {
 *   menu.addSeparator()
 *       .addSubMenu(ui.createMenu('🔗 Connettori')
 *         .addSubMenu(creaSubmenuConnettoreFornitoriQE()));
 * }
 */

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║  VERSIONE MINIMA - SOLO MENU CONNETTORI                              ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Se vuoi un menu separato "🔗 Connettori" invece di integrarlo
 * nel menu principale, usa questa funzione.
 * Chiamala da onOpen(): aggiungiMenuConnettoriQE();
 */
function aggiungiMenuConnettoriQE() {
  var ui = SpreadsheetApp.getUi();
  
  var menuConnettori = ui.createMenu('🔗 Connettori');
  
  // Aggiungi Fornitori se presente
  if (typeof creaSubmenuConnettoreFornitoriQE === 'function') {
    menuConnettori.addSubMenu(creaSubmenuConnettoreFornitoriQE());
  }
  
  // Placeholder per futuri connettori
  // if (typeof creaSubmenuConnettoreEmailQE === 'function') {
  //   menuConnettori.addSubMenu(creaSubmenuConnettoreEmailQE());
  // }
  
  menuConnettori.addToUi();
}

/**
 * ╔══════════════════════════════════════════════════════════════════════╗
 * ║  WRAPPER RAPIDO PER INTEGRAZIONE                                     ║
 * ╚══════════════════════════════════════════════════════════════════════╝
 * 
 * Chiama questa funzione alla FINE della tua onOpen() esistente
 * per aggiungere automaticamente tutti i connettori disponibili.
 */
function integraConnettoriInMenu(menuPrincipale) {
  // Check se ci sono connettori da aggiungere
  var haConnettori = (typeof creaSubmenuConnettoreFornitoriQE === 'function');
  
  if (!haConnettori) return menuPrincipale;
  
  // Aggiungi separatore e submenu Connettori
  menuPrincipale.addSeparator();
  
  var subMenuConnettori = SpreadsheetApp.getUi().createMenu('🔗 Connettori');
  
  if (typeof creaSubmenuConnettoreFornitoriQE === 'function') {
    subMenuConnettori.addSubMenu(creaSubmenuConnettoreFornitoriQE());
  }
  
  menuPrincipale.addSubMenu(subMenuConnettori);
  
  return menuPrincipale;
}

// ═══════════════════════════════════════════════════════════════════════
// REPORT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera report completo connettore
 */
function generaReportConnettoreFornitoriQE() {
  var report = {
    nome: CONNECTOR_FORNITORI_QE.NOME,
    versione: CONNECTOR_FORNITORI_QE.VERSIONE,
    attivo: isConnectorFornitoriQEAttivo(),
    pronto: checkConnectorFornitoriQEReady(),
    statsSync: getStatsSyncFornitoriQE(),
    timestamp: new Date()
  };
  
  return report;
}