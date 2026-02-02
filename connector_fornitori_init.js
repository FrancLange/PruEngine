/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - INIT v1.0.0
 * ==========================================================================================
 * Inizializzazione, setup e menu per il connettore Fornitori
 * 
 * AUTONOMO: Gestisce il ciclo di vita del connettore
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// INIZIALIZZAZIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Setup iniziale del connettore Fornitori
 * Aggiunge configurazione in SETUP e prepara tab sync
 */
function initConnectorFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "🔗 Inizializza Connector Fornitori",
    "Questa operazione:\n\n" +
    "1. Aggiunge chiavi configurazione in SETUP\n" +
    "2. Crea tab " + CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC + "\n" +
    "3. Prepara il connettore per il sync\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Verifica SETUP esiste
    var setupSheet = ss.getSheetByName("SETUP");
    if (!setupSheet) {
      ui.alert("❌ Errore", "Foglio SETUP non trovato. Esegui prima il setup di Email Engine.", ui.ButtonSet.OK);
      return;
    }
    
    // 2. Aggiungi chiavi in SETUP
    aggiungiChiaviSetupFornitori(setupSheet);
    
    // 3. Crea tab FORNITORI_SYNC (vuota, pronta per sync)
    var syncSheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    if (!syncSheet) {
      syncSheet = ss.insertSheet(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
      syncSheet.setTabColor("#A5D6A7");
      syncSheet.getRange("A1").setValue("In attesa di sync dal master...");
      syncSheet.getRange("A2").setValue("Configura ID master in SETUP, poi esegui Sync.");
      
      logConnectorFornitori("Creata tab FORNITORI_SYNC");
    }
    
    logConnectorFornitori("Connettore inizializzato");
    
    ui.alert(
      "✅ Connettore Inizializzato",
      "Prossimi step:\n\n" +
      "1. Vai nel foglio SETUP\n" +
      "2. Trova: CONNECTOR_FORNITORI_MASTER_ID\n" +
      "3. Incolla l'ID del foglio Fornitori Engine\n" +
      "4. Menu > Connettori > Sync Fornitori\n\n" +
      "💡 Tip: L'ID del foglio è nella URL dopo /d/",
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
    logConnectorFornitori("Errore init: " + e.toString());
  }
}

/**
 * Aggiunge chiavi configurazione in SETUP
 */
function aggiungiChiaviSetupFornitori(setupSheet) {
  var data = setupSheet.getDataRange().getValues();
  var chiavi = data.map(function(row) { return row[0]; });
  
  var nuoveChiavi = [
    ["", ""],
    ["═══ CONNECTOR FORNITORI ═══", ""],
    [CONNECTOR_FORNITORI.SETUP_KEYS.ENABLED, "SI"],
    [CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "Incolla qui ID foglio Fornitori Engine"],
    [CONNECTOR_FORNITORI.SETUP_KEYS.SYNC_INTERVAL, CONNECTOR_FORNITORI.DEFAULTS.SYNC_INTERVAL_MIN],
    [CONNECTOR_FORNITORI.SETUP_KEYS.LAST_SYNC, ""]
  ];
  
  // Aggiungi solo se non esistono già
  var hasConnectorKeys = chiavi.indexOf(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID) >= 0;
  
  if (!hasConnectorKeys) {
    nuoveChiavi.forEach(function(row) {
      setupSheet.appendRow(row);
    });
    
    logConnectorFornitori("Aggiunte chiavi configurazione in SETUP");
  }
}

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea sottomenu Connettori (da chiamare in onOpen)
 * @returns {Object} Submenu object
 */
function creaSubmenuConnettoreFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  return ui.createMenu("📦 Fornitori")
    .addItem("🔗 Inizializza Connettore", "initConnectorFornitori")
    .addSeparator()
    .addItem("🔄 Sync da Master", "syncFornitoriMenu")
    .addItem("⏰ Configura Trigger Sync", "configuraTriggerSyncFornitori")
    .addSeparator()
    .addItem("🔍 Test Lookup Email", "testLookupEmailFornitori")
    .addItem("🧪 Test Connessione Master", "testScritturaMasterFornitori")
    .addSeparator()
    .addItem("📊 Stato Connettore", "mostraStatoConnettoreFornitori");
}

/**
 * Aggiungi menu Connettori al menu principale
 * Da chiamare in onOpen() del file principale
 */
function aggiungiMenuConnettoreFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("🔗 Connettori")
    .addSubMenu(creaSubmenuConnettoreFornitori())
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Menu: Mostra stato connettore
 */
function mostraStatoConnettoreFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var attivo = isConnectorFornitoriAttivo();
  var check = checkConnectorFornitoriReady();
  var stats = getStatsSyncFornitori();
  
  var report = "🔗 CONNECTOR FORNITORI\n";
  report += "══════════════════════════════\n\n";
  
  report += "Stato: " + (attivo ? "✅ ATTIVO" : "❌ NON ATTIVO") + "\n\n";
  
  report += "━━━ CONFIGURAZIONE ━━━\n";
  report += "Master configurato: " + (stats.masterConfigurato ? "✅" : "❌") + "\n";
  report += "Tab sync presente: " + (stats.tabEsiste ? "✅" : "❌") + "\n";
  report += "Fornitori locali: " + stats.righe + "\n\n";
  
  report += "━━━ SYNC ━━━\n";
  report += "Ultimo sync: " + (stats.ultimoSync ? stats.ultimoSync.toLocaleString('it-IT') : "Mai") + "\n";
  report += "Intervallo: " + stats.intervallo + " minuti\n";
  
  if (stats.prossimoSync) {
    report += "Prossimo sync: " + stats.prossimoSync.toLocaleString('it-IT') + "\n";
  }
  
  if (!attivo && !check.canActivate) {
    report += "\n━━━ AZIONI RICHIESTE ━━━\n";
    report += "⚠️ " + check.reason + "\n";
  }
  
  ui.alert("Stato Connettore Fornitori", report, ui.ButtonSet.OK);
}

/**
 * Menu: Test lookup email
 */
function testLookupEmailFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  // Verifica connettore attivo
  if (!isConnectorFornitoriAttivo()) {
    ui.alert(
      "❌ Connettore Non Attivo",
      "Inizializza e configura il connettore prima di eseguire il test.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Chiedi email
  var risposta = ui.prompt(
    "🔍 Test Lookup Email",
    "Inserisci un'email mittente da cercare:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var email = risposta.getResponseText().trim();
  if (!email) {
    ui.alert("Email non valida");
    return;
  }
  
  // Esegui lookup
  var risultato = preCheckEmailFornitore(email);
  
  var report = "🔍 RISULTATO LOOKUP\n";
  report += "══════════════════════════════\n\n";
  report += "Email: " + email + "\n\n";
  
  if (risultato.isFornitoreNoto) {
    report += "✅ FORNITORE TROVATO!\n\n";
    report += "ID: " + risultato.fornitore.id + "\n";
    report += "Nome: " + risultato.fornitore.nome + "\n";
    report += "Status: " + risultato.fornitore.status + "\n";
    report += "Skip L0: " + (risultato.skipL0 ? "SI" : "NO") + "\n\n";
    report += "━━━ CONTESTO L1 ━━━\n";
    report += risultato.contesto;
  } else {
    report += "❌ Nessun fornitore trovato per questa email.\n\n";
    report += "L'email verrà processata normalmente (L0 → L1 → L2 → L3).";
  }
  
  ui.alert("Risultato Lookup", report, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// REPORT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera report completo connettore
 */
function generaReportConnettoreFornitori() {
  var report = {
    nome: CONNECTOR_FORNITORI.NOME,
    versione: CONNECTOR_FORNITORI.VERSIONE,
    attivo: isConnectorFornitoriAttivo(),
    pronto: checkConnectorFornitoriReady(),
    stats: getStatsSyncFornitori(),
    timestamp: new Date()
  };
  
  return report;
}

// ═══════════════════════════════════════════════════════════════════════
// DISATTIVAZIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Disattiva connettore Fornitori
 */
function disattivaConnectorFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "⚠️ Disattiva Connettore",
    "Vuoi disattivare il Connector Fornitori?\n\n" +
    "Il connettore non verrà più usato durante l'analisi email.\n" +
    "I dati locali rimarranno, potrai riattivarlo in qualsiasi momento.",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  // Disattiva in SETUP
  setConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.ENABLED, "NO");
  
  // Rimuovi trigger
  rimuoviTriggerSyncFornitori();
  
  logConnectorFornitori("Connettore disattivato");
  
  ui.alert("✅ Connettore Disattivato", "Puoi riattivarlo modificando ENABLED = SI in SETUP.", ui.ButtonSet.OK);
}

/**
 * Riattiva connettore Fornitori
 */
function riattivaConnectorFornitori() {
  setConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.ENABLED, "SI");
  logConnectorFornitori("Connettore riattivato");
  
  SpreadsheetApp.getUi().alert("✅ Connettore Riattivato");
}

// ═══════════════════════════════════════════════════════════════════════
// PULIZIA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Rimuove completamente il connettore
 * (tab sync, chiavi setup, trigger)
 */
function rimuoviConnectorFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "⚠️ RIMUOVI CONNETTORE",
    "ATTENZIONE: Questa operazione rimuoverà:\n\n" +
    "- Tab " + CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC + "\n" +
    "- Chiavi configurazione da SETUP\n" +
    "- Trigger sync automatico\n\n" +
    "Sei sicuro?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Rimuovi tab sync
    var syncSheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    if (syncSheet) {
      ss.deleteSheet(syncSheet);
    }
    
    // 2. Rimuovi trigger
    rimuoviTriggerSyncFornitori();
    
    // 3. Le chiavi in SETUP le lasciamo (non causano problemi)
    
    logConnectorFornitori("Connettore rimosso completamente");
    
    ui.alert("✅ Connettore Rimosso", "Per reinstallare, esegui di nuovo Inizializza Connettore.", ui.ButtonSet.OK);
    
  } catch(e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
  }
}