/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - SYNC v1.0.0
 * ==========================================================================================
 * Sincronizzazione dati da Fornitori Engine (Master) → Email Engine (Tab locale)
 * 
 * AUTONOMO: Copia completa tab FORNITORI dal master al locale
 * ==========================================================================================
 */

/**
 * MAIN: Sincronizza tab FORNITORI dal master
 * @param {Boolean} silent - Se true, non mostra alert
 * @returns {Object} {success, righe, durata, errore}
 */
function syncFornitoriDaMaster(silent) {
  var startTime = new Date();
  var risultato = { 
    success: false, 
    righe: 0, 
    durata: 0, 
    errore: null 
  };
  
  logConnectorFornitori("Sync START");
  
  try {
    // 1. Ottieni ID master
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    
    if (!masterId || masterId === "") {
      throw new Error("ID Fornitori Master non configurato in SETUP");
    }
    
    // 2. Apri foglio master
    var masterSs;
    try {
      masterSs = SpreadsheetApp.openById(masterId);
    } catch(e) {
      throw new Error("Impossibile aprire master. Verifica ID e permessi: " + e.message);
    }
    
    var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_MASTER);
    if (!masterSheet) {
      throw new Error("Tab 'FORNITORI' non trovata nel master");
    }
    
    // 3. Leggi tutti i dati
    var masterData = masterSheet.getDataRange().getValues();
    
    if (masterData.length <= 1) {
      logConnectorFornitori("⚠️ Tab master vuota o solo header");
      risultato.success = true;
      risultato.righe = 0;
      return risultato;
    }
    
    logConnectorFornitori("Letti " + (masterData.length - 1) + " fornitori da master");
    
    // 4. Prepara tab locale
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var localSheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    
    // Crea se non esiste
    if (!localSheet) {
      localSheet = ss.insertSheet(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
      localSheet.setTabColor("#A5D6A7"); // Verde chiaro
      logConnectorFornitori("Creata tab FORNITORI_SYNC");
    }
    
    // 5. Pulisci e scrivi
    localSheet.clear();
    
    if (masterData.length > 0 && masterData[0].length > 0) {
      localSheet.getRange(1, 1, masterData.length, masterData[0].length)
        .setValues(masterData);
    }
    
    // 6. Formattazione
    localSheet.setFrozenRows(1);
    localSheet.getRange(1, 1, 1, masterData[0].length)
      .setFontWeight("bold")
      .setBackground("#E8F5E9");
    
    // 7. Aggiorna tracking
    setConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.LAST_SYNC, new Date().toISOString());
    
    // 8. Invalida cache
    CONNECTOR_FORNITORI._cache.fornitori = null;
    CONNECTOR_FORNITORI._cache.timestamp = null;
    
    // Risultato
    risultato.success = true;
    risultato.righe = masterData.length - 1; // Escludi header
    risultato.durata = Math.round((new Date() - startTime) / 1000);
    
    logConnectorFornitori("✅ Sync OK: " + risultato.righe + " fornitori in " + risultato.durata + "s");
    
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "✅ Sync Fornitori Completato",
        "Sincronizzati " + risultato.righe + " fornitori\n" +
        "Durata: " + risultato.durata + " secondi\n\n" +
        "Fonte: " + masterSs.getName(),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch(e) {
    risultato.errore = e.toString();
    logConnectorFornitori("❌ Sync ERRORE: " + risultato.errore);
    
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        "❌ Errore Sync Fornitori",
        risultato.errore,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
  
  return risultato;
}

/**
 * Sync silenzioso (per trigger automatici)
 */
function syncFornitoriSilent() {
  return syncFornitoriDaMaster(true);
}

/**
 * Sync da menu
 */
function syncFornitoriMenu() {
  return syncFornitoriDaMaster(false);
}

/**
 * Verifica se il sync è necessario
 * @param {Number} maxMinuti - Minuti massimi dal ultimo sync (default da config)
 * @returns {Boolean}
 */
function isSyncFornitoriNecessario(maxMinuti) {
  maxMinuti = maxMinuti || getConnectorSetupValue(
    CONNECTOR_FORNITORI.SETUP_KEYS.SYNC_INTERVAL,
    CONNECTOR_FORNITORI.DEFAULTS.SYNC_INTERVAL_MIN
  );
  
  var lastSync = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.LAST_SYNC, null);
  
  if (!lastSync) {
    return true; // Mai sincronizzato
  }
  
  try {
    var lastDate = new Date(lastSync);
    var now = new Date();
    var diffMinuti = (now - lastDate) / (1000 * 60);
    
    return diffMinuti >= maxMinuti;
    
  } catch(e) {
    return true; // In caso di errore, sincronizza
  }
}

/**
 * Auto-sync se necessario (chiamato prima di lookup)
 * @returns {Boolean} true se sync eseguito o non necessario
 */
function autoSyncFornitoriSeNecessario() {
  // Non sincronizzare se connettore non attivo
  if (!isConnectorFornitoriAttivo()) {
    return false;
  }
  
  if (isSyncFornitoriNecessario()) {
    logConnectorFornitori("Auto-sync necessario");
    var result = syncFornitoriDaMaster(true);
    return result.success;
  }
  
  return true; // Non necessario, dati freschi
}

// ═══════════════════════════════════════════════════════════════════════
// TRIGGER AUTOMATICO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Configura trigger per sync automatico
 */
function configuraTriggerSyncFornitori() {
  var ui = SpreadsheetApp.getUi();
  
  // Verifica connettore pronto
  var check = checkConnectorFornitoriReady();
  if (!check.canActivate) {
    ui.alert("❌ Connettore Non Pronto", check.reason, ui.ButtonSet.OK);
    return;
  }
  
  // Rimuovi trigger esistenti
  rimuoviTriggerSyncFornitori();
  
  // Intervallo da SETUP o default
  var intervallo = parseInt(getConnectorSetupValue(
    CONNECTOR_FORNITORI.SETUP_KEYS.SYNC_INTERVAL,
    CONNECTOR_FORNITORI.DEFAULTS.SYNC_INTERVAL_MIN
  ));
  
  // Crea trigger
  ScriptApp.newTrigger('syncFornitoriSilent')
    .timeBased()
    .everyMinutes(intervallo)
    .create();
  
  logConnectorFornitori("Trigger sync configurato: ogni " + intervallo + " minuti");
  
  ui.alert(
    "✅ Trigger Configurato",
    "Sync automatico ogni " + intervallo + " minuti.\n\n" +
    "Verifica in: Estensioni > Apps Script > Trigger",
    ui.ButtonSet.OK
  );
}

/**
 * Rimuove trigger sync fornitori
 */
function rimuoviTriggerSyncFornitori() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  
  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === 'syncFornitoriSilent' || 
        handler === 'syncFornitoriDaMaster' ||
        handler === 'syncFornitoriMenu') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  if (count > 0) {
    logConnectorFornitori("Rimossi " + count + " trigger sync");
  }
}

// ═══════════════════════════════════════════════════════════════════════
// STATISTICHE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Ottiene statistiche sync
 * @returns {Object}
 */
function getStatsSyncFornitori() {
  var stats = {
    tabEsiste: false,
    righe: 0,
    ultimoSync: null,
    intervallo: CONNECTOR_FORNITORI.DEFAULTS.SYNC_INTERVAL_MIN,
    prossimoSync: null,
    masterConfigurato: false
  };
  
  try {
    // Tab locale
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    
    if (sheet) {
      stats.tabEsiste = true;
      stats.righe = Math.max(0, sheet.getLastRow() - 1);
    }
    
    // Ultimo sync
    var lastSync = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.LAST_SYNC, null);
    if (lastSync) {
      stats.ultimoSync = new Date(lastSync);
      
      var intervallo = parseInt(getConnectorSetupValue(
        CONNECTOR_FORNITORI.SETUP_KEYS.SYNC_INTERVAL,
        CONNECTOR_FORNITORI.DEFAULTS.SYNC_INTERVAL_MIN
      ));
      stats.intervallo = intervallo;
      
      stats.prossimoSync = new Date(stats.ultimoSync.getTime() + intervallo * 60 * 1000);
    }
    
    // Master
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    stats.masterConfigurato = !!(masterId && masterId.length > 10);
    
  } catch(e) {
    // Ignora errori
  }
  
  return stats;
}