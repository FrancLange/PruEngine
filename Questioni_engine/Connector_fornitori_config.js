/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - CONFIG v1.0.0 (per Questioni Engine)
 * ==========================================================================================
 * Configurazione connettore per integrazione Fornitori Engine ↔ Questioni Engine
 * 
 * RICICLATO: Basato su connector fornitori di Email Engine, adattato per Questioni
 * ==========================================================================================
 */

var CONNECTOR_FORNITORI_QE = {
  
  // ═══════════════════════════════════════════════════════════
  // IDENTIFICATIVO
  // ═══════════════════════════════════════════════════════════
  NOME: "Connector_Fornitori_Questioni",
  VERSIONE: "1.0.0",
  DESCRIZIONE: "Integrazione Fornitori Engine con Questioni Engine",
  
  // ═══════════════════════════════════════════════════════════
  // CONFIGURAZIONE FOGLI
  // ═══════════════════════════════════════════════════════════
  SHEETS: {
    // Tab locale (copia sync)
    FORNITORI_SYNC: "FORNITORI_SYNC",
    
    // Tab nel master (fonte verità)
    FORNITORI_MASTER: "FORNITORI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // CHIAVI SETUP (in SETUP di Questioni Engine)
  // ═══════════════════════════════════════════════════════════
  SETUP_KEYS: {
    // ID del foglio Fornitori Engine (master)
    MASTER_SHEET_ID: "CONNECTOR_FORNITORI_MASTER_ID",
    
    // Abilitazione connettore
    ENABLED: "CONNECTOR_FORNITORI_ENABLED",
    
    // Ultimo sync
    LAST_SYNC: "CONNECTOR_FORNITORI_LAST_SYNC",
    
    // Intervallo sync in minuti
    SYNC_INTERVAL: "CONNECTOR_FORNITORI_SYNC_INTERVAL",
    
    // Ultimo update stats verso master
    LAST_STATS_UPDATE: "CONNECTOR_FORNITORI_LAST_STATS"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    ENABLED: true,
    SYNC_INTERVAL_MIN: 60,   // Ogni ora (meno frequente di Email)
    CACHE_TTL_SEC: 120,      // Cache lookup 2 minuti
    STATS_UPDATE_INTERVAL: 6 // Ore tra update stats
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE FORNITORI (mapping)
  // Deve matchare Config_Fornitori.gs del Fornitori Engine
  // ═══════════════════════════════════════════════════════════
  COLONNE: {
    // Anagrafica (per lookup)
    ID_FORNITORE: "ID_FORNITORE",
    PRIORITA_URGENTE: "PRIORITA_URGENTE",
    NOME_AZIENDA: "NOME_AZIENDA",
    EMAIL_ORDINI: "EMAIL_ORDINI",
    EMAIL_ALTRI: "EMAIL_ALTRI",
    CONTATTO: "CONTATTO",
    TELEFONO: "TELEFONO",
    
    // Per contesto questione
    SCONTO_PERCENTUALE: "SCONTO_PERCENTUALE",
    STATUS_ULTIMA_AZIONE: "STATUS_ULTIMA_AZIONE",
    PERFORMANCE_SCORE: "PERFORMANCE_SCORE",
    STATUS_FORNITORE: "STATUS_FORNITORE",
    
    // 🆕 COLONNE QUESTIONI (da aggiungere a Fornitori Engine)
    QUESTIONI_APERTE: "QUESTIONI_APERTE",
    QUESTIONI_TOTALI: "QUESTIONI_TOTALI",
    ULTIMA_QUESTIONE: "ULTIMA_QUESTIONE",
    TIPO_QUESTIONI: "TIPO_QUESTIONI",
    NOTE_QUESTIONI: "NOTE_QUESTIONI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // MAPPING CATEGORIE QUESTIONI
  // ═══════════════════════════════════════════════════════════
  CATEGORIE_QUESTIONI: {
    RECLAMO: "RECLAMO",
    ORDINE: "ORDINE",
    FATTURA: "FATTURA",
    SPEDIZIONE: "SPEDIZIONE",
    QUALITA: "QUALITA",
    URGENTE: "URGENTE",
    ALTRO: "ALTRO"
  },
  
  // ═══════════════════════════════════════════════════════════
  // STATUS QUESTIONI (per calcolo stats)
  // ═══════════════════════════════════════════════════════════
  STATUS_QUESTIONI: {
    APERTA: "APERTA",
    IN_GESTIONE: "IN_GESTIONE",
    IN_ATTESA: "IN_ATTESA",
    RISOLTA: "RISOLTA",
    CHIUSA: "CHIUSA"
  },
  
  // Status che contano come "aperte"
  STATUS_APERTE: ["APERTA", "IN_GESTIONE", "IN_ATTESA"],
  
  // ═══════════════════════════════════════════════════════════
  // LOGGING
  // ═══════════════════════════════════════════════════════════
  LOG_PREFIX: "🔗 [Fornitori→QE]",
  
  // ═══════════════════════════════════════════════════════════
  // CACHE RUNTIME (per batch processing)
  // ═══════════════════════════════════════════════════════════
  _cache: {
    fornitori: null,
    timestamp: null,
    colMap: null
  }
};

// ═══════════════════════════════════════════════════════════════════════
// UTILITY CONFIG
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se il connettore è attivo e configurato
 * @returns {Boolean}
 */
function isConnectorFornitoriQEAttivo() {
  try {
    // 1. Check se abilitato in SETUP
    var enabled = getConnectorQESetupValue(
      CONNECTOR_FORNITORI_QE.SETUP_KEYS.ENABLED, 
      CONNECTOR_FORNITORI_QE.DEFAULTS.ENABLED
    );
    
    if (enabled === "NO" || enabled === false) {
      return false;
    }
    
    // 2. Check se ID master configurato
    var masterId = getConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.MASTER_SHEET_ID, "");
    if (!masterId || masterId === "" || masterId.indexOf("Incolla") >= 0) {
      return false;
    }
    
    // 3. Check se tab FORNITORI_SYNC esiste
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_SYNC);
    
    return !!sheet;
    
  } catch(e) {
    return false;
  }
}

/**
 * Verifica se il connettore può essere attivato (master configurato)
 * @returns {Object} {canActivate, reason}
 */
function checkConnectorFornitoriQEReady() {
  var result = { canActivate: false, reason: "" };
  
  try {
    var masterId = getConnectorQESetupValue(CONNECTOR_FORNITORI_QE.SETUP_KEYS.MASTER_SHEET_ID, "");
    
    if (!masterId || masterId === "" || masterId.indexOf("Incolla") >= 0) {
      result.reason = "ID foglio Fornitori Master non configurato in SETUP";
      return result;
    }
    
    // Prova ad aprire il master
    try {
      var masterSs = SpreadsheetApp.openById(masterId);
      var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI_QE.SHEETS.FORNITORI_MASTER);
      
      if (!masterSheet) {
        result.reason = "Tab FORNITORI non trovato nel foglio master";
        return result;
      }
      
      result.canActivate = true;
      result.reason = "OK - Master raggiungibile";
      
    } catch(e) {
      result.reason = "Impossibile accedere al foglio master: " + e.message;
    }
    
  } catch(e) {
    result.reason = "Errore verifica: " + e.message;
  }
  
  return result;
}

/**
 * Ottiene valore da SETUP per il connettore
 * @param {String} key - Chiave
 * @param {*} defaultValue - Valore default
 * @returns {*}
 */
function getConnectorQESetupValue(key, defaultValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SETUP");
    
    if (!sheet) return defaultValue;
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        var val = data[i][1];
        return (val !== "" && val !== null && val !== undefined) ? val : defaultValue;
      }
    }
    
    return defaultValue;
    
  } catch(e) {
    return defaultValue;
  }
}

/**
 * Imposta valore in SETUP per il connettore
 * @param {String} key - Chiave
 * @param {*} value - Valore
 */
function setConnectorQESetupValue(key, value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SETUP");
    
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    var found = false;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    
    // Se non trovato, aggiungi
    if (!found) {
      sheet.appendRow([key, value]);
    }
    
  } catch(e) {
    Logger.log("Errore setConnectorQESetupValue: " + e.toString());
  }
}

/**
 * Log specifico per connettore
 * @param {String} messaggio
 */
function logConnectorFornitoriQE(messaggio) {
  var fullMsg = CONNECTOR_FORNITORI_QE.LOG_PREFIX + " " + messaggio;
  
  // Usa logQuestioni se esiste
  if (typeof logQuestioni === 'function') {
    logQuestioni(fullMsg);
  } else if (typeof logSistema === 'function') {
    logSistema(fullMsg);
  } else {
    Logger.log(fullMsg);
  }
}