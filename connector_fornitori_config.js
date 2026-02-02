/**
 * ==========================================================================================
 * CONNECTOR FORNITORI - CONFIG v1.1.0
 * ==========================================================================================
 * Configurazione connettore per integrazione Fornitori Engine ↔ Email Engine
 * 
 * v1.1.0: Aggiornato mapping colonne per matchare FORNITORI_SYNC reale
 * ==========================================================================================
 */

var CONNECTOR_FORNITORI = {
  
  // ═══════════════════════════════════════════════════════════
  // IDENTIFICATIVO
  // ═══════════════════════════════════════════════════════════
  NOME: "Connector_Fornitori",
  VERSIONE: "1.1.0",
  DESCRIZIONE: "Integrazione Fornitori Engine con Email Engine",
  
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
  // CHIAVI SETUP (in SETUP di Email Engine)
  // ═══════════════════════════════════════════════════════════
  SETUP_KEYS: {
    // ID del foglio Fornitori Engine (master)
    MASTER_SHEET_ID: "CONNECTOR_FORNITORI_MASTER_ID",
    
    // Abilitazione connettore
    ENABLED: "CONNECTOR_FORNITORI_ENABLED",
    
    // Ultimo sync
    LAST_SYNC: "CONNECTOR_FORNITORI_LAST_SYNC",
    
    // Intervallo sync in minuti
    SYNC_INTERVAL: "CONNECTOR_FORNITORI_SYNC_INTERVAL"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    ENABLED: true,
    SYNC_INTERVAL_MIN: 30,  // Ogni 30 minuti
    CACHE_TTL_SEC: 60       // Cache lookup 60 secondi
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE FORNITORI (mapping con nomi reali del foglio)
  // AGGIORNATO v1.1.0 per matchare FORNITORI_SYNC
  // ═══════════════════════════════════════════════════════════
  COLONNE: {
    // Anagrafica (per lookup)
    ID_FORNITORE: "ID_FORNITORE",
    PRIORITA_URGENTE: "PRIORITA",
    NOME_AZIENDA: "RAGIONE_SOCIALE",
    EMAIL_ORDINI: "EMAIL_PRINCIPALE",
    EMAIL_ALTRI: "EMAIL_SECONDARIE",
    DOMINIO_EMAIL: "DOMINIO_EMAIL",
    CONTATTO: "CONTATTO",
    TELEFONO: "TELEFONO",
    
    // Per contesto L1
    SCONTO_PERCENTUALE: "SCONTO_BASE",
    DATA_PROSSIMA_PROMO: "DATA_PROSSIMA_PROMO",
    STATUS_ULTIMA_AZIONE: "STATUS_ULTIMA_AZIONE",
    PERFORMANCE_SCORE: "PERFORMANCE",
    STATUS_FORNITORE: "STATUS",
    
    // Per update
    DATA_ULTIMA_EMAIL: "DATA_ULTIMA_EMAIL",
    DATA_ULTIMO_ORDINE: "ULTIMO_ORDINE",
    DATA_ULTIMA_ANALISI: "DATA_ULTIMA_ANALISI",
    EMAIL_ANALIZZATE_COUNT: "EMAIL_ANALIZZATE_COUNT",
    NOTE_INTERNE: "NOTE_INTERNE",
    ULTIMA_SYNC: "ULTIMA_SYNC"
  },
  
  // ═══════════════════════════════════════════════════════════
  // STATUS VALIDI (per skip L0)
  // ═══════════════════════════════════════════════════════════
  STATUS_SKIP_L0: ["ATTIVO", "IN_VALUTAZIONE", "NUOVO"],
  
  // ═══════════════════════════════════════════════════════════
  // CAMPI PER CONTESTO L1
  // ═══════════════════════════════════════════════════════════
  CAMPI_CONTESTO: [
    "RAGIONE_SOCIALE",
    "PRIORITA",
    "SCONTO_BASE",
    "DATA_PROSSIMA_PROMO",
    "STATUS_ULTIMA_AZIONE",
    "PERFORMANCE"
  ],
  
  // ═══════════════════════════════════════════════════════════
  // LOGGING
  // ═══════════════════════════════════════════════════════════
  LOG_PREFIX: "🔗 [Fornitori]",
  
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
function isConnectorFornitoriAttivo() {
  try {
    // 1. Check se abilitato in SETUP
    var enabled = getConnectorSetupValue(
      CONNECTOR_FORNITORI.SETUP_KEYS.ENABLED, 
      CONNECTOR_FORNITORI.DEFAULTS.ENABLED
    );
    
    if (enabled === "NO" || enabled === false) {
      return false;
    }
    
    // 2. Check se ID master configurato
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    if (!masterId || masterId === "" || masterId.indexOf("Incolla") >= 0) {
      return false;
    }
    
    // 3. Check se tab FORNITORI_SYNC esiste
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC);
    
    return !!sheet;
    
  } catch(e) {
    return false;
  }
}

/**
 * Verifica se il connettore può essere attivato (master configurato)
 * @returns {Object} {canActivate, reason}
 */
function checkConnectorFornitoriReady() {
  var result = { canActivate: false, reason: "" };
  
  try {
    var masterId = getConnectorSetupValue(CONNECTOR_FORNITORI.SETUP_KEYS.MASTER_SHEET_ID, "");
    
    if (!masterId || masterId === "" || masterId.indexOf("Incolla") >= 0) {
      result.reason = "ID foglio Fornitori Master non configurato in SETUP";
      return result;
    }
    
    // Prova ad aprire il master
    try {
      var masterSs = SpreadsheetApp.openById(masterId);
      var masterSheet = masterSs.getSheetByName(CONNECTOR_FORNITORI.SHEETS.FORNITORI_MASTER);
      
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
function getConnectorSetupValue(key, defaultValue) {
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
function setConnectorSetupValue(key, value) {
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
    Logger.log("Errore setConnectorSetupValue: " + e.toString());
  }
}

/**
 * Log specifico per connettore
 * @param {String} messaggio
 */
function logConnectorFornitori(messaggio) {
  var fullMsg = CONNECTOR_FORNITORI.LOG_PREFIX + " " + messaggio;
  
  // Usa logSistema se esiste (da Email Engine)
  if (typeof logSistema === 'function') {
    logSistema(fullMsg);
  } else {
    Logger.log(fullMsg);
  }
}