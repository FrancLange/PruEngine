/**
 * ==========================================================================================
 * QUESTIONI_CONFIG.js - Configurazione Motore Questioni v1.0.0 STANDALONE
 * ==========================================================================================
 * Configurazione centralizzata per gestione questioni/ticket fornitori
 * STANDALONE: Non dipende da altri CONFIG, funziona su foglio separato
 * ==========================================================================================
 */

var QUESTIONI_CONFIG = {
  
  // ═══════════════════════════════════════════════════════════
  // SISTEMA
  // ═══════════════════════════════════════════════════════════
  VERSION: "1.0.0",
  NOME_SISTEMA: "Questioni Engine",
  DATA_RELEASE: "2025-02-02",
  
  // ═══════════════════════════════════════════════════════════
  // FOGLI
  // ═══════════════════════════════════════════════════════════
  SHEETS: {
    // Tab locali
    EMAIL_IN: "EMAIL_IN",
    QUESTIONI: "QUESTIONI",
    PROMPTS: "PROMPTS",
    SETUP: "SETUP",
    LOG_SISTEMA: "LOG_SISTEMA",
    
    // Tab future
    CAMPAGNE_IN: "CAMPAGNE_IN",
    STORICO_AZIONI: "STORICO_AZIONI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // CHIAVI SETUP
  // ═══════════════════════════════════════════════════════════
  SETUP_KEYS: {
    // Connector Email
    EMAIL_ENGINE_ID: "CONNECTOR_EMAIL_ENGINE_ID",
    EMAIL_SYNC_ENABLED: "CONNECTOR_EMAIL_ENABLED",
    EMAIL_LAST_SYNC: "CONNECTOR_EMAIL_LAST_SYNC",
    EMAIL_MAX_RIGHE: "CONNECTOR_EMAIL_MAX_RIGHE",
    EMAIL_GIORNI_INDIETRO: "CONNECTOR_EMAIL_GIORNI",
    EMAIL_FILTRO_SCORE: "CONNECTOR_EMAIL_FILTRO_SCORE",
    
    // Connector Campagne (futuro)
    CAMPAGNE_ENGINE_ID: "CONNECTOR_CAMPAGNE_ENGINE_ID",
    CAMPAGNE_SYNC_ENABLED: "CONNECTOR_CAMPAGNE_ENABLED",
    
    // AI
    OPENAI_THRESHOLD_MATCH: "AI_THRESHOLD_MATCH",
    OPENAI_THRESHOLD_CLUSTER: "AI_THRESHOLD_CLUSTER",
    
    // Sistema
    AUTO_CREA_QUESTIONI: "AUTO_CREA_QUESTIONI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    EMAIL_MAX_RIGHE: 2000,
    EMAIL_GIORNI_INDIETRO: 90,
    EMAIL_FILTRO_SCORE: 70,      // score_problema minimo
    SYNC_INTERVAL_MIN: 30,
    CACHE_TTL_SEC: 60,
    MATCH_THRESHOLD: 70,
    CLUSTER_THRESHOLD: 80,
    AUTO_CREA_QUESTIONI: true
  },
  
  // ═══════════════════════════════════════════════════════════
  // AI CONFIG
  // ═══════════════════════════════════════════════════════════
  AI_CONFIG: {
    MODEL_CLUSTER: "gpt-4o",
    MODEL_MATCH: "gpt-4o-mini",
    TEMPERATURE_CLUSTER: 0.3,
    TEMPERATURE_MATCH: 0.2,
    MAX_TOKENS_CLUSTER: 800,
    MAX_TOKENS_MATCH: 300
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE EMAIL_IN (sync da Email Engine LOG_IN)
  // ═══════════════════════════════════════════════════════════
  COLONNE_EMAIL_IN: {
    // Anagrafica
    TIMESTAMP_SYNC: "TIMESTAMP_SYNC",
    ID_EMAIL: "ID_EMAIL",
    TIMESTAMP_ORIGINALE: "TIMESTAMP_ORIGINALE",
    MITTENTE: "MITTENTE",
    OGGETTO: "OGGETTO",
    CORPO: "CORPO",
    
    // Analisi (da Email Engine)
    L3_TAGS_FINALI: "L3_TAGS_FINALI",
    L3_SINTESI_FINALE: "L3_SINTESI_FINALE",
    L3_SCORES_JSON: "L3_SCORES_JSON",
    L3_CONFIDENCE: "L3_CONFIDENCE",
    L3_AZIONI_SUGGERITE: "L3_AZIONI_SUGGERITE",
    STATUS_EMAIL: "STATUS_EMAIL",
    
    // Questioni (locale)
    ID_QUESTIONE: "ID_QUESTIONE",
    STATUS_QUESTIONE: "STATUS_QUESTIONE",
    PROCESSATO: "PROCESSATO"
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE QUESTIONI
  // ═══════════════════════════════════════════════════════════
  COLONNE_QUESTIONI: {
    // Identificazione
    ID_QUESTIONE: "ID_QUESTIONE",
    TIMESTAMP_CREAZIONE: "TIMESTAMP_CREAZIONE",
    FONTE: "FONTE",
    
    // Fornitore
    ID_FORNITORE: "ID_FORNITORE",
    NOME_FORNITORE: "NOME_FORNITORE",
    EMAIL_FORNITORE: "EMAIL_FORNITORE",
    
    // Contenuto
    TITOLO: "TITOLO",
    DESCRIZIONE: "DESCRIZIONE",
    CATEGORIA: "CATEGORIA",
    TAGS: "TAGS",
    
    // Priorità e Stato
    PRIORITA: "PRIORITA",
    URGENTE: "URGENTE",
    STATUS: "STATUS",
    
    // Tracking
    DATA_SCADENZA: "DATA_SCADENZA",
    ASSEGNATO_A: "ASSEGNATO_A",
    
    // Email collegate
    EMAIL_COLLEGATE: "EMAIL_COLLEGATE",
    NUM_EMAIL: "NUM_EMAIL",
    
    // Risoluzione
    TIMESTAMP_RISOLUZIONE: "TIMESTAMP_RISOLUZIONE",
    RISOLUZIONE_NOTE: "RISOLUZIONE_NOTE",
    
    // AI
    AI_CONFIDENCE: "AI_CONFIDENCE",
    AI_CLUSTER_ID: "AI_CLUSTER_ID"
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE PROMPTS
  // ═══════════════════════════════════════════════════════════
  COLONNE_PROMPTS: {
    KEY: "CHIAVE_PROMPT",
    LIVELLO_AI: "LIVELLO_AI",
    TESTO: "TESTO_PROMPT",
    CATEGORIA: "CATEGORIA",
    INPUT_ATTESO: "INPUT_ATTESO",
    OUTPUT_ATTESO: "OUTPUT_ATTESO",
    TEMPERATURA: "TEMPERATURA",
    MAX_TOKENS: "MAX_TOKENS",
    ATTIVO: "ATTIVO",
    NOTE: "NOTE"
  },
  
  // ═══════════════════════════════════════════════════════════
  // ENUM STATUS QUESTIONI
  // ═══════════════════════════════════════════════════════════
  STATUS_QUESTIONE: {
    APERTA: "APERTA",
    IN_LAVORAZIONE: "IN_LAVORAZIONE",
    IN_ATTESA_RISPOSTA: "IN_ATTESA_RISPOSTA",
    RISOLTA: "RISOLTA",
    CHIUSA: "CHIUSA",
    ANNULLATA: "ANNULLATA"
  },
  
  // ═══════════════════════════════════════════════════════════
  // ENUM PRIORITÀ
  // ═══════════════════════════════════════════════════════════
  PRIORITA: {
    CRITICA: "CRITICA",
    ALTA: "ALTA",
    MEDIA: "MEDIA",
    BASSA: "BASSA"
  },
  
  // ═══════════════════════════════════════════════════════════
  // ENUM CATEGORIE
  // ═══════════════════════════════════════════════════════════
  CATEGORIE_QUESTIONE: {
    PROBLEMA_ORDINE: "PROBLEMA_ORDINE",
    PROBLEMA_QUALITA: "PROBLEMA_QUALITA",
    PROBLEMA_CONSEGNA: "PROBLEMA_CONSEGNA",
    PROBLEMA_FATTURAZIONE: "PROBLEMA_FATTURAZIONE",
    RICHIESTA_SCONTO: "RICHIESTA_SCONTO",
    RICHIESTA_INFO: "RICHIESTA_INFO",
    RECLAMO: "RECLAMO",
    URGENZA: "URGENZA",
    ALTRO: "ALTRO"
  },
  
  // ═══════════════════════════════════════════════════════════
  // ENUM FONTI
  // ═══════════════════════════════════════════════════════════
  FONTI_QUESTIONE: {
    EMAIL: "EMAIL",
    CAMPAGNA: "CAMPAGNA",
    MANUALE: "MANUALE",
    IMPORT: "IMPORT"
  },
  
  // ═══════════════════════════════════════════════════════════
  // PROMPT KEYS
  // ═══════════════════════════════════════════════════════════
  PROMPT_KEYS: {
    CLUSTER_EMAIL: "QUESTIONI_CLUSTER_EMAIL",
    MATCH_QUESTIONE: "QUESTIONI_MATCH_ESISTENTE",
    GENERA_TITOLO: "QUESTIONI_GENERA_TITOLO",
    ANALIZZA_URGENZA: "QUESTIONI_ANALIZZA_URGENZA"
  },
  
  // ═══════════════════════════════════════════════════════════
  // MAPPING COLONNE EMAIL ENGINE → EMAIL_IN
  // Per il sync, mappa le colonne sorgente
  // ═══════════════════════════════════════════════════════════
  MAPPING_EMAIL_ENGINE: {
    // Colonna sorgente (Email Engine) → Colonna destinazione (EMAIL_IN)
    "TIMESTAMP": "TIMESTAMP_ORIGINALE",
    "ID_EMAIL": "ID_EMAIL",
    "MITTENTE": "MITTENTE",
    "OGGETTO": "OGGETTO",
    "CORPO": "CORPO",
    "L3_TAGS_FINALI": "L3_TAGS_FINALI",
    "L3_SINTESI_FINALE": "L3_SINTESI_FINALE",
    "L3_SCORES_JSON": "L3_SCORES_JSON",
    "L3_CONFIDENCE": "L3_CONFIDENCE",
    "L3_AZIONI_SUGGERITE": "L3_AZIONI_SUGGERITE",
    "STATUS": "STATUS_EMAIL"
  },
  
  // ═══════════════════════════════════════════════════════════
  // LOGGING
  // ═══════════════════════════════════════════════════════════
  LOG_PREFIX: "🎫 [Questioni]"
};

// ═══════════════════════════════════════════════════════════════════════
// UTILITY GLOBALI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Log nel foglio LOG_SISTEMA
 */
function logQuestioni(messaggio) {
  var fullMsg = QUESTIONI_CONFIG.LOG_PREFIX + " " + messaggio;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.LOG_SISTEMA);
    if (sheet) {
      sheet.appendRow([new Date(), fullMsg]);
    }
  } catch(e) {
    Logger.log(fullMsg);
  }
}

/**
 * Genera ID univoco per questione
 */
function generaIdQuestione() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return "QST-001";
  }
  
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 1).getValue();
  
  if (lastId && typeof lastId === 'string' && lastId.indexOf("QST-") === 0) {
    var num = parseInt(lastId.replace("QST-", "")) + 1;
    return "QST-" + ("000" + num).slice(-3);
  }
  
  return "QST-" + ("000" + lastRow).slice(-3);
}

/**
 * Ottieni valore da SETUP
 */
function getQuestioniSetupValue(key, defaultValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.SETUP);
    
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
 * Imposta valore in SETUP
 */
function setQuestioniSetupValue(key, value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.SETUP);
    
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
    
    if (!found) {
      sheet.appendRow([key, value]);
    }
    
  } catch(e) {
    Logger.log("Errore setQuestioniSetupValue: " + e.toString());
  }
}