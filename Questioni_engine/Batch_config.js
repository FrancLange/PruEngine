/**
 * ==========================================================================================
 * BATCH_CONFIG.js - Configurazione Infrastruttura Batch v1.0.0
 * ==========================================================================================
 * Configurazione centralizzata per OpenAI Batch API
 * Risparmio 50% sui costi API con finestra 24h
 * ==========================================================================================
 */

var BATCH_CONFIG = {
  
  // ═══════════════════════════════════════════════════════════
  // IDENTIFICATIVO
  // ═══════════════════════════════════════════════════════════
  NOME: "Batch Infrastructure",
  VERSIONE: "1.0.0",
  
  // ═══════════════════════════════════════════════════════════
  // OPENAI BATCH API
  // ═══════════════════════════════════════════════════════════
  API: {
    BASE_URL: "https://api.openai.com/v1",
    ENDPOINTS: {
      FILES: "/files",
      BATCHES: "/batches"
    },
    COMPLETION_WINDOW: "24h",  // Obbligatorio per sconto 50%
    MAX_REQUESTS_PER_BATCH: 50000,
    MAX_FILE_SIZE_MB: 100
  },
  
  // ═══════════════════════════════════════════════════════════
  // FOGLI
  // ═══════════════════════════════════════════════════════════
  SHEETS: {
    BATCH_QUEUE: "BATCH_QUEUE",   // Richieste in attesa
    BATCH_JOBS: "BATCH_JOBS"      // Job inviati a OpenAI
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE BATCH_QUEUE (richieste singole)
  // ═══════════════════════════════════════════════════════════
  COLONNE_QUEUE: {
    ID_REQUEST: "ID_REQUEST",           // UUID richiesta
    TIMESTAMP_CREAZIONE: "TIMESTAMP_CREAZIONE",
    TIPO_OPERAZIONE: "TIPO_OPERAZIONE", // QUESTIONI_CLUSTER, QUESTIONI_MATCH, EMAIL_L1, etc.
    CUSTOM_ID: "CUSTOM_ID",             // ID per ricollego (es. email ID, fornitore ID)
    SYSTEM_PROMPT: "SYSTEM_PROMPT",
    USER_PROMPT: "USER_PROMPT",
    MODEL: "MODEL",
    MAX_TOKENS: "MAX_TOKENS",
    TEMPERATURE: "TEMPERATURE",
    JSON_MODE: "JSON_MODE",             // SI/NO
    PRIORITA: "PRIORITA",               // 1=Alta, 2=Media, 3=Bassa
    STATUS: "STATUS",                   // PENDING, QUEUED, PROCESSING, COMPLETED, ERROR
    BATCH_JOB_ID: "BATCH_JOB_ID",       // Riferimento a BATCH_JOBS
    RISULTATO_JSON: "RISULTATO_JSON",   // Risposta AI
    TIMESTAMP_COMPLETAMENTO: "TIMESTAMP_COMPLETAMENTO",
    ERRORE: "ERRORE",
    METADATA_JSON: "METADATA_JSON"      // Dati extra per post-processing
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE BATCH_JOBS (job inviati a OpenAI)
  // ═══════════════════════════════════════════════════════════
  COLONNE_JOBS: {
    BATCH_JOB_ID: "BATCH_JOB_ID",       // ID interno (BJB-001, BJB-002)
    OPENAI_BATCH_ID: "OPENAI_BATCH_ID", // ID restituito da OpenAI
    OPENAI_FILE_ID: "OPENAI_FILE_ID",   // ID file JSONL uploadato
    TIMESTAMP_INVIO: "TIMESTAMP_INVIO",
    TIMESTAMP_COMPLETAMENTO: "TIMESTAMP_COMPLETAMENTO",
    TIPO_OPERAZIONE: "TIPO_OPERAZIONE", // Tipo prevalente nel batch
    NUM_REQUESTS: "NUM_REQUESTS",       // Quante richieste nel batch
    STATUS: "STATUS",                   // UPLOADING, SUBMITTED, IN_PROGRESS, COMPLETED, FAILED, EXPIRED
    OPENAI_STATUS: "OPENAI_STATUS",     // Status raw da OpenAI
    OUTPUT_FILE_ID: "OUTPUT_FILE_ID",   // ID file risultati
    ERROR_FILE_ID: "ERROR_FILE_ID",     // ID file errori (se presente)
    COMPLETED_COUNT: "COMPLETED_COUNT",
    FAILED_COUNT: "FAILED_COUNT",
    ERRORE_MSG: "ERRORE_MSG",
    COSTO_STIMATO: "COSTO_STIMATO",
    NOTE: "NOTE"
  },
  
  // ═══════════════════════════════════════════════════════════
  // STATUS
  // ═══════════════════════════════════════════════════════════
  STATUS_QUEUE: {
    PENDING: "PENDING",       // In attesa di essere inclusa in un batch
    QUEUED: "QUEUED",         // Assegnata a un batch non ancora inviato
    PROCESSING: "PROCESSING", // Batch inviato, in elaborazione
    COMPLETED: "COMPLETED",   // Risultato ricevuto
    ERROR: "ERROR"            // Errore
  },
  
  STATUS_JOB: {
    UPLOADING: "UPLOADING",     // File in upload
    SUBMITTED: "SUBMITTED",     // Batch creato, in attesa
    IN_PROGRESS: "IN_PROGRESS", // OpenAI sta elaborando
    COMPLETED: "COMPLETED",     // Completato con successo
    FAILED: "FAILED",           // Fallito
    EXPIRED: "EXPIRED",         // Scaduto (oltre 24h)
    CANCELLED: "CANCELLED"      // Cancellato manualmente
  },
  
  // ═══════════════════════════════════════════════════════════
  // TIPI OPERAZIONE
  // ═══════════════════════════════════════════════════════════
  TIPI_OPERAZIONE: {
    // Questioni
    QUESTIONI_CLUSTER: "QUESTIONI_CLUSTER",
    QUESTIONI_MATCH: "QUESTIONI_MATCH",
    
    // Email Engine (futuro)
    EMAIL_L0_SPAM: "EMAIL_L0_SPAM",
    EMAIL_L1_ANALISI: "EMAIL_L1_ANALISI",
    EMAIL_L2_VERIFICA: "EMAIL_L2_VERIFICA",
    
    // Generici
    CUSTOM: "CUSTOM"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    MODEL: "gpt-4o",
    MODEL_MINI: "gpt-4o-mini",
    TEMPERATURE: 0.3,
    MAX_TOKENS: 500,
    JSON_MODE: true,
    
    // Batching
    MIN_REQUESTS_PER_BATCH: 5,    // Minimo richieste prima di inviare
    MAX_REQUESTS_PER_BATCH: 100,  // Massimo per batch (nostro limite)
    MAX_WAIT_MINUTES: 60,         // Attesa massima prima di forzare invio
    POLL_INTERVAL_MINUTES: 15,    // Ogni quanto controllare status
    
    // Priorità
    PRIORITA_ALTA: 1,
    PRIORITA_MEDIA: 2,
    PRIORITA_BASSA: 3
  },
  
  // ═══════════════════════════════════════════════════════════
  // COSTI STIMATI (per 1M token, sconto 50% già applicato)
  // ═══════════════════════════════════════════════════════════
  COSTI: {
    "gpt-4o": {
      input: 1.25,    // $2.50/M → $1.25/M con batch
      output: 5.00    // $10/M → $5/M con batch
    },
    "gpt-4o-mini": {
      input: 0.075,   // $0.15/M → $0.075/M con batch
      output: 0.30    // $0.60/M → $0.30/M con batch
    }
  },
  
  // ═══════════════════════════════════════════════════════════
  // SCHEDULING
  // ═══════════════════════════════════════════════════════════
  SCHEDULING: {
    // Orari preferiti per invio batch (minor carico)
    ORE_INVIO_PREFERITE: [2, 3, 4, 19, 20, 21, 22, 23],
    
    // Trigger automatici
    TRIGGER_INVIO_BATCH: "inviaQueuedBatches",
    TRIGGER_POLL_STATUS: "pollBatchStatus",
    TRIGGER_PROCESS_RESULTS: "processCompletedBatches"
  },
  
  // ═══════════════════════════════════════════════════════════
  // LOG
  // ═══════════════════════════════════════════════════════════
  LOG_PREFIX: "📦 [Batch]",
  
  // ═══════════════════════════════════════════════════════════
  // CACHE RUNTIME
  // ═══════════════════════════════════════════════════════════
  _cache: {
    queue: null,
    jobs: null,
    timestamp: null
  }
};

// ═══════════════════════════════════════════════════════════════════════
// UTILITY LOGGING
// ═══════════════════════════════════════════════════════════════════════

/**
 * Log specifico per batch
 */
function logBatch(messaggio) {
  var fullMsg = BATCH_CONFIG.LOG_PREFIX + " " + messaggio;
  
  if (typeof logSistema === 'function') {
    logSistema(fullMsg);
  } else {
    Logger.log(fullMsg);
  }
}

/**
 * Genera UUID per richieste
 */
function generaUUID() {
  return Utilities.getUuid();
}

/**
 * Genera ID incrementale per batch job
 */
function generaBatchJobId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return "BJB-001";
  }
  
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 1).getValue();
  
  if (lastId && lastId.toString().indexOf("BJB-") === 0) {
    var num = parseInt(lastId.replace("BJB-", "")) + 1;
    return "BJB-" + ("000" + num).slice(-3);
  }
  
  return "BJB-" + ("000" + lastRow).slice(-3);
}