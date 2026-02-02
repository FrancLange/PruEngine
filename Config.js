/**
 * ==========================================================================================
 * CONFIG - MOTORE EMAIL AI v1.1.0
 * ==========================================================================================
 * Sistema di analisi email con AI stratificato (L0 Spam Filter + GPT → Claude → Merge)
 * ==========================================================================================
 */

const CONFIG = {
  VERSION: "1.1.0",
  NOME_SISTEMA: "Email Intelligence Engine",
  DATA_RELEASE: "2025-01-31",
  
  // ═══════════════════════════════════════════════════════════
  // AI MODELS
  // ═══════════════════════════════════════════════════════════
  AI_CONFIG: {
    MODEL_SPAM_FILTER: "gpt-4o-mini",
    L0_TEMPERATURE: 0.1,
    L0_MAX_TOKENS: 100,
    L0_CONFIDENCE_THRESHOLD: 90,
    
    MODEL_GPT: "gpt-4o",
    L1_TEMPERATURE: 0.3,
    L1_MAX_TOKENS: 500,
    
    MODEL_CLAUDE: "claude-3-5-sonnet-20241022",
    L2_TEMPERATURE: 0.2,
    L2_MAX_TOKENS: 600,
    
    CONFIDENCE_THRESHOLD: 70,
    DIVERGENCE_THRESHOLD: 30
  },
  
  // ═══════════════════════════════════════════════════════════
  // FOGLI GOOGLE SHEETS
  // ═══════════════════════════════════════════════════════════
  SHEETS: {
    HOME: "HOME",
    SETUP: "SETUP",
    AUTOMAZIONI: "AUTOMAZIONI",
    PROMPTS: "PROMPTS",
    LOG_IN: "LOG_IN",
    LOG_OUT: "LOG_OUT",
    MANIFEST: "MANIFEST",
    CHANGELOG: "CHANGELOG",
    LOG_SISTEMA: "LOG_SISTEMA"
  },
  
  // ═══════════════════════════════════════════════════════════
  // CHIAVI SETUP
  // ═══════════════════════════════════════════════════════════
  KEYS_SETUP: {
    ID_MASTERSKU: "ID_MASTERSKU_SOURCE",
    ID_FORNITORI: "ID_FORNITORI_SOURCE",
    ID_OUTPUT_BI: "ID_OUTPUT_BI",
    CONFIDENCE_THRESHOLD: "AI_CONFIDENCE_THRESHOLD",
    DIVERGENCE_THRESHOLD: "AI_DIVERGENCE_THRESHOLD",
    DISPATCHER_ULTIMO_RUN: "DISPATCHER_ULTIMO_RUN",
    DISPATCHER_STATUS: "DISPATCHER_STATUS",
    DISPATCHER_INTERVALLO_MIN: "DISPATCHER_INTERVALLO_MIN"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    CONFIDENCE_THRESHOLD: 70,
    DIVERGENCE_THRESHOLD: 30,
    DISPATCHER_INTERVALLO_MIN: 60,
    TIMEOUT_SEC: 300,
    MAX_RETRY: 3,
    RETRY_DELAY_SEC: 60
  },
  
  // ═══════════════════════════════════════════════════════════
  // LIVELLI AI
  // ═══════════════════════════════════════════════════════════
  LIVELLI_AI: {
    L0_SPAM_FILTER: "L0_SPAM_FILTER",
    L1_GPT: "L1_GPT",
    L2_CLAUDE: "L2_CLAUDE",
    L3_MERGE: "L3_MERGE",
    L4_HUMAN: "L4_HUMAN",
    L5_AI_REFINED: "L5_AI_REFINED"
  },
  
  // ═══════════════════════════════════════════════════════════
  // CATEGORIE JOB
  // ═══════════════════════════════════════════════════════════
  CATEGORIE_JOB: {
    SISTEMA: "SISTEMA",
    EMAIL: "EMAIL",
    AI: "AI",
    SYNC: "SYNC",
    BI: "BI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // FREQUENZE
  // ═══════════════════════════════════════════════════════════
  FREQUENZE: {
    ORARIA: "ORARIA",
    GIORNALIERA: "GIORNALIERA",
    SETTIMANALE: "SETTIMANALE",
    CUSTOM: "CUSTOM"
  },
  
  // ═══════════════════════════════════════════════════════════
  // ESITI
  // ═══════════════════════════════════════════════════════════
  ESITI: {
    OK: "OK",
    ERRORE: "ERRORE",
    SKIP: "SKIP",
    IN_CORSO: "IN_CORSO",
    TIMEOUT: "TIMEOUT",
    WAITING: "WAITING"
  },
  
  // ═══════════════════════════════════════════════════════════
  // SCORE CATEGORIES
  // ═══════════════════════════════════════════════════════════
  SCORE_CATEGORIES: {
    NOVITA: "novita",
    PROMO: "promo",
    FATTURA: "fattura",
    CATALOGO: "catalogo",
    PREZZI: "prezzi",
    PROBLEMA: "problema",
    RISPOSTA_CAMPAGNA: "risposta_campagna",
    URGENTE: "urgente",
    ORDINE: "ordine"
  },
  
  // ═══════════════════════════════════════════════════════════
  // SPAM CATEGORIES
  // ═══════════════════════════════════════════════════════════
  SPAM_CATEGORIES: {
    NEWSLETTER_GENERICA: "Newsletter marketing generica",
    PROMO_IRRILEVANTE: "Promozione servizi non pertinenti",
    PHISHING: "Tentativo phishing/truffa",
    AUTO_REPLY: "Auto-reply/Out of office",
    NOTIFICA_SISTEMA: "Notifica automatica sistema",
    LINGUA_SBAGLIATA: "Lingua non italiano/inglese",
    DUPLICATO: "Email duplicata",
    SPAM_GENERICO: "Spam generico"
  },
  
  // ═══════════════════════════════════════════════════════════
  // STATUS EMAIL
  // ═══════════════════════════════════════════════════════════
  STATUS_EMAIL: {
    DA_FILTRARE: "DA_FILTRARE",
    SPAM: "SPAM",
    DA_ANALIZZARE: "DA_ANALIZZARE",
    IN_ANALISI: "IN_ANALISI",
    ANALIZZATO: "ANALIZZATO",
    NEEDS_REVIEW: "NEEDS_REVIEW",
    GESTITO: "GESTITO"
  },
  
  // ═══════════════════════════════════════════════════════════
  // PROMPTS SISTEMA
  // ═══════════════════════════════════════════════════════════
  PROMPTS_SISTEMA: {
    L0_SPAM_SYSTEM: "Sei un filtro anti-spam specializzato in email B2B commerciali.\n\n" +
      "Il tuo compito e' classificare email come SPAM o LEGIT.\n\n" +
      "SPAM include:\n" +
      "- Newsletter marketing generiche non richieste\n" +
      "- Promozioni servizi non pertinenti al business\n" +
      "- Tentativi phishing/truffa\n" +
      "- Auto-reply e out-of-office\n" +
      "- Notifiche automatiche di sistema\n" +
      "- Email in lingue diverse da italiano/inglese\n" +
      "- Email duplicate\n\n" +
      "LEGIT include:\n" +
      "- Email da fornitori commerciali conosciuti\n" +
      "- Conferme ordini, fatture, listini prezzi\n" +
      "- Comunicazioni su prodotti/servizi\n" +
      "- Risposte a campagne marketing nostre\n" +
      "- Problemi/reclami clienti\n" +
      "- Qualsiasi email con contenuto commerciale rilevante\n\n" +
      "IMPORTANTE: In caso di dubbio, classifica come LEGIT.\n" +
      "Meglio analizzare un'email in piu' che perderne una importante.\n\n" +
      "Rispondi SOLO in formato JSON:\n" +
      "{\n" +
      '  "isSpam": true/false,\n' +
      '  "confidence": 0-100,\n' +
      '  "reason": "Motivo breve se spam, altrimenti null"\n' +
      "}",
    
    L0_SPAM_USER: "Classifica questa email:\n\n" +
      "MITTENTE: {{mittente}}\n" +
      "OGGETTO: {{oggetto}}\n" +
      "CORPO (prime 500 caratteri):\n" +
      "{{corpo_preview}}\n\n" +
      "Rispondi SOLO con il JSON, senza altri commenti."
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE AUTOMAZIONI
  // ═══════════════════════════════════════════════════════════
  COLONNE_AUTOMAZIONI: {
    ID_JOB: "ID_JOB",
    ATTIVA: "ATTIVA",
    NOME: "NOME",
    CATEGORIA: "CATEGORIA",
    CODICE_FUNZIONE: "CODICE_FUNZIONE",
    FREQUENZA: "FREQUENZA",
    INTERVALLO_ORE: "INTERVALLO_ORE",
    ORA_ESECUZIONE: "ORA_ESECUZIONE",
    DIPENDE_DA: "DIPENDE_DA",
    PRIORITA: "PRIORITA",
    TIMEOUT_SEC: "TIMEOUT_SEC",
    MAX_RETRY: "MAX_RETRY",
    PARAMETRI_JSON: "PARAMETRI_JSON",
    ULTIMA_EXEC: "ULTIMA_EXEC",
    PROSSIMA_EXEC: "PROSSIMA_EXEC",
    ESITO: "ESITO",
    DURATA_SEC: "DURATA_SEC",
    ERRORE_MSG: "ERRORE_MSG",
    EXEC_COUNT: "EXEC_COUNT",
    SUCCESS_RATE: "SUCCESS_RATE",
    NOTE: "NOTE"
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
  // COLONNE LOG_IN (31 colonne)
  // ═══════════════════════════════════════════════════════════
  COLONNE_LOG_IN: {
    TIMESTAMP: "TIMESTAMP",
    ID_EMAIL: "ID_EMAIL",
    MITTENTE: "MITTENTE",
    OGGETTO: "OGGETTO",
    CORPO: "CORPO",
    
    L0_TIMESTAMP: "L0_TIMESTAMP",
    L0_IS_SPAM: "L0_IS_SPAM",
    L0_CONFIDENCE: "L0_CONFIDENCE",
    L0_REASON: "L0_REASON",
    L0_MODELLO: "L0_MODELLO",
    
    L1_TIMESTAMP: "L1_TIMESTAMP",
    L1_TAGS: "L1_TAGS",
    L1_SINTESI: "L1_SINTESI",
    L1_SCORES_JSON: "L1_SCORES_JSON",
    L1_MODELLO: "L1_MODELLO",
    
    L2_TIMESTAMP: "L2_TIMESTAMP",
    L2_TAGS: "L2_TAGS",
    L2_SINTESI: "L2_SINTESI",
    L2_SCORES_JSON: "L2_SCORES_JSON",
    L2_CONFIDENCE: "L2_CONFIDENCE",
    L2_RICHIESTA_RETRY: "L2_RICHIESTA_RETRY",
    L2_MODELLO: "L2_MODELLO",
    
    L3_TIMESTAMP: "L3_TIMESTAMP",
    L3_TAGS_FINALI: "L3_TAGS_FINALI",
    L3_SINTESI_FINALE: "L3_SINTESI_FINALE",
    L3_SCORES_JSON: "L3_SCORES_JSON",
    L3_CONFIDENCE: "L3_CONFIDENCE",
    L3_DIVERGENZA: "L3_DIVERGENZA",
    L3_NEEDS_REVIEW: "L3_NEEDS_REVIEW",
    L3_AZIONI_SUGGERITE: "L3_AZIONI_SUGGERITE",
    
    STATUS: "STATUS"
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE LOG_OUT
  // ═══════════════════════════════════════════════════════════
  COLONNE_LOG_OUT: {
    TIMESTAMP: "Data Creazione",
    ID_EMAIL: "ID_Email",
    DESTINATARIO: "Email Destinatario",
    OGGETTO: "Oggetto",
    CORPO: "Corpo Email",
    ALLEGATI: "Link Allegati",
    LIVELLO_GENERAZIONE: "Livello_Generazione",
    PROMPT_USATO: "Prompt_Usato",
    MODELLO_AI: "Modello_AI",
    STATUS: "Status Invio"
  }
};