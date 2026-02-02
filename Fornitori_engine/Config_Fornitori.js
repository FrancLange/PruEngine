/**
 * ==========================================================================================
 * CONFIG - MOTORE FORNITORI v1.1.0
 * ==========================================================================================
 * Gestione anagrafica fornitori con integrazione Email Engine + Questioni Engine
 * 
 * CHANGELOG v1.1.0:
 * - Aggiunte 5 colonne per integrazione Questioni Engine (33-37)
 * ==========================================================================================
 */

const CONFIG_FORNITORI = {
  VERSION: "1.1.0",
  NOME_SISTEMA: "Fornitori Intelligence Engine",
  DATA_RELEASE: "2025-02-02",
  
  // ═══════════════════════════════════════════════════════════
  // FOGLI GOOGLE SHEETS
  // ═══════════════════════════════════════════════════════════
  SHEETS: {
    HOME: "HOME",
    FORNITORI: "FORNITORI",
    SETUP: "SETUP",
    PROMPTS: "PROMPTS",
    LOG_SISTEMA: "LOG_SISTEMA",
    STORICO_AZIONI: "STORICO_AZIONI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE FORNITORI (37 colonne)
  // ═══════════════════════════════════════════════════════════
  COLONNE_FORNITORI: {
    // ANAGRAFICA BASE (1-7)
    ID_FORNITORE: "ID_FORNITORE",
    PRIORITA_URGENTE: "PRIORITA_URGENTE",
    NOME_AZIENDA: "NOME_AZIENDA",
    EMAIL_ORDINI: "EMAIL_ORDINI",
    EMAIL_ALTRI: "EMAIL_ALTRI",
    CONTATTO: "CONTATTO",
    TELEFONO: "TELEFONO",
    
    // DATI FISCALI (8-10)
    PARTITA_IVA: "PARTITA_IVA",
    INDIRIZZO: "INDIRIZZO",
    NOTE: "NOTE",
    
    // OPERATIVO (11-15)
    METODO_INVIO: "METODO_INVIO",
    TIPO_LISTINO: "TIPO_LISTINO",
    LINK_LISTINO: "LINK_LISTINO",
    MIN_ORDINE: "MIN_ORDINE",
    LEAD_TIME_GIORNI: "LEAD_TIME_GIORNI",
    
    // PROMO & SCONTI (16-24)
    PROMO_SU_RICHIESTA: "PROMO_SU_RICHIESTA",
    DATA_PROSSIMA_PROMO: "DATA_PROSSIMA_PROMO",
    PROMO_ANNUALI_NUM: "PROMO_ANNUALI_NUM",
    PROMO_ANNUALI_RICHIESTE: "PROMO_ANNUALI_RICHIESTE",
    SCONTO_TARGET: "SCONTO_TARGET",
    RICHIESTA_SCONTO: "RICHIESTA_SCONTO",
    ESITO_RICHIESTA_SCONTO: "ESITO_RICHIESTA_SCONTO",
    TIPO_SCONTO_CONCESSO: "TIPO_SCONTO_CONCESSO",
    SCONTO_PERCENTUALE: "SCONTO_PERCENTUALE",
    
    // TRACKING EMAIL (25-32)
    DATA_ULTIMO_ORDINE: "DATA_ULTIMO_ORDINE",
    DATA_ULTIMA_EMAIL: "DATA_ULTIMA_EMAIL",
    DATA_ULTIMA_ANALISI: "DATA_ULTIMA_ANALISI",
    STATUS_ULTIMA_AZIONE: "STATUS_ULTIMA_AZIONE",
    EMAIL_ANALIZZATE_COUNT: "EMAIL_ANALIZZATE_COUNT",
    PERFORMANCE_SCORE: "PERFORMANCE_SCORE",
    STATUS_FORNITORE: "STATUS_FORNITORE",
    DATA_CREAZIONE: "DATA_CREAZIONE",
    
    // ═══ QUESTIONI ENGINE (33-37) ═══
    QUESTIONI_APERTE: "QUESTIONI_APERTE",
    QUESTIONI_TOTALI: "QUESTIONI_TOTALI",
    ULTIMA_QUESTIONE: "ULTIMA_QUESTIONE",
    TIPO_QUESTIONI: "TIPO_QUESTIONI",
    NOTE_QUESTIONI: "NOTE_QUESTIONI"
  },
  
  // ═══════════════════════════════════════════════════════════
  // COLONNE STORICO AZIONI
  // ═══════════════════════════════════════════════════════════
  COLONNE_STORICO: {
    TIMESTAMP: "TIMESTAMP",
    ID_FORNITORE: "ID_FORNITORE",
    NOME_FORNITORE: "NOME_FORNITORE",
    TIPO_AZIONE: "TIPO_AZIONE",
    DESCRIZIONE: "DESCRIZIONE",
    EMAIL_COLLEGATA: "EMAIL_COLLEGATA",
    ESEGUITO_DA: "ESEGUITO_DA",
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
  // ENUM VALUES
  // ═══════════════════════════════════════════════════════════
  
  // Status Fornitore
  STATUS_FORNITORE: {
    ATTIVO: "ATTIVO",
    SOSPESO: "SOSPESO",
    IN_VALUTAZIONE: "IN_VALUTAZIONE",
    BLACKLIST: "BLACKLIST",
    NUOVO: "NUOVO"
  },
  
  // Metodi Invio Ordine
  METODO_INVIO: {
    SITO: "SITO",
    MODULO: "MODULO",
    LISTA_EMAIL: "LISTA_EMAIL",
    EMAIL_DIRETTA: "EMAIL_DIRETTA",
    TELEFONO: "TELEFONO",
    WHATSAPP: "WHATSAPP"
  },
  
  // Tipo Listino
  TIPO_LISTINO: {
    SITO_WEB: "SITO_WEB",
    EXCEL: "EXCEL",
    PDF: "PDF",
    NO_LISTINO: "NO_LISTINO",
    SU_RICHIESTA: "SU_RICHIESTA"
  },
  
  // Esito Richiesta Sconto
  ESITO_SCONTO: {
    IN_ATTESA: "IN_ATTESA",
    ACCETTATO: "ACCETTATO",
    RIFIUTATO: "RIFIUTATO",
    CONTROPROPOSTA: "CONTROPROPOSTA",
    NON_RICHIESTO: "NON_RICHIESTO"
  },
  
  // Tipo Azioni (per storico)
  TIPO_AZIONE: {
    CREAZIONE: "CREAZIONE",
    MODIFICA: "MODIFICA",
    EMAIL_RICEVUTA: "EMAIL_RICEVUTA",
    EMAIL_INVIATA: "EMAIL_INVIATA",
    ORDINE: "ORDINE",
    RICHIESTA_SCONTO: "RICHIESTA_SCONTO",
    PROMO_ATTIVATA: "PROMO_ATTIVATA",
    PROBLEMA: "PROBLEMA",
    NOTA_MANUALE: "NOTA_MANUALE",
    // ═══ NUOVI TIPI PER QUESTIONI ═══
    QUESTIONE_APERTA: "QUESTIONE_APERTA",
    QUESTIONE_CHIUSA: "QUESTIONE_CHIUSA"
  },
  
  // ═══════════════════════════════════════════════════════════
  // DEFAULTS
  // ═══════════════════════════════════════════════════════════
  DEFAULTS: {
    STATUS_INIZIALE: "NUOVO",
    PERFORMANCE_SCORE: 3,
    SCONTO_TARGET: 10,
    MIN_ORDINE: 0,
    LEAD_TIME_GIORNI: 5,
    // ═══ DEFAULTS QUESTIONI ═══
    QUESTIONI_APERTE: 0,
    QUESTIONI_TOTALI: 0
  },
  
  // ═══════════════════════════════════════════════════════════
  // CONFIGURAZIONE LOOKUP EMAIL
  // ═══════════════════════════════════════════════════════════
  EMAIL_LOOKUP: {
    // Se true, cerca anche nei domini (es. @biocosmetics.it)
    MATCH_DOMINIO: true,
    // Colonne da cercare per match email
    COLONNE_EMAIL: ["EMAIL_ORDINI", "EMAIL_ALTRI"],
    // Cache TTL in secondi (0 = no cache)
    CACHE_TTL: 300
  },
  
  // ═══════════════════════════════════════════════════════════
  // CONTESTO PER EMAIL ENGINE
  // ═══════════════════════════════════════════════════════════
  CONTESTO_EMAIL: {
    // Campi da includere nel contesto per L1
    CAMPI_CONTESTO: [
      "NOME_AZIENDA",
      "PRIORITA_URGENTE", 
      "SCONTO_PERCENTUALE",
      "DATA_PROSSIMA_PROMO",
      "STATUS_ULTIMA_AZIONE",
      "PERFORMANCE_SCORE",
      // ═══ CONTESTO QUESTIONI ═══
      "QUESTIONI_APERTE"
    ]
  },
  
  // ═══════════════════════════════════════════════════════════
  // CONFIGURAZIONE QUESTIONI ENGINE (NUOVO)
  // ═══════════════════════════════════════════════════════════
  QUESTIONI: {
    // Colonne per stats (devono matchare CONNECTOR_FORNITORI_QE)
    COLONNE_STATS: [
      "QUESTIONI_APERTE",
      "QUESTIONI_TOTALI", 
      "ULTIMA_QUESTIONE",
      "TIPO_QUESTIONI",
      "NOTE_QUESTIONI"
    ],
    // Categorie questioni (per breakdown)
    CATEGORIE: {
      RECLAMO: "RECLAMO",
      ORDINE: "ORDINE",
      FATTURA: "FATTURA",
      SPEDIZIONE: "SPEDIZIONE",
      QUALITA: "QUALITA",
      URGENTE: "URGENTE",
      ALTRO: "ALTRO"
    }
  }
};