/**
 * ==========================================================================================
 * SETUP MOTORE EMAIL AI v1.0.0
 * ==========================================================================================
 * Inizializzazione completa con un solo clic
 * ==========================================================================================
 */

function onOpen_PATCHED() {
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('🤖 Email Intelligence v1.2')
    .addItem('🛠️ Setup Completo', 'setupCompleto')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔑 API Keys')
      .addItem('OpenAI', 'impostaApiKeyOpenAI')
      .addItem('Claude', 'impostaApiKeyClaude')
      .addItem('🧪 Test Connessioni', 'testConnessioniAI'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Analisi Email')
      .addItem('🛡️ Test Layer 0 (Spam Filter)', 'testLayer0SpamFilter')
      .addItem('🧪 Test Layer 1 (GPT)', 'testLayer1')
      .addItem('🧪 Test Analisi Completa', 'testAnalisiSingola')
      .addItem('▶️ Analizza Email in Coda', 'menuAnalizzaEmailInCoda'));
  
  // 🆕 PATCH: Aggiungi submenu Connettori
  if (typeof creaSubmenuConnettoreFornitori === 'function') {
    menu.addSeparator()
        .addSubMenu(ui.createMenu('🔗 Connettori')
          .addSubMenu(creaSubmenuConnettoreFornitori()));
  }
  
  menu.addSeparator()
      .addItem('⏰ Configura Trigger', 'configuraTrigger')
      .addItem('ℹ️ Info & Changelog', 'showChangelog')
      .addToUi();
}


// ═══════════════════════════════════════════════════════════════════════
// SETUP COMPLETO
// ═══════════════════════════════════════════════════════════════════════

function setupCompleto() {
  const ui = SpreadsheetApp.getUi();
  
  const risposta = ui.alert(
    '🛠️ Setup Completo Motore Email AI',
    'Questa operazione creerà:\n\n' +
    '✅ 9 Fogli preconfigurati\n' +
    '✅ 5 Automazioni predefinite\n' +
    '✅ 8 Prompt stratificati\n' +
    '✅ Collegamenti file esterni\n' +
    '✅ Manifest & Changelog\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Crea Fogli
    creaFogli(ss);
    
    // 2. Popola SETUP
    popolaSetup(ss);
    
    // 3. Popola AUTOMAZIONI
    popolaAutomazioni(ss);
    
    // 4. Popola PROMPTS
    popolaPrompts(ss);
    
    // 5. Crea MANIFEST
    creaManifest(ss);
    
    // 6. Crea CHANGELOG
    creaChangelog(ss);
    
    // 7. Email di test
    creaEmailTest(ss);
    
    logSistema("Setup completo v1.0.0 eseguito con successo");
    
    ui.alert(
      '✅ Sistema Pronto!',
      'Setup completato:\n\n' +
      '📁 9 Fogli creati\n' +
      '🤖 5 Automazioni configurate\n' +
      '📝 8 Prompt stratificati\n' +
      '📧 1 Email di test\n\n' +
      'Prossimi step:\n' +
      '1. Configura API Keys (Menu)\n' +
      '2. Imposta ID file esterni (SETUP)\n' +
      '3. Test: Analisi Email > Test Singola',
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore Setup', e.toString(), ui.ButtonSet.OK);
    logSistema("ERRORE Setup: " + e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE FOGLI
// ═══════════════════════════════════════════════════════════════════════

function creaFogli(ss) {
  const fogli = [
    { nome: CONFIG.SHEETS.HOME, colore: "#134F5C" },
    { nome: CONFIG.SHEETS.SETUP, colore: "#666666" },
    { nome: CONFIG.SHEETS.AUTOMAZIONI, colore: "#9B59B6" },
    { nome: CONFIG.SHEETS.PROMPTS, colore: "#F1C232" },
    { nome: CONFIG.SHEETS.LOG_IN, colore: "#93C47D" },
    { nome: CONFIG.SHEETS.LOG_OUT, colore: "#E06666" },
    { nome: CONFIG.SHEETS.MANIFEST, colore: "#3D85C6" },
    { nome: CONFIG.SHEETS.CHANGELOG, colore: "#FF9900" },
    { nome: CONFIG.SHEETS.LOG_SISTEMA, colore: "#000000" }
  ];
  
  fogli.forEach(({ nome, colore }) => {
    let sheet = ss.getSheetByName(nome);
    if (!sheet) {
      sheet = ss.insertSheet(nome);
      sheet.setTabColor(colore);
      Logger.log(`✅ Creato: ${nome}`);
    }
    
    // Headers specifici
    if (nome === CONFIG.SHEETS.AUTOMAZIONI) {
      const headers = Object.values(CONFIG.COLONNE_AUTOMAZIONI);
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
      sheet.setFrozenRows(1);
    }
    
    if (nome === CONFIG.SHEETS.PROMPTS) {
      const headers = Object.values(CONFIG.COLONNE_PROMPTS);
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
      sheet.setFrozenRows(1);
    }
    
    if (nome === CONFIG.SHEETS.LOG_IN) {
      const headers = Object.values(CONFIG.COLONNE_LOG_IN);
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(4);
    }
    
    if (nome === CONFIG.SHEETS.LOG_OUT) {
      const headers = Object.values(CONFIG.COLONNE_LOG_OUT);
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
      sheet.setFrozenRows(1);
    }
    
    if (nome === CONFIG.SHEETS.LOG_SISTEMA) {
      sheet.getRange("A1:B1")
        .setValues([["TIMESTAMP", "MESSAGGIO"]])
        .setFontWeight("bold")
        .setBackground("#EFEFEF");
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// POPOLAMENTO SETUP
// ═══════════════════════════════════════════════════════════════════════

function popolaSetup(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SETUP);
  if (sheet.getLastRow() > 1) return; // Già popolato
  
  sheet.getRange("A1:B1")
    .setValues([["CHIAVE", "VALORE"]])
    .setFontWeight("bold")
    .setBackground("#EFEFEF");
  
  const dati = [
    ["=== FILE ESTERNI ===", ""],
    [CONFIG.KEYS_SETUP.ID_MASTERSKU, "Incolla qui ID MasterSku"],
    [CONFIG.KEYS_SETUP.ID_FORNITORI, "Incolla qui ID Fornitori"],
    [CONFIG.KEYS_SETUP.ID_OUTPUT_BI, ""],
    ["", ""],
    ["=== CONFIGURAZIONE AI ===", ""],
    [CONFIG.KEYS_SETUP.CONFIDENCE_THRESHOLD, CONFIG.DEFAULTS.CONFIDENCE_THRESHOLD],
    [CONFIG.KEYS_SETUP.DIVERGENCE_THRESHOLD, CONFIG.DEFAULTS.DIVERGENCE_THRESHOLD],
    ["", ""],
    ["=== DISPATCHER ===", ""],
    [CONFIG.KEYS_SETUP.DISPATCHER_INTERVALLO_MIN, CONFIG.DEFAULTS.DISPATCHER_INTERVALLO_MIN],
    [CONFIG.KEYS_SETUP.DISPATCHER_ULTIMO_RUN, ""],
    [CONFIG.KEYS_SETUP.DISPATCHER_STATUS, ""]
  ];
  
  sheet.getRange(2, 1, dati.length, 2).setValues(dati);
  
  // Formattazione
  sheet.getRange("A:A").setFontWeight("bold");
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 400);
}

// ═══════════════════════════════════════════════════════════════════════
// POPOLAMENTO AUTOMAZIONI
// ═══════════════════════════════════════════════════════════════════════

function popolaAutomazioni(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.AUTOMAZIONI);
  if (sheet.getLastRow() > 1) return;
  
  const jobs = [
    {
      id: "JOB-001",
      attiva: "SI",
      nome: "Health Check Sistema",
      categoria: "SISTEMA",
      codice: "healthCheckSistema",
      frequenza: "ORARIA",
      intervallo: 1,
      ora: "*",
      dipende: "",
      priorita: 1,
      timeout: 60,
      retry: 1,
      parametri: "{}",
      note: "Verifica API, spazio log, errori"
    },
    {
      id: "JOB-002",
      attiva: "SI",
      nome: "Analisi Email Layer 1 (GPT)",
      categoria: "EMAIL",
      codice: "analizzaEmailL1",
      frequenza: "ORARIA",
      intervallo: 1,
      ora: "*",
      dipende: "",
      priorita: 20,
      timeout: 300,
      retry: 3,
      parametri: '{"max_email": 10}',
      note: "Prima elaborazione GPT-4o"
    },
    {
      id: "JOB-003",
      attiva: "SI",
      nome: "Analisi Email Layer 2 (Claude)",
      categoria: "AI",
      codice: "analizzaEmailL2",
      frequenza: "ORARIA",
      intervallo: 1,
      ora: "*",
      dipende: "JOB-002",
      priorita: 21,
      timeout: 300,
      retry: 3,
      parametri: '{"max_email": 10}',
      note: "Verifica e raffinamento Claude"
    },
    {
      id: "JOB-004",
      attiva: "SI",
      nome: "Merge Analisi Layer 3",
      categoria: "AI",
      codice: "mergeAnalisiL3",
      frequenza: "ORARIA",
      intervallo: 1,
      ora: "*",
      dipende: "JOB-003",
      priorita: 22,
      timeout: 120,
      retry: 2,
      parametri: "{}",
      note: "Merge finale + confidence"
    },
    {
      id: "JOB-005",
      attiva: "NO",
      nome: "Sync File Esterni",
      categoria: "SYNC",
      codice: "syncFileEsterni",
      frequenza: "GIORNALIERA",
      intervallo: "",
      ora: "05:00",
      dipende: "",
      priorita: 5,
      timeout: 300,
      retry: 2,
      parametri: "{}",
      note: "Sincronizza MasterSku, Fornitori"
    }
  ];
  
  jobs.forEach(job => {
    const row = [
      job.id, job.attiva, job.nome, job.categoria, job.codice,
      job.frequenza, job.intervallo, job.ora, job.dipende, job.priorita,
      job.timeout, job.retry, job.parametri,
      "", "", "", "", "", "", "", // Tracking (vuoti)
      job.note
    ];
    sheet.appendRow(row);
  });
  
  Logger.log("✅ Automazioni: " + jobs.length + " job configurati");
}

// ═══════════════════════════════════════════════════════════════════════
// POPOLAMENTO PROMPTS STRATIFICATI
// ═══════════════════════════════════════════════════════════════════════

function popolaPrompts(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROMPTS);
  if (sheet.getLastRow() > 1) return;
  
  const prompts = [
    // LAYER 1 - GPT Prima Elaborazione
    {
      key: "L1_ANALISI_EMAIL",
      livello: "L1_GPT",
      testo: `Sei un analista esperto di email commerciali B2B.
Analizza l'email e rispondi SOLO in formato JSON valido con questa struttura esatta:
{
  "tags": ["tag1", "tag2"],
  "sintesi": "Breve sintesi in italiano (max 100 parole)",
  "scores": {
    "novita": 0,
    "promo": 0,
    "fattura": 0,
    "catalogo": 0,
    "prezzi": 0,
    "problema": 0,
    "risposta_campagna": 0,
    "urgente": 0,
    "ordine": 0
  }
}

I punteggi vanno da 0 a 100. Assegna punteggi alti solo se l'evidenza è chiara.
Tags possibili: NOVITA, PROMO, FATTURA, CATALOGO, PREZZI, PROBLEMA, RISPOSTA_CAMPAGNA, URGENTE, ORDINE, INFO, CONFERMA, LISTINO`,
      categoria: "ANALISI_EMAIL",
      input: "Email grezza (mittente, oggetto, corpo)",
      output: "JSON con tags, sintesi, scores",
      temp: 0.3,
      tokens: 500,
      attivo: "SI",
      note: "Prima categorizzazione rapida"
    },
    
    // LAYER 2 - Claude Verifica
    {
      key: "L2_VERIFICA_EMAIL",
      livello: "L2_CLAUDE",
      testo: `Sei un revisore esperto di analisi email commerciali B2B.
Ti viene fornita un'email e l'analisi preliminare fatta da un altro sistema.
Il tuo compito è:
1. Verificare l'accuratezza dell'analisi
2. Correggere eventuali errori
3. Assegnare un punteggio di confidence (0-100) alla tua analisi
4. Indicare se hai bisogno di una terza opinione (richiestaRetry)

Rispondi SOLO in formato JSON valido:
{
  "tags": ["tag1", "tag2"],
  "sintesi": "Sintesi raffinata (max 100 parole)",
  "scores": {
    "novita": 0,
    "promo": 0,
    "fattura": 0,
    "catalogo": 0,
    "prezzi": 0,
    "problema": 0,
    "risposta_campagna": 0,
    "urgente": 0,
    "ordine": 0
  },
  "confidence": 85,
  "richiestaRetry": false,
  "note": "Eventuali note sulla revisione"
}`,
      categoria: "ANALISI_EMAIL",
      input: "Email + Analisi Layer 1",
      output: "JSON raffinato + confidence",
      temp: 0.2,
      tokens: 600,
      attivo: "SI",
      note: "Verifica e raffinamento"
    },
    
    // LAYER 3 - Prompts per azioni
    {
      key: "L3_GENERA_RISPOSTA_PROBLEMA",
      livello: "L3_MERGE",
      testo: `Genera una risposta professionale per un problema segnalato da un fornitore.

DATI DISPONIBILI:
- Email originale: {{email_corpo}}
- Fornitore: {{nome_fornitore}}
- Analisi AI: {{sintesi_problema}}

ISTRUZIONI:
- Tono professionale ma empatico
- Conferma ricezione problema
- Richiedi dettagli se necessario
- Proponi tempistica di risposta
- Firma: Team Yumibio

Rispondi SOLO con:
OGGETTO: [oggetto email]
CORPO: [corpo email]`,
      categoria: "GENERAZIONE_EMAIL",
      input: "Analisi merged + dati fornitore",
      output: "Email formattata (oggetto + corpo)",
      temp: 0.7,
      tokens: 400,
      attivo: "SI",
      note: "Per score problema > 70"
    },
    
    {
      key: "L3_GENERA_CONFERMA_PROMO",
      livello: "L3_MERGE",
      testo: `Genera una conferma di partecipazione a promo ricevuta.

DATI DISPONIBILI:
- Email originale: {{email_corpo}}
- Fornitore: {{nome_fornitore}}
- Sconto proposto: {{sconto_proposto}}
- Data evento: {{data_evento}}

ISTRUZIONI:
- Ringrazia per adesione
- Conferma sconto e condizioni
- Riepiloga prossimi step
- Chiedi conferma finale

Rispondi SOLO con:
OGGETTO: [oggetto email]
CORPO: [corpo email]`,
      categoria: "GENERAZIONE_EMAIL",
      input: "Analisi merged + dati campagna",
      output: "Email conferma",
      temp: 0.7,
      tokens: 400,
      attivo: "SI",
      note: "Per risposta_campagna > 70"
    },
    
    // LAYER 4 - Review Umana
    {
      key: "L4_SUGGERIMENTI_REVIEW",
      livello: "L4_HUMAN",
      testo: `Genera suggerimenti per review umana di email con bassa confidence.

DATI DISPONIBILI:
- Email: {{email_corpo}}
- Confidence L2: {{confidence}}%
- Divergenza L1-L2: {{divergenza}}%
- Tags L1: {{tags_l1}}
- Tags L2: {{tags_l2}}

ISTRUZIONI:
Spiega in italiano:
1. Perché serve review umana
2. Punti di incertezza
3. Domande da farsi
4. Azioni alternative possibili

Max 150 parole, bullet points.`,
      categoria: "ASSISTENZA_REVIEW",
      input: "Analisi completa con metadati",
      output: "Guida testuale",
      temp: 0.5,
      tokens: 300,
      attivo: "SI",
      note: "Quando confidence < 70"
    }
  ];
  
  prompts.forEach(p => {
    sheet.appendRow([
      p.key, p.livello, p.testo, p.categoria,
      p.input, p.output, p.temp, p.tokens, p.attivo, p.note
    ]);
  });
  
  // Color coding per livelli
  sheet.getRange(2, 2, prompts.length, 1).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  
  Logger.log("✅ Prompts: " + prompts.length + " template configurati");
}

// ═══════════════════════════════════════════════════════════════════════
// MANIFEST
// ═══════════════════════════════════════════════════════════════════════

function creaManifest(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MANIFEST);
  sheet.clear();
  
  const manifest = `
╔══════════════════════════════════════════════════════════════════════╗
║                    MOTORE EMAIL AI - MANIFEST                        ║
╚══════════════════════════════════════════════════════════════════════╝

📦 INFORMAZIONI SISTEMA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Nome Sistema:     ${CONFIG.NOME_SISTEMA}
Versione:         ${CONFIG.VERSION}
Data Release:     ${CONFIG.DATA_RELEASE}
Stato:            PRODUCTION-READY

🎯 SCOPO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Sistema di analisi automatica email B2B con AI stratificato:
- Layer 1 (GPT-4o): Prima categorizzazione rapida
- Layer 2 (Claude): Verifica e raffinamento
- Layer 3 (Merge): Consolidamento e confidence scoring
- Layer 4 (Human): Review assistita per casi incerti
- Layer 5 (AI-Refined): Ri-elaborazione post-correzione

🏗️ ARCHITETTURA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
┌─────────────┐      ┌──────────────┐      ┌─────────────┐
│  LOG_IN     │ ───> │  DISPATCHER  │ ───> │  AILayer.gs │
│  (Email)    │      │  (Jobs)      │      │  (3-Layer)  │
└─────────────┘      └──────────────┘      └─────────────┘
                            │
                            v
                    ┌──────────────┐
                    │   LOG_OUT    │
                    │  (Risposte)  │
                    └──────────────┘

📁 STRUTTURA FILE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Apps Script:
  • Config.gs                - Configurazione centrale
  • InizializzazioneFoglio.gs - Setup e menu
  • Dispatcher.gs            - Orchestratore automazioni
  • AILayer.gs               - Wrapper chiamate AI
  • Logic.gs                 - Logica elaborazione

Fogli Google Sheets:
  • HOME          - Dashboard KPI
  • SETUP         - Collegamenti file esterni
  • AUTOMAZIONI   - Jobs schedulati
  • PROMPTS       - Template AI stratificati
  • LOG_IN        - Email ricevute + analisi
  • LOG_OUT       - Email generate da inviare
  • MANIFEST      - Documentazione (questo file)
  • CHANGELOG     - Storico modifiche
  • LOG_SISTEMA   - Debug

🔧 CONFIGURAZIONE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
API Keys Required:
  ✓ OpenAI API Key (GPT-4o)
  ✓ Claude API Key (Anthropic)

File Esterni (opzionali):
  • MasterSku      - Dati prodotti
  • Fornitori      - Anagrafica fornitori
  • Output BI      - Report business intelligence

Thresholds:
  • Confidence: ${CONFIG.DEFAULTS.CONFIDENCE_THRESHOLD}% (sotto = NEEDS_REVIEW)
  • Divergenza: ${CONFIG.DEFAULTS.DIVERGENCE_THRESHOLD}% (L1 vs L2 troppo diversi)

⚙️ AUTOMAZIONI PREDEFINITE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
JOB-001: Health Check Sistema        [ORARIA - Priorità 1]
JOB-002: Analisi Email Layer 1 (GPT)  [ORARIA - Priorità 20]
JOB-003: Analisi Email Layer 2 (Claude) [ORARIA - Priorità 21]
JOB-004: Merge Analisi Layer 3        [ORARIA - Priorità 22]
JOB-005: Sync File Esterni            [GIORNALIERA 05:00]

📊 METRICHE & KPI
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Tracciamento automatico:
  • Email analizzate/ora
  • Confidence media
  • Tasso di review umana richiesta
  • Divergenza media L1-L2
  • Tempo medio elaborazione
  • Success rate per layer

🚀 QUICK START
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. Menu: 🔑 API Keys > Configura entrambe
2. SETUP: Inserisci ID file esterni (opzionale)
3. LOG_IN: Inserisci email di test
4. Menu: 📧 Analisi Email > Test Singola
5. Verifica risultati in LOG_IN (colonne L1, L2, L3)
6. Menu: 🤖 Automazioni > Configura Trigger Orario
7. Dashboard: Monitora KPI in tempo reale

📝 DOCUMENTAZIONE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Per modifiche strutturali:
  • SEMPRE modificare Config.gs per primo
  • Aggiornare CHANGELOG con data e versione
  • Testare in sandbox prima di deploy

Per nuovi prompt:
  • Specificare LIVELLO_AI corretto
  • Documentare INPUT_ATTESO e OUTPUT_ATTESO
  • Testare con temperatura diversa se output variabile

Per nuove automazioni:
  • Definire dipendenze in DIPENDE_DA
  • Impostare timeout appropriato
  • Gestire errori con retry intelligente

⚠️ LIMITI & BEST PRACTICES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Google Apps Script:
  • 6 minuti max per esecuzione
  • 20 trigger contemporanei max
  • 90 minuti totali/giorno

AI APIs:
  • Rate limiting: controllare usage
  • Cost monitoring: GPT-4o più costoso di mini
  • Fallback: usare GPT-4o-mini per retry

Sheets:
  • Max 10M celle per foglio
  • getValues() una volta, non in loop
  • Batch write quando possibile

🔐 SICUREZZA
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  • API Keys in PropertiesService (mai hardcoded)
  • Log sensibili solo in LOG_SISTEMA
  • File esterni con permessi controllati

📧 SUPPORTO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Per bug o feature request:
  • Aggiorna CHANGELOG con proposta
  • Testa in sandbox
  • Documenta in Manifest

═══════════════════════════════════════════════════════════════════════
Generato automaticamente - ${new Date().toLocaleDateString('it-IT')}
`;
  
  sheet.getRange("A1").setValue(manifest);
  sheet.getRange("A1").setWrap(true);
  sheet.getRange("A1").setFontFamily("Courier New");
  sheet.getRange("A1").setFontSize(9);
  sheet.setColumnWidth(1, 800);
  sheet.setRowHeight(1, 2000);
}

// ═══════════════════════════════════════════════════════════════════════
// CHANGELOG
// ═══════════════════════════════════════════════════════════════════════

function creaChangelog(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CHANGELOG);
  sheet.clear();
  
  const changelog = `
╔══════════════════════════════════════════════════════════════════════╗
║                    MOTORE EMAIL AI - CHANGELOG                       ║
╚══════════════════════════════════════════════════════════════════════╝

[2025-01-30] v1.0.0 - RELEASE INIZIALE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NUOVE FUNZIONALITÀ:
  ✅ Sistema analisi 3-layer (GPT → Claude → Merge)
  ✅ 9 fogli preconfigurati
  ✅ 5 automazioni predefinite
  ✅ 8 prompt stratificati con LIVELLO_AI
  ✅ Collegamenti file esterni (MasterSku, Fornitori)
  ✅ Dispatcher con dipendenze e priorità
  ✅ Manifest & Changelog auto-generati
  ✅ Test automatici integrati
  ✅ Dashboard KPI (placeholder)

STRUTTURA COLONNE:
  • LOG_IN: 29 colonne (6 base + 23 AI layered)
  • LOG_OUT: 10 colonne (include LIVELLO_GENERAZIONE)
  • PROMPTS: 10 colonne (+ LIVELLO_AI, INPUT/OUTPUT)
  • AUTOMAZIONI: 21 colonne (sistema enterprise)

API INTEGRATE:
  • OpenAI GPT-4o (Layer 1)
  • Claude 3.5 Sonnet (Layer 2)
  • GPT-4o-mini (task leggeri/retry)

FILE MODIFICATI:
  • Config.gs (v1.0.0)
  • InizializzazioneFoglio.gs (v1.0.0)
  • AILayer.gs (incluso da Il Segretario)

PROSSIMI SVILUPPI:
  🚧 Dispatcher.gs - Orchestratore completo
  🚧 Logic.gs - Funzioni elaborazione
  🚧 Dashboard.gs - KPI visuali
  🚧 Test suite completa

NOTE TECNICHE:
  • Sintassi italiana per formule (;)
  • Hardcoding ID file in sandbox OK
  • Logger.log() preferito vs foglio in test
  • One-click setup funzionante

═══════════════════════════════════════════════════════════════════════

TEMPLATE PER FUTURE ENTRY:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[DATA] vX.Y.Z - TITOLO BREVE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MODIFICHE:
  • Dettaglio 1
  • Dettaglio 2

FILE MODIFICATI:
  • Nome file (versione)

NOTE:
  • Eventuali note tecniche

═══════════════════════════════════════════════════════════════════════
Ultimo Aggiornamento: ${new Date().toLocaleDateString('it-IT')}
`;
  
  sheet.getRange("A1").setValue(changelog);
  sheet.getRange("A1").setWrap(true);
  sheet.getRange("A1").setFontFamily("Courier New");
  sheet.getRange("A1").setFontSize(9);
  sheet.setColumnWidth(1, 800);
}

// ═══════════════════════════════════════════════════════════════════════
// EMAIL TEST
// ═══════════════════════════════════════════════════════════════════════

function creaEmailTest(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  if (sheet.getLastRow() > 1) return; // Già presente
  
  const emailTest = [
    new Date(),
    "TEST-001",
    "ordini@fornitoretest.it",
    "Re: Promo Shopping Night - Conferma partecipazione",
    `Gentile Team Yumibio,

Confermiamo la nostra partecipazione alla Shopping Night del 15 Febbraio.
Possiamo offrire uno sconto del 20% su tutta la linea corpo.

In allegato il nuovo listino aggiornato con le ultime novità primavera 2025.

Cordiali saluti,
Mario Rossi
Fornitore Test SRL`,
    "",
    // Resto colonne vuote (saranno popolate dall'analisi)
    ...Array(23).fill("")
  ];
  
  sheet.appendRow(emailTest);
  Logger.log("✅ Email test creata");
}

// ═══════════════════════════════════════════════════════════════════════
// API KEYS
// ═══════════════════════════════════════════════════════════════════════

function impostaApiKeyOpenAI() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configurazione OpenAI',
    'Incolla qui la tua OpenAI API Key (sk-...):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
      ui.alert('✅ API Key OpenAI salvata con successo.');
      logSistema("API Key OpenAI configurata");
    }
  }
}

function impostaApiKeyClaude() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Configurazione Claude',
    'Incolla qui la tua Claude API Key (sk-ant-...):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty("CLAUDE_API_KEY", key);
      ui.alert('✅ API Key Claude salvata con successo.');
      logSistema("API Key Claude configurata");
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════

function logSistema(messaggio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_SISTEMA);
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}

function mostraInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '📋 Motore Email AI v1.0.0',
    'Sistema: ' + CONFIG.NOME_SISTEMA + '\n' +
    'Versione: ' + CONFIG.VERSION + '\n' +
    'Release: ' + CONFIG.DATA_RELEASE + '\n\n' +
    '✅ Funzionalità:\n' +
    '• Analisi 3-layer (GPT→Claude→Merge)\n' +
    '• 5 automazioni predefinite\n' +
    '• 8 prompt stratificati\n' +
    '• Collegamenti file esterni\n\n' +
    'Vedi MANIFEST e CHANGELOG per dettagli.',
    ui.ButtonSet.OK
  );
}
/**
 * Menu wrapper per analisi email in coda
 */
function menuAnalizzaEmailInCoda() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '📧 Analizza Email in Coda',
    'Questa operazione eseguirà il flusso completo:\n\n' +
    '1️⃣ L0: Spam Filter\n' +
    '2️⃣ L1: Categorizzazione GPT\n' +
    '3️⃣ L2: Verifica Claude\n' +
    '4️⃣ L3: Merge e Confidence\n\n' +
    'Tempo stimato: 5-10 secondi per email\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    ui.alert('⏳ Elaborazione in corso...', 
             'Attendi il completamento.\nControlla LOG_SISTEMA per il progresso.', 
             ui.ButtonSet.OK);
    
    var risultati = analizzaEmailInCoda(50, false);
    
    var report = '✅ Analisi Completata!\n\n' +
      '📊 RISULTATI:\n' +
      '━━━━━━━━━━━━━━━━━━\n' +
      '🛡️ L0 Spam Filter:\n' +
      '   • SPAM: ' + risultati.l0.spam + '\n' +
      '   • LEGIT: ' + risultati.l0.legit + '\n\n' +
      '🔵 L1 GPT: ' + risultati.l1.ok + ' analizzate\n' +
      '🟣 L2 Claude: ' + risultati.l2.ok + ' verificate\n' +
      '🟢 L3 Merge: ' + risultati.l3.ok + ' completate\n\n' +
      '⚠️ Needs Review: ' + risultati.l3.needsReview + '\n\n' +
      'Controlla il foglio LOG_IN per i dettagli.';
    
    ui.alert('Analisi Completata', report, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('❌ Errore', e.toString() + '\n\nControlla LOG_SISTEMA per dettagli.', ui.ButtonSet.OK);
    logSistema("❌ Errore menuAnalizzaEmailInCoda: " + e.toString());
  }
}