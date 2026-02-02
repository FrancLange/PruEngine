/**
 * ==========================================================================================
 * QUESTIONI_INIT.js - Inizializzazione e Menu v1.0.0 STANDALONE
 * ==========================================================================================
 * Setup ambiente, creazione fogli, menu
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

/**
 * Menu principale - Da chiamare in onOpen()
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('🎫 Questioni Engine')
    .addItem('🛠️ Setup Completo', 'setupQuestioniEngine')
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 Connector Email')
      .addItem('🔗 Configura Connettore', 'configuraConnectorEmail')
      .addItem('🔄 Sync da Email Engine', 'syncEmailMenu')
      .addItem('⏰ Configura Trigger Sync', 'configuraTriggerSyncEmail')
      .addSeparator()
      .addItem('🧪 Test Connessione', 'testConnectorEmail')
      .addItem('📊 Stato Connettore', 'mostraStatoConnectorEmail'));
  
  // ═══ PATCH: CONNECTOR FORNITORI ═══
  if (typeof creaSubmenuConnettoreFornitoriQE === 'function') {
    menu.addSubMenu(creaSubmenuConnettoreFornitoriQE());
  }
  
  menu.addSeparator()
    .addSubMenu(ui.createMenu('🎫 Gestione Questioni')
      .addItem('▶️ Processa Email → Questioni', 'menuProcessaEmailInQuestioni')
      .addItem('📊 Stato Questioni', 'mostraStatoQuestioni')
      .addSeparator()
      .addItem('✏️ Crea Questione Manuale', 'menuCreaQuestioneManuale')
      .addItem('🧪 Test Creazione', 'testCreazioneQuestione'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔑 API Keys')
      .addItem('OpenAI', 'impostaApiKeyOpenAI')
      .addItem('🧪 Test Connessione AI', 'testConnessioneAI'))
    .addSeparator()
    .addItem('ℹ️ Info Sistema', 'mostraInfoSistema')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP COMPLETO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Setup completo Questioni Engine
 */
function setupQuestioniEngine() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    '🛠️ Setup Questioni Engine',
    'Questa operazione creerà:\n\n' +
    '✅ Tab SETUP (configurazione)\n' +
    '✅ Tab EMAIL_IN (sync da Email Engine)\n' +
    '✅ Tab QUESTIONI (ticket)\n' +
    '✅ Tab PROMPTS (AI templates)\n' +
    '✅ Tab LOG_SISTEMA (debug)\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Crea fogli
    creaFogliQuestioni(ss);
    
    // 2. Popola SETUP
    popolaSetupQuestioni(ss);
    
    // 3. Popola PROMPTS
    popolaPromptsQuestioni(ss);
    
    logQuestioni("Setup completo eseguito con successo");
    
    ui.alert(
      '✅ Setup Completato!',
      'Prossimi step:\n\n' +
      '1. Vai in SETUP\n' +
      '2. Trova: CONNECTOR_EMAIL_ENGINE_ID\n' +
      '3. Incolla ID del foglio Email Engine\n' +
      '4. Menu > Connector Email > Sync\n\n' +
      '💡 L\'ID del foglio è nella URL dopo /d/',
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore Setup', e.toString(), ui.ButtonSet.OK);
    logQuestioni("ERRORE Setup: " + e.toString());
  }
}

/**
 * Crea tutti i fogli necessari
 */
function creaFogliQuestioni(ss) {
  var fogli = [
    { nome: QUESTIONI_CONFIG.SHEETS.SETUP, colore: "#666666" },
    { nome: QUESTIONI_CONFIG.SHEETS.EMAIL_IN, colore: "#4CAF50" },
    { nome: QUESTIONI_CONFIG.SHEETS.QUESTIONI, colore: "#2196F3" },
    { nome: QUESTIONI_CONFIG.SHEETS.PROMPTS, colore: "#FFC107" },
    { nome: QUESTIONI_CONFIG.SHEETS.LOG_SISTEMA, colore: "#000000" }
  ];
  
  fogli.forEach(function(foglio) {
    var sheet = ss.getSheetByName(foglio.nome);
    
    if (!sheet) {
      sheet = ss.insertSheet(foglio.nome);
      sheet.setTabColor(foglio.colore);
      logQuestioni("Creato foglio: " + foglio.nome);
    }
    
    // Headers specifici - ORDINE ESPLICITO (importante!)
    if (foglio.nome === QUESTIONI_CONFIG.SHEETS.EMAIL_IN) {
      var headers = [
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.TIMESTAMP_SYNC,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_EMAIL,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.TIMESTAMP_ORIGINALE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.MITTENTE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.OGGETTO,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.CORPO,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_TAGS_FINALI,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_SINTESI_FINALE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_SCORES_JSON,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_CONFIDENCE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.L3_AZIONI_SUGGERITE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_EMAIL,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.ID_QUESTIONE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.STATUS_QUESTIONE,
        QUESTIONI_CONFIG.COLONNE_EMAIL_IN.PROCESSATO
      ];
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#E8F5E9");
      sheet.setFrozenRows(1);
    }
    
    if (foglio.nome === QUESTIONI_CONFIG.SHEETS.QUESTIONI) {
      var headers = [
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.ID_QUESTIONE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.TIMESTAMP_CREAZIONE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.FONTE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.ID_FORNITORE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.NOME_FORNITORE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.EMAIL_FORNITORE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.TITOLO,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.DESCRIZIONE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.CATEGORIA,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.TAGS,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.PRIORITA,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.URGENTE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.STATUS,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.DATA_SCADENZA,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.ASSEGNATO_A,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.EMAIL_COLLEGATE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.NUM_EMAIL,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.TIMESTAMP_RISOLUZIONE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.RISOLUZIONE_NOTE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.AI_CONFIDENCE,
        QUESTIONI_CONFIG.COLONNE_QUESTIONI.AI_CLUSTER_ID
      ];
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#BBDEFB");
      sheet.setFrozenRows(1);
    }
    
    if (foglio.nome === QUESTIONI_CONFIG.SHEETS.PROMPTS) {
      var headers = [
        QUESTIONI_CONFIG.COLONNE_PROMPTS.KEY,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.LIVELLO_AI,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.TESTO,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.CATEGORIA,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.INPUT_ATTESO,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.OUTPUT_ATTESO,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.TEMPERATURA,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.MAX_TOKENS,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.ATTIVO,
        QUESTIONI_CONFIG.COLONNE_PROMPTS.NOTE
      ];
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#FFF9C4");
      sheet.setFrozenRows(1);
    }
    
    if (foglio.nome === QUESTIONI_CONFIG.SHEETS.LOG_SISTEMA) {
      sheet.getRange("A1:B1")
        .setValues([["TIMESTAMP", "MESSAGGIO"]])
        .setFontWeight("bold")
        .setBackground("#EEEEEE");
      sheet.setFrozenRows(1);
    }
  });
}

/**
 * Popola foglio SETUP
 */
function popolaSetupQuestioni(ss) {
  var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.SETUP);
  
  if (sheet.getLastRow() > 1) return; // Già popolato
  
  sheet.getRange("A1:B1")
    .setValues([["CHIAVE", "VALORE"]])
    .setFontWeight("bold")
    .setBackground("#EEEEEE");
  
  var dati = [
    ["═══ CONNECTOR EMAIL ENGINE ═══", ""],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, "Incolla qui ID foglio Email Engine"],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_SYNC_ENABLED, "SI"],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_MAX_RIGHE, QUESTIONI_CONFIG.DEFAULTS.EMAIL_MAX_RIGHE],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_GIORNI_INDIETRO, QUESTIONI_CONFIG.DEFAULTS.EMAIL_GIORNI_INDIETRO],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_FILTRO_SCORE, QUESTIONI_CONFIG.DEFAULTS.EMAIL_FILTRO_SCORE],
    [QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_LAST_SYNC, ""],
    ["", ""],
    ["═══ CONNECTOR CAMPAGNE (futuro) ═══", ""],
    [QUESTIONI_CONFIG.SETUP_KEYS.CAMPAGNE_ENGINE_ID, ""],
    [QUESTIONI_CONFIG.SETUP_KEYS.CAMPAGNE_SYNC_ENABLED, "NO"],
    ["", ""],
    ["═══ CONFIGURAZIONE AI ═══", ""],
    [QUESTIONI_CONFIG.SETUP_KEYS.OPENAI_THRESHOLD_MATCH, QUESTIONI_CONFIG.DEFAULTS.MATCH_THRESHOLD],
    [QUESTIONI_CONFIG.SETUP_KEYS.OPENAI_THRESHOLD_CLUSTER, QUESTIONI_CONFIG.DEFAULTS.CLUSTER_THRESHOLD],
    ["", ""],
    ["═══ SISTEMA ═══", ""],
    [QUESTIONI_CONFIG.SETUP_KEYS.AUTO_CREA_QUESTIONI, "SI"]
  ];
  
  sheet.getRange(2, 1, dati.length, 2).setValues(dati);
  sheet.getRange("A:A").setFontWeight("bold");
  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 400);
}

/**
 * Popola foglio PROMPTS
 */
function popolaPromptsQuestioni(ss) {
  var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.PROMPTS);
  
  if (sheet.getLastRow() > 1) return; // Già popolato
  
  var prompts = [
    {
      key: QUESTIONI_CONFIG.PROMPT_KEYS.CLUSTER_EMAIL,
      livello: "CLUSTERING",
      testo: "Sei un sistema di clustering email per gestione problemi fornitori B2B.\n\n" +
        "Ti vengono fornite N email dallo stesso fornitore/dominio.\n" +
        "Il tuo compito è raggrupparle in CLUSTER tematici (stesso problema/argomento).\n\n" +
        "REGOLE:\n" +
        "- Email sullo stesso ordine/problema → stesso cluster\n" +
        "- Email su argomenti diversi → cluster diversi\n" +
        "- Una email può appartenere a UN solo cluster\n\n" +
        "Rispondi SOLO in JSON:\n" +
        "{\n" +
        '  "clusters": [\n' +
        "    {\n" +
        '      "id": 1,\n' +
        '      "titolo": "Problema consegna ordine #123",\n' +
        '      "email_ids": ["EMAIL-001", "EMAIL-003"],\n' +
        '      "categoria": "PROBLEMA_CONSEGNA",\n' +
        '      "priorita": "ALTA",\n' +
        '      "urgente": true,\n' +
        '      "descrizione": "Breve descrizione del problema"\n' +
        "    }\n" +
        "  ]\n" +
        "}",
      categoria: "QUESTIONI",
      input: "Lista email con id, oggetto, sintesi",
      output: "JSON con clusters",
      temp: 0.3,
      tokens: 800,
      attivo: "SI",
      note: "Clustering email per fornitore"
    },
    {
      key: QUESTIONI_CONFIG.PROMPT_KEYS.MATCH_QUESTIONE,
      livello: "MATCHING",
      testo: "Sei un sistema di matching per questioni/ticket B2B.\n\n" +
        "Ti viene fornita una NUOVA email e una lista di QUESTIONI APERTE.\n" +
        "Determina se l'email si riferisce a una questione esistente.\n\n" +
        "CRITERI MATCH:\n" +
        "- Stesso fornitore E stesso argomento/ordine → MATCH\n" +
        "- Riferimento esplicito a comunicazione precedente → MATCH\n" +
        "- Argomento completamente diverso → NO MATCH\n\n" +
        "Rispondi SOLO in JSON:\n" +
        "{\n" +
        '  "match": true/false,\n' +
        '  "id_questione": "QST-001" o null,\n' +
        '  "confidence": 0-100,\n' +
        '  "motivo": "Breve spiegazione"\n' +
        "}",
      categoria: "QUESTIONI",
      input: "Email + lista questioni aperte",
      output: "JSON con match result",
      temp: 0.2,
      tokens: 300,
      attivo: "SI",
      note: "Match email a questioni esistenti"
    },
    {
      key: QUESTIONI_CONFIG.PROMPT_KEYS.GENERA_TITOLO,
      livello: "GENERAZIONE",
      testo: "Genera un titolo breve e descrittivo per questa questione/ticket.\n\n" +
        "REGOLE:\n" +
        "- Max 60 caratteri\n" +
        "- Includi riferimento ordine se presente\n" +
        "- Chiaro e specifico\n\n" +
        "Rispondi SOLO con il titolo, senza virgolette.",
      categoria: "QUESTIONI",
      input: "Sintesi email/problema",
      output: "Titolo breve",
      temp: 0.5,
      tokens: 50,
      attivo: "SI",
      note: "Genera titolo questione"
    },
    {
      key: QUESTIONI_CONFIG.PROMPT_KEYS.ANALIZZA_URGENZA,
      livello: "ANALISI",
      testo: "Analizza l'urgenza di questa questione.\n\n" +
        "CRITERI URGENZA:\n" +
        "- CRITICA: Blocco produzione, scadenza imminente (<24h)\n" +
        "- ALTA: Problema grave, scadenza vicina (<3 giorni)\n" +
        "- MEDIA: Problema standard, tempo normale\n" +
        "- BASSA: Informativa, nessuna scadenza\n\n" +
        "Rispondi SOLO in JSON:\n" +
        "{\n" +
        '  "priorita": "CRITICA|ALTA|MEDIA|BASSA",\n' +
        '  "urgente": true/false,\n' +
        '  "giorni_scadenza": numero o null,\n' +
        '  "motivo": "Breve spiegazione"\n' +
        "}",
      categoria: "QUESTIONI",
      input: "Descrizione problema",
      output: "JSON con analisi urgenza",
      temp: 0.2,
      tokens: 150,
      attivo: "SI",
      note: "Analizza urgenza questione"
    }
  ];
  
  prompts.forEach(function(p) {
    sheet.appendRow([
      p.key, p.livello, p.testo, p.categoria,
      p.input, p.output, p.temp, p.tokens, p.attivo, p.note
    ]);
  });
  
  logQuestioni("Popolati " + prompts.length + " prompts");
}

// ═══════════════════════════════════════════════════════════════════════
// CONFIGURAZIONE CONNETTORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Wizard configurazione connettore email
 */
function configuraConnectorEmail() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.prompt(
    '📧 Configura Connector Email',
    'Incolla l\'ID del foglio Email Engine:\n\n' +
    '(Lo trovi nella URL dopo /d/ e prima di /edit)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  var engineId = risposta.getResponseText().trim();
  
  if (!engineId || engineId.length < 20) {
    ui.alert("❌ ID non valido", "L'ID deve essere una stringa lunga circa 40 caratteri.", ui.ButtonSet.OK);
    return;
  }
  
  // Verifica accesso
  try {
    var testSs = SpreadsheetApp.openById(engineId);
    var testSheet = testSs.getSheetByName("LOG_IN");
    
    if (!testSheet) {
      ui.alert("❌ Tab LOG_IN non trovata", "Il foglio esiste ma non ha una tab LOG_IN.", ui.ButtonSet.OK);
      return;
    }
    
    // Salva in SETUP
    setQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_ENGINE_ID, engineId);
    setQuestioniSetupValue(QUESTIONI_CONFIG.SETUP_KEYS.EMAIL_SYNC_ENABLED, "SI");
    
    ui.alert(
      "✅ Connettore Configurato!",
      "Email Engine: " + testSs.getName() + "\n\n" +
      "Ora puoi eseguire:\n" +
      "Menu > Connector Email > Sync da Email Engine",
      ui.ButtonSet.OK
    );
    
    logQuestioni("Connector Email configurato: " + engineId);
    
  } catch(e) {
    ui.alert("❌ Errore Accesso", "Impossibile accedere al foglio:\n" + e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// API KEY
// ═══════════════════════════════════════════════════════════════════════

/**
 * Imposta API Key OpenAI
 */
function impostaApiKeyOpenAI() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.prompt(
    'Configurazione OpenAI',
    'Incolla qui la tua OpenAI API Key (sk-...):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var key = result.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
      ui.alert('✅ API Key OpenAI salvata con successo.');
      logQuestioni("API Key OpenAI configurata");
    }
  }
}

/**
 * Test connessione AI
 */
function testConnessioneAI() {
  var ui = SpreadsheetApp.getUi();
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  if (!apiKey) {
    ui.alert("❌ API Key non configurata", "Configura prima la API Key OpenAI.", ui.ButtonSet.OK);
    return;
  }
  
  try {
    var result = chiamataGPTQuestioni("Rispondi solo: OK", "Test", { max_tokens: 10 });
    
    if (result.success) {
      ui.alert("✅ Connessione OK", "OpenAI risponde correttamente.", ui.ButtonSet.OK);
    } else {
      ui.alert("❌ Errore", result.error, ui.ButtonSet.OK);
    }
  } catch(e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// INFO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Mostra info sistema
 */
function mostraInfoSistema() {
  var ui = SpreadsheetApp.getUi();
  
  var stats = getStatsConnectorEmail();
  var questioni = getStatisticheQuestioni();
  
  var info = "🎫 QUESTIONI ENGINE v" + QUESTIONI_CONFIG.VERSION + "\n";
  info += "══════════════════════════════════════\n\n";
  
  info += "📧 CONNECTOR EMAIL\n";
  info += "Stato: " + (stats.attivo ? "✅ Attivo" : "❌ Non attivo") + "\n";
  info += "Email sincronizzate: " + stats.righe + "\n";
  info += "Da processare: " + stats.daProcessare + "\n";
  info += "Ultimo sync: " + (stats.ultimoSync ? stats.ultimoSync.toLocaleString('it-IT') : "Mai") + "\n\n";
  
  info += "🎫 QUESTIONI\n";
  info += "Totali: " + questioni.totali + "\n";
  info += "Aperte: " + questioni.aperte + "\n";
  info += "In lavorazione: " + questioni.inLavorazione + "\n";
  info += "Risolte: " + questioni.risolte + "\n";
  
  ui.alert("Info Sistema", info, ui.ButtonSet.OK);
}

/**
 * Statistiche questioni
 */
function getStatisticheQuestioni() {
  var stats = {
    totali: 0,
    aperte: 0,
    inLavorazione: 0,
    inAttesa: 0,
    risolte: 0,
    chiuse: 0
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) return stats;
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colStatus = headers.indexOf(QUESTIONI_CONFIG.COLONNE_QUESTIONI.STATUS);
    
    stats.totali = data.length - 1;
    
    for (var i = 1; i < data.length; i++) {
      var status = data[i][colStatus];
      
      switch(status) {
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.APERTA:
          stats.aperte++;
          break;
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_LAVORAZIONE:
          stats.inLavorazione++;
          break;
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_ATTESA_RISPOSTA:
          stats.inAttesa++;
          break;
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.RISOLTA:
          stats.risolte++;
          break;
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.CHIUSA:
        case QUESTIONI_CONFIG.STATUS_QUESTIONE.ANNULLATA:
          stats.chiuse++;
          break;
      }
    }
    
  } catch(e) {
    // Ignora
  }
  
  return stats;
}

// ═══════════════════════════════════════════════════════════════════════
// MENU WRAPPER
// ═══════════════════════════════════════════════════════════════════════

/**
 * Menu: Processa email in questioni
 */
function menuProcessaEmailInQuestioni() {
  var ui = SpreadsheetApp.getUi();
  
  // Auto-sync prima
  if (isConnectorEmailAttivo()) {
    autoSyncEmailSeNecessario();
  }
  
  var stats = getStatsConnectorEmail();
  
  if (stats.daProcessare === 0) {
    ui.alert(
      "ℹ️ Nessuna Email da Processare",
      "Non ci sono email da trasformare in questioni.\n\n" +
      "Suggerimenti:\n" +
      "1. Verifica che Email Engine abbia email analizzate\n" +
      "2. Esegui Sync da Email Engine\n" +
      "3. Controlla filtro score in SETUP (default: problema >= 70)",
      ui.ButtonSet.OK
    );
    return;
  }
  
  var risposta = ui.alert(
    '▶️ Processa Email → Questioni',
    'Trovate ' + stats.daProcessare + ' email da processare.\n\n' +
    'Questa operazione:\n' +
    '1. Raggruppa email per fornitore\n' +
    '2. Usa AI per clustering\n' +
    '3. Crea questioni automaticamente\n\n' +
    'Continuare?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    var risultato = processaEmailInQuestioni();
    
    ui.alert(
      '✅ Elaborazione Completata',
      'Email processate: ' + risultato.emailProcessate + '\n' +
      'Questioni create: ' + risultato.questioniCreate + '\n' +
      'Errori: ' + risultato.errori + '\n\n' +
      'Controlla tab QUESTIONI per dettagli.',
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore', e.toString(), ui.ButtonSet.OK);
    logQuestioni("Errore processaEmailInQuestioni: " + e.toString());
  }
}

/**
 * Menu: Crea questione manuale
 */
function menuCreaQuestioneManuale() {
  var ui = SpreadsheetApp.getUi();
  
  var rispostaTitolo = ui.prompt(
    '✏️ Crea Questione Manuale',
    'Inserisci il titolo della questione:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (rispostaTitolo.getSelectedButton() !== ui.Button.OK) return;
  
  var titolo = rispostaTitolo.getResponseText().trim();
  if (!titolo) {
    ui.alert("Titolo obbligatorio");
    return;
  }
  
  var rispostaDesc = ui.prompt(
    'Descrizione',
    'Inserisci una descrizione (opzionale):',
    ui.ButtonSet.OK_CANCEL
  );
  
  var descrizione = "";
  if (rispostaDesc.getSelectedButton() === ui.Button.OK) {
    descrizione = rispostaDesc.getResponseText().trim();
  }
  
  try {
    var questione = creaQuestione({
      titolo: titolo,
      descrizione: descrizione,
      fonte: QUESTIONI_CONFIG.FONTI_QUESTIONE.MANUALE,
      categoria: QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.ALTRO,
      priorita: QUESTIONI_CONFIG.PRIORITA.MEDIA
    });
    
    ui.alert(
      '✅ Questione Creata',
      'ID: ' + questione.id + '\n' +
      'Titolo: ' + questione.titolo,
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu: Mostra stato questioni
 */
function mostraStatoQuestioni() {
  var ui = SpreadsheetApp.getUi();
  var stats = getStatisticheQuestioni();
  
  var report = "📊 STATO QUESTIONI\n";
  report += "══════════════════════════\n\n";
  report += "Totali: " + stats.totali + "\n\n";
  report += "📌 APERTE: " + stats.aperte + "\n";
  report += "🔧 IN LAVORAZIONE: " + stats.inLavorazione + "\n";
  report += "⏳ IN ATTESA: " + stats.inAttesa + "\n";
  report += "✅ RISOLTE: " + stats.risolte + "\n";
  report += "📁 CHIUSE: " + stats.chiuse + "\n";
  
  ui.alert("Stato Questioni", report, ui.ButtonSet.OK);
}

/**
 * Test creazione questione
 */
function testCreazioneQuestione() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var questione = creaQuestione({
      titolo: "TEST - Questione di prova",
      descrizione: "Questa è una questione creata per test",
      fonte: QUESTIONI_CONFIG.FONTI_QUESTIONE.MANUALE,
      categoria: QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.ALTRO,
      priorita: QUESTIONI_CONFIG.PRIORITA.BASSA
    });
    
    ui.alert(
      '✅ Test Riuscito!',
      'Creata questione: ' + questione.id + '\n\n' +
      'Controlla tab QUESTIONI.',
      ui.ButtonSet.OK
    );
    
  } catch(e) {
    ui.alert('❌ Errore Test', e.toString(), ui.ButtonSet.OK);
  }
}