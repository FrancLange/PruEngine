/**
 * ==========================================================================================
 * BATCH_INIT.js - Inizializzazione Infrastruttura Batch v1.0.0
 * ==========================================================================================
 * Setup e integrazione menu per OpenAI Batch API
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// SETUP
// ═══════════════════════════════════════════════════════════════════════

/**
 * Setup completo infrastruttura batch
 */
function setupBatchInfrastructure() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "📦 Setup Batch Infrastructure",
    "Questa operazione creerà:\n\n" +
    "✅ Foglio BATCH_QUEUE (coda richieste)\n" +
    "✅ Foglio BATCH_JOBS (tracking job)\n\n" +
    "Requisiti:\n" +
    "- API Key OpenAI configurata\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    // Verifica API Key
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if (!apiKey) {
      ui.alert("❌ Errore", "API Key OpenAI non configurata.\n\nVai su Menu > API Keys > OpenAI", ui.ButtonSet.OK);
      return;
    }
    
    // Crea fogli
    initBatchSheets();
    
    logBatch("✅ Setup batch infrastructure completato");
    
    ui.alert(
      "✅ Setup Completato",
      "Infrastruttura batch pronta!\n\n" +
      "Fogli creati:\n" +
      "- BATCH_QUEUE\n" +
      "- BATCH_JOBS\n\n" +
      "Prossimi step:\n" +
      "1. Menu > Batch API > Avvia Batch Questioni\n" +
      "2. Attendi 1-24h\n" +
      "3. Menu > Batch API > Processa Risultati",
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
    logBatch("❌ Setup error: " + e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════
// TRIGGER AUTOMATICI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Configura trigger per polling automatico
 */
function configuraTriggerBatch() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "⏰ Configura Trigger Batch",
    "Creerà trigger automatici per:\n\n" +
    "1. Poll status ogni 15 minuti\n" +
    "2. Recupero risultati automatico\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    // Rimuovi trigger esistenti
    rimuoviTriggerBatch();
    
    // Poll ogni 15 minuti
    ScriptApp.newTrigger('triggerPollBatch')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    logBatch("✅ Trigger batch configurato");
    
    ui.alert(
      "✅ Trigger Configurato",
      "Polling automatico ogni 15 minuti.\n\n" +
      "Verifica: Estensioni > Apps Script > Trigger",
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Rimuovi trigger batch
 */
function rimuoviTriggerBatch() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  
  triggers.forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === 'triggerPollBatch' || 
        handler === 'triggerInviaBatch' ||
        handler === 'triggerProcessaRisultati') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  if (count > 0) {
    logBatch("Rimossi " + count + " trigger batch");
  }
}

/**
 * Trigger: Poll e processa automaticamente
 */
function triggerPollBatch() {
  try {
    // Poll status
    var pollResult = pollAllBatchJobs();
    
    // Se ci sono completati, processa
    if (pollResult.completed.length > 0) {
      logBatch("Trigger: " + pollResult.completed.length + " batch completati, processo risultati...");
      
      // Recupera risultati
      recuperaTuttiRisultatiCompletati();
      
      // Processa per questioni
      processaRisultatiBatchQuestioni();
    }
    
  } catch (e) {
    logBatch("❌ Trigger error: " + e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════════════
// TEST
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test connessione OpenAI Batch API
 */
function testBatchApiConnection() {
  var ui = SpreadsheetApp.getUi();
  
  logBatch("🧪 Test connessione Batch API...");
  
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    
    if (!apiKey) {
      ui.alert("❌ Test Fallito", "API Key OpenAI non configurata", ui.ButtonSet.OK);
      return;
    }
    
    // Test: lista file esistenti
    var options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      muteHttpExceptions: true
    };
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.FILES + "?purpose=batch&limit=1";
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      logBatch("✅ Test connessione OK");
      
      // Test: lista batch esistenti
      var urlBatches = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.BATCHES + "?limit=5";
      var responseBatches = UrlFetchApp.fetch(urlBatches, options);
      var batchesData = JSON.parse(responseBatches.getContentText());
      
      var numBatches = batchesData.data ? batchesData.data.length : 0;
      
      ui.alert(
        "✅ Test OK",
        "Connessione OpenAI Batch API funzionante!\n\n" +
        "Batch recenti: " + numBatches,
        ui.ButtonSet.OK
      );
      
    } else {
      var errorBody = JSON.parse(response.getContentText());
      var errorMsg = errorBody.error ? errorBody.error.message : "Errore " + responseCode;
      
      ui.alert("❌ Test Fallito", errorMsg, ui.ButtonSet.OK);
      logBatch("❌ Test fallito: " + errorMsg);
    }
    
  } catch (e) {
    ui.alert("❌ Test Fallito", e.toString(), ui.ButtonSet.OK);
    logBatch("❌ Test exception: " + e.toString());
  }
}

/**
 * Test mini batch (1 richiesta)
 */
function testMiniBatch() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "🧪 Test Mini Batch",
    "Invierà 1 richiesta test a OpenAI Batch API.\n\n" +
    "Costo stimato: ~$0.001\n" +
    "Tempo attesa: 1-60 minuti\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  try {
    // Accoda richiesta test
    var accodaResult = accodaRichiestaBatch({
      tipoOperazione: "TEST",
      customId: "TEST-" + Date.now(),
      systemPrompt: "Rispondi in JSON: {\"status\": \"ok\", \"message\": \"test\"}",
      userPrompt: "Questo è un test. Rispondi con il JSON richiesto.",
      model: BATCH_CONFIG.DEFAULTS.MODEL_MINI,
      maxTokens: 50,
      temperature: 0,
      jsonMode: true,
      priorita: 1
    });
    
    if (!accodaResult.success) {
      ui.alert("❌ Errore", "Accodamento fallito: " + accodaResult.error, ui.ButtonSet.OK);
      return;
    }
    
    // Invia batch
    var invioResult = inviaBatchPending({ forza: true });
    
    if (invioResult.success) {
      ui.alert(
        "✅ Test Inviato",
        "Batch Job: " + invioResult.batchJobId + "\n" +
        "OpenAI Batch: " + invioResult.openAiBatchId + "\n\n" +
        "Controlla tra qualche minuto:\n" +
        "Menu > Batch API > Stato Batch",
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("❌ Errore", "Invio fallito: " + invioResult.error, ui.ButtonSet.OK);
    }
    
  } catch (e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// MENU INTEGRATO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiunge menu batch al menu principale
 * Da chiamare in onOpen()
 */
function aggiungiMenuBatch() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("📦 Batch API")
    .addItem("🛠️ Setup Infrastructure", "setupBatchInfrastructure")
    .addItem("🧪 Test Connessione", "testBatchApiConnection")
    .addSeparator()
    .addSubMenu(ui.createMenu("🎫 Questioni")
      .addItem("🚀 Avvia Batch Questioni", "menuAvviaQuestioniBatch")
      .addItem("📥 Processa Risultati", "menuProcessaRisultatiBatchQuestioni"))
    .addSeparator()
    .addItem("🔄 Poll Status Job", "pollAllBatchJobsMenu")
    .addItem("📊 Stato Batch", "menuStatoBatch")
    .addSeparator()
    .addItem("⏰ Configura Trigger Auto", "configuraTriggerBatch")
    .addItem("🗑️ Rimuovi Trigger", "rimuoviTriggerBatch")
    .addSeparator()
    .addItem("🧪 Test Mini Batch", "testMiniBatch")
    .addToUi();
}

/**
 * Poll con UI feedback
 */
function pollAllBatchJobsMenu() {
  var ui = SpreadsheetApp.getUi();
  
  var result = pollAllBatchJobs();
  
  var msg = "🔄 POLL COMPLETATO\n\n" +
    "Controllati: " + result.checked + "\n" +
    "Completati: " + result.completed.length + "\n" +
    "In corso: " + result.inProgress.length + "\n" +
    "Falliti: " + result.failed.length;
  
  if (result.inProgress.length > 0) {
    msg += "\n\n━━━ IN CORSO ━━━";
    result.inProgress.forEach(function(job) {
      msg += "\n" + job.batchJobId + ": " + job.progress.completed + "/" + job.progress.total;
    });
  }
  
  ui.alert("Poll Status", msg, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// STIMA COSTI UI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Mostra stima costi per batch questioni
 */
function mostraStimaCostiBatch() {
  var ui = SpreadsheetApp.getUi();
  
  // Scan email per stima
  var candidati = scanEmailPerQuestioniBatch({ giorniIndietro: 90, maxEmail: 1000 });
  var gruppi = raggruppaEmailPerFornitoreBatch(candidati);
  var numGruppi = Object.keys(gruppi).length;
  
  // Stima richieste
  var richiesteCluster = numGruppi;
  var richiesteMatch = candidati.length;
  
  // Stima costi (GPT-4o per cluster, GPT-4o-mini per match)
  var costoCluster = stimaCostoBatch(richiesteCluster, "gpt-4o", 800, 500);
  var costoMatch = stimaCostoBatch(richiesteMatch, "gpt-4o-mini", 400, 100);
  
  var costoTotale = costoCluster.totalCost + costoMatch.totalCost;
  
  var msg = "💰 STIMA COSTI BATCH\n\n" +
    "Email candidati: " + candidati.length + "\n" +
    "Gruppi fornitore: " + numGruppi + "\n\n" +
    "━━━ CLUSTERING (GPT-4o) ━━━\n" +
    "Richieste: " + richiesteCluster + "\n" +
    "Costo: $" + costoCluster.totalCost.toFixed(3) + "\n\n" +
    "━━━ MATCHING (GPT-4o-mini) ━━━\n" +
    "Richieste: " + richiesteMatch + "\n" +
    "Costo: $" + costoMatch.totalCost.toFixed(3) + "\n\n" +
    "━━━ TOTALE ━━━\n" +
    "Costo batch (sconto 50%): $" + costoTotale.toFixed(3) + "\n" +
    "Costo realtime equiv.: $" + (costoTotale * 2).toFixed(3);
  
  ui.alert("Stima Costi", msg, ui.ButtonSet.OK);
}