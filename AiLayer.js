/**
 * AI LAYER - Wrapper Chiamate GPT e Claude v1.0.3
 * Gestisce tutte le interazioni con i modelli AI
 * 
 * CHANGELOG v1.0.3:
 * - Aggiunto fallback GPT per Layer 2 se Claude non configurato
 * - Test connessioni più flessibile (Claude opzionale)
 */
function chiamataGPT(systemPrompt, userPrompt, options) {
  options = options || {};
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  if (!apiKey) {
    return { success: false, content: null, error: "API Key OpenAI non configurata" };
  }
  
  var model = options.model || "gpt-4o";
  var temperature = options.temperature !== undefined ? options.temperature : 0.7;
  var maxTokens = options.max_tokens || 1000;
  
  var payload = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt }
    ],
    temperature: temperature,
    max_tokens: maxTokens
  };
  
  if (options.json_mode) {
    payload.response_format = { type: "json_object" };
  }
  
  var requestOptions = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200 && responseBody.choices && responseBody.choices[0]) {
      return {
        success: true,
        content: responseBody.choices[0].message.content,
        model: model,
        usage: responseBody.usage,
        error: null
      };
    } else {
      var errorMsg = responseBody.error ? responseBody.error.message : "Errore sconosciuto";
      logSistema("Errore GPT (" + responseCode + "): " + errorMsg);
      return { success: false, content: null, error: errorMsg };
    }
  } catch (e) {
    logSistema("Eccezione chiamata GPT: " + e.toString());
    return { success: false, content: null, error: e.toString() };
  }
}

function chiamataClaude(systemPrompt, userPrompt, options) {
  options = options || {};
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  
  if (!apiKey) {
    return { success: false, content: null, error: "API Key Claude non configurata" };
  }
  
  var model = options.model || "claude-3-5-sonnet-20241022";
  var maxTokens = options.max_tokens || 1000;
  
  var payload = {
    model: model,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [
      { role: "user", content: userPrompt }
    ]
  };
  
  var requestOptions = {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "Content-Type": "application/json",
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", requestOptions);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200 && responseBody.content && responseBody.content[0]) {
      return {
        success: true,
        content: responseBody.content[0].text,
        model: model,
        usage: responseBody.usage,
        error: null
      };
    } else {
      var errorMsg = responseBody.error ? responseBody.error.message : "Errore sconosciuto";
      logSistema("Errore Claude (" + responseCode + "): " + errorMsg);
      return { success: false, content: null, error: errorMsg };
    }
  } catch (e) {
    logSistema("Eccezione chiamata Claude: " + e.toString());
    return { success: false, content: null, error: e.toString() };
  }
}

function analisiLayer1(emailData) {
  var systemPrompt = "Sei un analista esperto di email commerciali B2B. " +
    "Analizza l'email e rispondi SOLO in formato JSON valido con questa struttura esatta: " +
    "{\"tags\": [\"tag1\", \"tag2\"], \"sintesi\": \"Breve sintesi in italiano (max 100 parole)\", " +
    "\"scores\": {\"novita\": 0, \"promo\": 0, \"fattura\": 0, \"catalogo\": 0, \"prezzi\": 0, " +
    "\"problema\": 0, \"risposta_campagna\": 0, \"urgente\": 0, \"ordine\": 0}} " +
    "I punteggi vanno da 0 a 100. Assegna punteggi alti solo se l'evidenza e chiara. " +
    "Tags possibili: NOVITA, PROMO, FATTURA, CATALOGO, PREZZI, PROBLEMA, RISPOSTA_CAMPAGNA, URGENTE, ORDINE, INFO, CONFERMA, LISTINO";

  var userPrompt = "Analizza questa email: " +
    "MITTENTE: " + emailData.mittente + " " +
    "OGGETTO: " + emailData.oggetto + " " +
    "CORPO: " + emailData.corpo + " " +
    "Rispondi SOLO con il JSON, senza altri commenti.";

  var result = chiamataGPT(systemPrompt, userPrompt, { 
    temperature: 0.3, 
    json_mode: true,
    max_tokens: 500 
  });
  
  if (result.success) {
    try {
      var parsed = JSON.parse(result.content);
      return {
        success: true,
        tags: parsed.tags || [],
        sintesi: parsed.sintesi || "",
        scores: parsed.scores || {},
        modello: result.model,
        timestamp: new Date()
      };
    } catch (e) {
      return { success: false, error: "Parsing JSON fallito: " + e.toString() };
    }
  }
  return { success: false, error: result.error };
}

function analisiLayer2(emailData, layer1Result) {
  // FALLBACK: Verifica se Claude e disponibile
  var claudeKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  
  if (!claudeKey) {
    logSistema("⚠️ Claude non configurato - Fallback su GPT-4o per Layer 2");
    return analisiLayer2ConGPT(emailData, layer1Result);
  }
  
  // Normale flusso Claude
  var systemPrompt = "Sei un revisore esperto di analisi email commerciali B2B. " +
    "Ti viene fornita un'email e l'analisi preliminare fatta da un altro sistema. " +
    "Il tuo compito e: 1. Verificare l'accuratezza dell'analisi 2. Correggere eventuali errori " +
    "3. Assegnare un punteggio di confidence (0-100) alla tua analisi " +
    "4. Indicare se hai bisogno di una terza opinione (richiestaRetry) " +
    "Rispondi SOLO in formato JSON valido: " +
    "{\"tags\": [\"tag1\", \"tag2\"], \"sintesi\": \"Sintesi raffinata (max 100 parole)\", " +
    "\"scores\": {\"novita\": 0, \"promo\": 0, \"fattura\": 0, \"catalogo\": 0, \"prezzi\": 0, " +
    "\"problema\": 0, \"risposta_campagna\": 0, \"urgente\": 0, \"ordine\": 0}, " +
    "\"confidence\": 85, \"richiestaRetry\": false, \"note\": \"Eventuali note sulla revisione\"}";

  var userPrompt = "EMAIL DA ANALIZZARE: " +
    "MITTENTE: " + emailData.mittente + " " +
    "OGGETTO: " + emailData.oggetto + " " +
    "CORPO: " + emailData.corpo + " " +
    "ANALISI PRELIMINARE (Layer 1): " +
    "Tags: " + layer1Result.tags.join(", ") + " " +
    "Sintesi: " + layer1Result.sintesi + " " +
    "Scores: " + JSON.stringify(layer1Result.scores) + " " +
    "Verifica e correggi se necessario. Rispondi SOLO con il JSON.";

  var result = chiamataClaude(systemPrompt, userPrompt, { max_tokens: 600 });
  
  if (result.success) {
    try {
      var content = result.content.trim();
      content = content.replace(/```json\n?/g, "").replace(/```$/g, "");
      
      var parsed = JSON.parse(content);
      return {
        success: true,
        tags: parsed.tags || [],
        sintesi: parsed.sintesi || "",
        scores: parsed.scores || {},
        confidence: parsed.confidence || 50,
        richiestaRetry: parsed.richiestaRetry || false,
        note: parsed.note || "",
        modello: result.model,
        timestamp: new Date()
      };
    } catch (e) {
      return { success: false, error: "Parsing JSON fallito: " + e.toString() };
    }
  }
  return { success: false, error: result.error };
}

function analisiLayer2ConGPT(emailData, layer1Result) {
  var systemPrompt = "Sei un revisore esperto di analisi email commerciali B2B. " +
    "Ti viene fornita un'email e l'analisi preliminare fatta da un altro sistema. " +
    "Il tuo compito e: 1. Verificare l'accuratezza dell'analisi 2. Correggere eventuali errori " +
    "3. Assegnare un punteggio di confidence (0-100) alla tua analisi " +
    "4. Indicare se hai bisogno di una terza opinione (richiestaRetry) " +
    "Rispondi SOLO in formato JSON valido: " +
    "{\"tags\": [\"tag1\", \"tag2\"], \"sintesi\": \"Sintesi raffinata (max 100 parole)\", " +
    "\"scores\": {\"novita\": 0, \"promo\": 0, \"fattura\": 0, \"catalogo\": 0, \"prezzi\": 0, " +
    "\"problema\": 0, \"risposta_campagna\": 0, \"urgente\": 0, \"ordine\": 0}, " +
    "\"confidence\": 85, \"richiestaRetry\": false, \"note\": \"Eventuali note sulla revisione\"}";

  var userPrompt = "EMAIL DA ANALIZZARE: " +
    "MITTENTE: " + emailData.mittente + " " +
    "OGGETTO: " + emailData.oggetto + " " +
    "CORPO: " + emailData.corpo + " " +
    "ANALISI PRELIMINARE (Layer 1): " +
    "Tags: " + layer1Result.tags.join(", ") + " " +
    "Sintesi: " + layer1Result.sintesi + " " +
    "Scores: " + JSON.stringify(layer1Result.scores) + " " +
    "Verifica e correggi se necessario. Rispondi SOLO con il JSON.";

  var result = chiamataGPT(systemPrompt, userPrompt, {
    temperature: 0.2,
    json_mode: true,
    max_tokens: 600
  });
  
  if (result.success) {
    try {
      var parsed = JSON.parse(result.content);
      return {
        success: true,
        tags: parsed.tags || [],
        sintesi: parsed.sintesi || "",
        scores: parsed.scores || {},
        confidence: parsed.confidence || 50,
        richiestaRetry: parsed.richiestaRetry || false,
        note: parsed.note || "",
        modello: "gpt-4o (fallback L2)",
        timestamp: new Date()
      };
    } catch (e) {
      return { success: false, error: "Parsing JSON fallito: " + e.toString() };
    }
  }
  return { success: false, error: result.error };
}

function mergeAnalisi(layer1, layer2) {
  var threshold = 70;
  var divergenceThreshold = 30;
  
  var divergenza = calcolaDivergenza(layer1.scores, layer2.scores);
  
  var tagsMerged = [];
  var allTags = layer1.tags.concat(layer2.tags);
  for (var i = 0; i < allTags.length; i++) {
    if (tagsMerged.indexOf(allTags[i]) === -1) {
      tagsMerged.push(allTags[i]);
    }
  }
  
  var pesoL2 = layer2.confidence / 100;
  var pesoL1 = 1 - pesoL2;
  var scoresMerged = {};
  
  var scoreKeys = ["novita", "promo", "fattura", "catalogo", "prezzi", "problema", "risposta_campagna", "urgente", "ordine"];
  for (var j = 0; j < scoreKeys.length; j++) {
    var key = scoreKeys[j];
    var s1 = layer1.scores[key] || 0;
    var s2 = layer2.scores[key] || 0;
    scoresMerged[key] = Math.round(s1 * pesoL1 + s2 * pesoL2);
  }
  
  var confidenceFinale = layer2.confidence;
  if (divergenza > divergenceThreshold) {
    confidenceFinale = Math.max(0, confidenceFinale - 20);
  }
  
  var needsReview = confidenceFinale < threshold || divergenza > divergenceThreshold || layer2.richiestaRetry;
  
  var sintesiFinale = layer2.confidence > 70 ? layer2.sintesi : layer1.sintesi;
  
  var azioniSuggerite = generaAzioniSuggerite(scoresMerged);
  
  return {
    success: true,
    tags: tagsMerged,
    sintesi: sintesiFinale,
    scores: scoresMerged,
    confidence: confidenceFinale,
    divergenza: divergenza,
    needsReview: needsReview,
    azioniSuggerite: azioniSuggerite,
    timestamp: new Date()
  };
}

function calcolaDivergenza(scores1, scores2) {
  var keys = ["novita", "promo", "fattura", "catalogo", "prezzi", "problema", "risposta_campagna", "urgente", "ordine"];
  var totalDiff = 0;
  var count = 0;
  
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var s1 = scores1[key] || 0;
    var s2 = scores2[key] || 0;
    totalDiff += Math.abs(s1 - s2);
    count++;
  }
  
  return count > 0 ? Math.round(totalDiff / count) : 0;
}

function generaAzioniSuggerite(scores) {
  var azioni = [];
  
  if (scores.problema > 70) azioni.push("CREA_QUESTIONE_PROBLEMA");
  if (scores.fattura > 70) azioni.push("FLAG_FATTURA_FORNITORE");
  if (scores.novita > 70) azioni.push("FLAG_NOVITA_FORNITORE");
  if (scores.catalogo > 70) azioni.push("FLAG_CATALOGO_FORNITORE");
  if (scores.prezzi > 70) azioni.push("FLAG_PREZZI_FORNITORE");
  if (scores.risposta_campagna > 70) azioni.push("AGGIORNA_ESITO_CAMPAGNA");
  if (scores.urgente > 80) azioni.push("ALERT_URGENTE");
  
  return azioni;
}

function testConnessioniAI() {
  var ui = SpreadsheetApp.getUi();
  var report = "TEST CONNESSIONI API\n\n";
  
  var gptResult = chiamataGPT("Sei un assistente.", "Rispondi solo con: OK", { max_tokens: 10 });
  report += "OpenAI GPT: " + (gptResult.success ? "OK" : "ERRORE - " + gptResult.error) + "\n";
  
  var claudeKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (claudeKey) {
    var claudeResult = chiamataClaude("Sei un assistente.", "Rispondi solo con: OK", { max_tokens: 10 });
    report += "Claude: " + (claudeResult.success ? "OK" : "ERRORE - " + claudeResult.error) + "\n";
  } else {
    report += "Claude: Non configurato (Fallback: GPT-4o)\n";
  }
  
  ui.alert("Test API", report, ui.ButtonSet.OK);
  logSistema(report.replace(/\n/g, " | "));
}

function logSistema(messaggio) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LOG_SISTEMA");
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}
/**
 * ═══════════════════════════════════════════════════════════════════════
 * LAYER 0 - SPAM FILTER (GPT-4o-mini)
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Layer 0: Spam Filter Ultra-Rapido
 * @param {Object} emailData - {mittente, oggetto, corpo}
 * @returns {Object} {success, isSpam, confidence, reason, modello}
 */
function analisiLayer0SpamFilter(emailData) {
  var systemPrompt = "Sei un filtro anti-spam specializzato in email B2B commerciali.\n\n" +
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
    "}";

  var userPrompt = "Classifica questa email:\n\n" +
    "MITTENTE: " + emailData.mittente + "\n" +
    "OGGETTO: " + emailData.oggetto + "\n" +
    "CORPO (prime 500 caratteri):\n" +
    emailData.corpo.substring(0, 500) + "...\n\n" +
    "Rispondi SOLO con il JSON, senza altri commenti.";

  var options = {
    model: CONFIG.AI_CONFIG.MODEL_SPAM_FILTER,
    temperature: CONFIG.AI_CONFIG.L0_TEMPERATURE,
    max_tokens: CONFIG.AI_CONFIG.L0_MAX_TOKENS,
    json_mode: true
  };

  var result = chiamataGPT(systemPrompt, userPrompt, options);
  
  if (result.success) {
    try {
      var parsed = JSON.parse(result.content);
      
      // Validazione
      if (typeof parsed.isSpam !== "boolean") {
        throw new Error("isSpam non e' boolean");
      }
      
      var confidence = parseInt(parsed.confidence) || 0;
      
      // Safety: Se confidence bassa, considera LEGIT
      var isSafeSpam = parsed.isSpam && 
                       confidence >= CONFIG.AI_CONFIG.L0_CONFIDENCE_THRESHOLD;
      
      return {
        success: true,
        isSpam: isSafeSpam,
        confidence: confidence,
        reason: isSafeSpam ? (parsed.reason || "Spam generico") : null,
        modello: result.model,
        timestamp: new Date()
      };
    } catch (e) {
      logSistema("Errore parsing L0: " + e.toString());
      // In caso di errore, assume LEGIT (safe default)
      return { 
        success: false, 
        isSpam: false,  // DEFAULT SAFE
        confidence: 0,
        reason: null,
        error: "Parsing fallito: " + e.toString() 
      };
    }
  }
  
  // Se chiamata API fallisce, assume LEGIT
  return { 
    success: false, 
    isSpam: false,  // DEFAULT SAFE
    confidence: 0,
    reason: null,
    error: result.error 
  };
}

/**
 * Test Layer 0 Spam Filter
 */
function testLayer0SpamFilter() {
  var ui = SpreadsheetApp.getUi();
  
  // Test 1: Email legit
  logSistema("🧪 Test L0: Email legit (conferma ordine)");
  var emailLegit = {
    mittente: "ordini@fornitoretest.it",
    oggetto: "Conferma Ordine #12345",
    corpo: "Gentile cliente, confermiamo ricezione ordine. Totale: 1250 euro. Grazie."
  };
  
  var resultLegit = analisiLayer0SpamFilter(emailLegit);
  logSistema("Risultato LEGIT: " + JSON.stringify(resultLegit));
  
  // Test 2: Email spam
  logSistema("🧪 Test L0: Email spam (newsletter)");
  var emailSpam = {
    mittente: "noreply@marketing-platform.com",
    oggetto: "🎉 Super Offerta! Sconto 90% Solo Oggi!!!",
    corpo: "Clicca qui per vincere un iPhone! Offerta limitata! Non perdere questa occasione unica!"
  };
  
  var resultSpam = analisiLayer0SpamFilter(emailSpam);
  logSistema("Risultato SPAM: " + JSON.stringify(resultSpam));
  
  // Report
  var report = "🧪 TEST LAYER 0 SPAM FILTER\n\n" +
    "━━━ EMAIL LEGIT ━━━\n" +
    "IsSpam: " + resultLegit.isSpam + "\n" +
    "Confidence: " + resultLegit.confidence + "%\n" +
    "Modello: " + resultLegit.modello + "\n\n" +
    "━━━ EMAIL SPAM ━━━\n" +
    "IsSpam: " + resultSpam.isSpam + "\n" +
    "Confidence: " + resultSpam.confidence + "%\n" +
    "Reason: " + (resultSpam.reason || "-") + "\n" +
    "Modello: " + resultSpam.modello + "\n\n" +
    (resultLegit.isSpam === false && resultSpam.isSpam === true ? 
     "✅ TEST SUPERATO!" : 
     "⚠️ Verifica risultati");
  
  ui.alert("Test L0 Completato", report, ui.ButtonSet.OK);
}
/**
 * Test Layer 0 su email reali nel foglio
 */
function testLayer0SuFoglio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Foglio LOG_IN non trovato");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  data.shift();
  
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colMittente = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colOggetto = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  var colCorpo = headers.indexOf(CONFIG.COLONNE_LOG_IN.CORPO);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  var emailDaFiltrare = [];
  
  // Carica email DA_FILTRARE
  for (var i = 0; i < data.length; i++) {
    if (data[i][colStatus] === CONFIG.STATUS_EMAIL.DA_FILTRARE) {
      emailDaFiltrare.push({
        rowNum: i + 2,
        id: data[i][colId],
        mittente: data[i][colMittente],
        oggetto: data[i][colOggetto],
        corpo: data[i][colCorpo]
      });
    }
  }
  
  if (emailDaFiltrare.length === 0) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Nessuna Email da Filtrare",
      "Non ci sono email con STATUS = DA_FILTRARE.\n\n" +
      "Esegui prima: creaEmailTestSpam()",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  logSistema("🛡️ TEST L0: Inizio filtro su " + emailDaFiltrare.length + " email");
  
  var risultati = [];
  
  emailDaFiltrare.forEach(function(email) {
    logSistema("Filtro " + email.id + "...");
    
    var l0Result = analisiLayer0SpamFilter(email);
    
    risultati.push({
      id: email.id,
      isSpam: l0Result.isSpam,
      confidence: l0Result.confidence,
      reason: l0Result.reason || "-"
    });
    
    // Scrivi risultati nel foglio
    scriviRisultatoL0(sheet, email.rowNum, l0Result);
    
    // Aggiorna status
    var nuovoStatus = l0Result.isSpam ? 
      CONFIG.STATUS_EMAIL.SPAM : 
      CONFIG.STATUS_EMAIL.DA_ANALIZZARE;
    
    aggiornaStatus(sheet, email.rowNum, nuovoStatus);
    
    logSistema(email.id + " → " + (l0Result.isSpam ? "SPAM" : "LEGIT") + 
               " (conf: " + l0Result.confidence + "%)");
  });
  
  // Report finale
  var spamCount = risultati.filter(function(r) { return r.isSpam; }).length;
  var legitCount = risultati.length - spamCount;
  
  var report = "🛡️ TEST LAYER 0 COMPLETATO\n\n" +
    "Email analizzate: " + risultati.length + "\n" +
    "🚫 SPAM: " + spamCount + "\n" +
    "✅ LEGIT: " + legitCount + "\n\n" +
    "━━━ DETTAGLIO ━━━\n";
  
  risultati.forEach(function(r) {
    var icon = r.isSpam ? "🚫" : "✅";
    report += icon + " " + r.id + ": " + 
              (r.isSpam ? "SPAM" : "LEGIT") + 
              " (" + r.confidence + "%) - " + r.reason + "\n";
  });
  
  report += "\n━━━ ATTESO ━━━\n" +
    "SPAM-001, SPAM-002, SPAM-003 → SPAM\n" +
    "LEGIT-001, LEGIT-002, LEGIT-003 → LEGIT\n\n" +
    "Verifica il foglio LOG_IN per dettagli.";
  
  SpreadsheetApp.getUi().alert("Test L0 Completato", report, SpreadsheetApp.getUi().ButtonSet.OK);
  logSistema("🛡️ TEST L0 COMPLETATO: " + spamCount + " spam, " + legitCount + " legit");
}