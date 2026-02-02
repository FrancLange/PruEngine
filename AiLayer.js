/**
 * ==========================================================================================
 * AI LAYER - Wrapper Chiamate GPT e Claude v1.1.0
 * ==========================================================================================
 * Gestisce tutte le interazioni con i modelli AI
 * 
 * CHANGELOG v1.1.0:
 * - RIMOSSO: analisiLayer0SpamFilter() - ora in Filters.js
 * - RIMOSSO: testLayer0SpamFilter() - ora in Tests.js
 * - RIMOSSO: testLayer0SuFoglio() - ora in Tests.js
 * - Pulizia codice duplicato
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CHIAMATE API BASE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Chiamata a OpenAI GPT
 * @param {String} systemPrompt - System prompt
 * @param {String} userPrompt - User prompt
 * @param {Object} options - {model, temperature, max_tokens, json_mode}
 * @returns {Object} {success, content, model, usage, error}
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

/**
 * Chiamata a Anthropic Claude
 * @param {String} systemPrompt - System prompt
 * @param {String} userPrompt - User prompt
 * @param {Object} options - {model, max_tokens}
 * @returns {Object} {success, content, model, usage, error}
 */
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


// ═══════════════════════════════════════════════════════════════════════
// LAYER 1 - GPT CATEGORIZATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Layer 1: Prima analisi con GPT-4o
 * @param {Object} emailData - {mittente, oggetto, corpo}
 * @returns {Object} {success, tags, sintesi, scores, modello, timestamp}
 */
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


// ═══════════════════════════════════════════════════════════════════════
// LAYER 2 - CLAUDE VERIFICATION (con fallback GPT)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Layer 2: Verifica con Claude (o GPT come fallback)
 * @param {Object} emailData - {mittente, oggetto, corpo}
 * @param {Object} layer1Result - Risultato Layer 1
 * @returns {Object} {success, tags, sintesi, scores, confidence, richiestaRetry, note, modello, timestamp}
 */
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

/**
 * Layer 2 con GPT (fallback se Claude non disponibile)
 */
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


// ═══════════════════════════════════════════════════════════════════════
// LAYER 3 - MERGE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Layer 3: Merge risultati L1 + L2
 * @param {Object} layer1 - Risultato Layer 1
 * @param {Object} layer2 - Risultato Layer 2
 * @returns {Object} {success, tags, sintesi, scores, confidence, divergenza, needsReview, azioniSuggerite, timestamp}
 */
function mergeAnalisi(layer1, layer2) {
  var threshold = CONFIG.AI_CONFIG.CONFIDENCE_THRESHOLD || 70;
  var divergenceThreshold = CONFIG.AI_CONFIG.DIVERGENCE_THRESHOLD || 30;
  
  var divergenza = calcolaDivergenza(layer1.scores, layer2.scores);
  
  // Merge tags (unione senza duplicati)
  var tagsMerged = [];
  var allTags = (layer1.tags || []).concat(layer2.tags || []);
  for (var i = 0; i < allTags.length; i++) {
    if (tagsMerged.indexOf(allTags[i]) === -1) {
      tagsMerged.push(allTags[i]);
    }
  }
  
  // Merge scores (media pesata su confidence L2)
  var pesoL2 = (layer2.confidence || 50) / 100;
  var pesoL1 = 1 - pesoL2;
  var scoresMerged = {};
  
  var scoreKeys = ["novita", "promo", "fattura", "catalogo", "prezzi", "problema", "risposta_campagna", "urgente", "ordine"];
  for (var j = 0; j < scoreKeys.length; j++) {
    var key = scoreKeys[j];
    var s1 = (layer1.scores && layer1.scores[key]) || 0;
    var s2 = (layer2.scores && layer2.scores[key]) || 0;
    scoresMerged[key] = Math.round(s1 * pesoL1 + s2 * pesoL2);
  }
  
  // Calcola confidence finale
  var confidenceFinale = layer2.confidence || 50;
  if (divergenza > divergenceThreshold) {
    confidenceFinale = Math.max(0, confidenceFinale - 20);
  }
  
  // Determina se serve review
  var needsReview = confidenceFinale < threshold || 
                    divergenza > divergenceThreshold || 
                    layer2.richiestaRetry;
  
  // Scegli sintesi
  var sintesiFinale = (layer2.confidence || 0) > 70 ? layer2.sintesi : layer1.sintesi;
  
  // Genera azioni suggerite
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

/**
 * Calcola divergenza media tra due set di scores
 */
function calcolaDivergenza(scores1, scores2) {
  var keys = ["novita", "promo", "fattura", "catalogo", "prezzi", "problema", "risposta_campagna", "urgente", "ordine"];
  var totalDiff = 0;
  var count = 0;
  
  scores1 = scores1 || {};
  scores2 = scores2 || {};
  
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var s1 = scores1[key] || 0;
    var s2 = scores2[key] || 0;
    totalDiff += Math.abs(s1 - s2);
    count++;
  }
  
  return count > 0 ? Math.round(totalDiff / count) : 0;
}

/**
 * Genera azioni suggerite basate sugli scores
 */
function generaAzioniSuggerite(scores) {
  var azioni = [];
  
  if (scores.problema > 70) azioni.push("CREA_QUESTIONE_PROBLEMA");
  if (scores.fattura > 70) azioni.push("FLAG_FATTURA_FORNITORE");
  if (scores.novita > 70) azioni.push("FLAG_NOVITA_FORNITORE");
  if (scores.catalogo > 70) azioni.push("FLAG_CATALOGO_FORNITORE");
  if (scores.prezzi > 70) azioni.push("FLAG_PREZZI_FORNITORE");
  if (scores.risposta_campagna > 70) azioni.push("AGGIORNA_ESITO_CAMPAGNA");
  if (scores.urgente > 80) azioni.push("ALERT_URGENTE");
  if (scores.ordine > 70) azioni.push("FLAG_ORDINE");
  
  return azioni;
}


// ═══════════════════════════════════════════════════════════════════════
// TEST CONNESSIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test connessioni API (GPT e Claude)
 */
function testConnessioniAI() {
  var ui = SpreadsheetApp.getUi();
  var report = "TEST CONNESSIONI API\n\n";
  
  // Test GPT
  var gptResult = chiamataGPT("Sei un assistente.", "Rispondi solo con: OK", { max_tokens: 10 });
  report += "OpenAI GPT: " + (gptResult.success ? "✅ OK" : "❌ ERRORE - " + gptResult.error) + "\n";
  
  // Test Claude (opzionale)
  var claudeKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (claudeKey) {
    var claudeResult = chiamataClaude("Sei un assistente.", "Rispondi solo con: OK", { max_tokens: 10 });
    report += "Claude: " + (claudeResult.success ? "✅ OK" : "❌ ERRORE - " + claudeResult.error) + "\n";
  } else {
    report += "Claude: ⚠️ Non configurato (Fallback: GPT-4o)\n";
  }
  
  ui.alert("Test API", report, ui.ButtonSet.OK);
  logSistema(report.replace(/\n/g, " | "));
}