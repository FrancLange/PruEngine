/**
 * ==========================================================================================
 * BATCH_CORE.js - Engine OpenAI Batch API v1.0.0
 * ==========================================================================================
 * Core functions per interagire con OpenAI Batch API:
 * - Upload file JSONL
 * - Creazione batch
 * - Polling status
 * - Download risultati
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// UPLOAD FILE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Upload file JSONL a OpenAI
 * 
 * @param {String} jsonlContent - Contenuto del file JSONL (righe separate da \n)
 * @param {String} filename - Nome file (default: batch_input.jsonl)
 * @returns {Object} {success, fileId, error}
 */
function uploadBatchFile(jsonlContent, filename) {
  var result = { success: false, fileId: null, error: null };
  
  if (!jsonlContent || jsonlContent.trim() === "") {
    result.error = "Contenuto JSONL vuoto";
    return result;
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    result.error = "API Key OpenAI non configurata";
    return result;
  }
  
  filename = filename || "batch_input_" + Date.now() + ".jsonl";
  
  try {
    // Crea blob
    var blob = Utilities.newBlob(jsonlContent, "application/jsonl", filename);
    
    // Prepara multipart request
    var boundary = "----BatchUploadBoundary" + Date.now();
    
    var payload = Utilities.newBlob(
      "--" + boundary + "\r\n" +
      'Content-Disposition: form-data; name="purpose"\r\n\r\n' +
      "batch\r\n" +
      "--" + boundary + "\r\n" +
      'Content-Disposition: form-data; name="file"; filename="' + filename + '"\r\n' +
      "Content-Type: application/jsonl\r\n\r\n"
    ).getBytes();
    
    var fileBytes = blob.getBytes();
    var endBoundary = Utilities.newBlob("\r\n--" + boundary + "--\r\n").getBytes();
    
    // Combina
    var fullPayload = [];
    for (var i = 0; i < payload.length; i++) fullPayload.push(payload[i]);
    for (var j = 0; j < fileBytes.length; j++) fullPayload.push(fileBytes[j]);
    for (var k = 0; k < endBoundary.length; k++) fullPayload.push(endBoundary[k]);
    
    var options = {
      method: "post",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      contentType: "multipart/form-data; boundary=" + boundary,
      payload: fullPayload,
      muteHttpExceptions: true
    };
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.FILES;
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200 && responseBody.id) {
      result.success = true;
      result.fileId = responseBody.id;
      logBatch("✅ File uploaded: " + result.fileId + " (" + filename + ")");
    } else {
      result.error = responseBody.error ? responseBody.error.message : "Upload fallito (" + responseCode + ")";
      logBatch("❌ Upload error: " + result.error);
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Upload exception: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE BATCH
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea batch su OpenAI
 * 
 * @param {String} inputFileId - ID del file JSONL uploadato
 * @param {Object} options - {metadata: {}}
 * @returns {Object} {success, batchId, error}
 */
function createBatch(inputFileId, options) {
  var result = { success: false, batchId: null, error: null };
  
  options = options || {};
  
  if (!inputFileId) {
    result.error = "File ID mancante";
    return result;
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    result.error = "API Key OpenAI non configurata";
    return result;
  }
  
  try {
    var payload = {
      input_file_id: inputFileId,
      endpoint: "/v1/chat/completions",
      completion_window: BATCH_CONFIG.API.COMPLETION_WINDOW
    };
    
    if (options.metadata) {
      payload.metadata = options.metadata;
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
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.BATCHES;
    var response = UrlFetchApp.fetch(url, requestOptions);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200 && responseBody.id) {
      result.success = true;
      result.batchId = responseBody.id;
      result.status = responseBody.status;
      logBatch("✅ Batch created: " + result.batchId);
    } else {
      result.error = responseBody.error ? responseBody.error.message : "Creazione batch fallita (" + responseCode + ")";
      logBatch("❌ Create batch error: " + result.error);
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Create batch exception: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// CHECK STATUS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Controlla status di un batch
 * 
 * @param {String} batchId - ID batch OpenAI
 * @returns {Object} {success, status, outputFileId, errorFileId, counts, error}
 */
function checkBatchStatus(batchId) {
  var result = { 
    success: false, 
    status: null, 
    outputFileId: null, 
    errorFileId: null,
    counts: null,
    error: null 
  };
  
  if (!batchId) {
    result.error = "Batch ID mancante";
    return result;
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    result.error = "API Key OpenAI non configurata";
    return result;
  }
  
  try {
    var options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      muteHttpExceptions: true
    };
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.BATCHES + "/" + batchId;
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200) {
      result.success = true;
      result.status = responseBody.status;
      result.outputFileId = responseBody.output_file_id || null;
      result.errorFileId = responseBody.error_file_id || null;
      result.counts = {
        total: responseBody.request_counts ? responseBody.request_counts.total : 0,
        completed: responseBody.request_counts ? responseBody.request_counts.completed : 0,
        failed: responseBody.request_counts ? responseBody.request_counts.failed : 0
      };
      result.rawResponse = responseBody;
      
      logBatch("📊 Batch " + batchId + " status: " + result.status + 
               " (" + result.counts.completed + "/" + result.counts.total + ")");
    } else {
      result.error = responseBody.error ? responseBody.error.message : "Check status fallito (" + responseCode + ")";
      logBatch("❌ Check status error: " + result.error);
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Check status exception: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// DOWNLOAD RISULTATI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Download contenuto file da OpenAI
 * 
 * @param {String} fileId - ID file da scaricare
 * @returns {Object} {success, content, lines, error}
 */
function downloadBatchFile(fileId) {
  var result = { success: false, content: null, lines: [], error: null };
  
  if (!fileId) {
    result.error = "File ID mancante";
    return result;
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    result.error = "API Key OpenAI non configurata";
    return result;
  }
  
  try {
    var options = {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      muteHttpExceptions: true
    };
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.FILES + "/" + fileId + "/content";
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      result.success = true;
      result.content = response.getContentText();
      
      // Parse linee JSONL
      var lines = result.content.split("\n").filter(function(line) {
        return line && line.trim() !== "";
      });
      
      result.lines = lines.map(function(line) {
        try {
          return JSON.parse(line);
        } catch (e) {
          return { _parseError: true, _raw: line };
        }
      });
      
      logBatch("✅ Downloaded file " + fileId + ": " + result.lines.length + " lines");
    } else {
      result.error = "Download fallito (" + responseCode + ")";
      logBatch("❌ Download error: " + result.error);
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Download exception: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// CANCEL BATCH
// ═══════════════════════════════════════════════════════════════════════

/**
 * Cancella un batch in corso
 * 
 * @param {String} batchId - ID batch da cancellare
 * @returns {Object} {success, error}
 */
function cancelBatch(batchId) {
  var result = { success: false, error: null };
  
  if (!batchId) {
    result.error = "Batch ID mancante";
    return result;
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    result.error = "API Key OpenAI non configurata";
    return result;
  }
  
  try {
    var options = {
      method: "post",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      muteHttpExceptions: true
    };
    
    var url = BATCH_CONFIG.API.BASE_URL + BATCH_CONFIG.API.ENDPOINTS.BATCHES + "/" + batchId + "/cancel";
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = JSON.parse(response.getContentText());
    
    if (responseCode === 200) {
      result.success = true;
      logBatch("🚫 Batch " + batchId + " cancelled");
    } else {
      result.error = responseBody.error ? responseBody.error.message : "Cancel fallito (" + responseCode + ")";
      logBatch("❌ Cancel error: " + result.error);
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Cancel exception: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// BUILD JSONL REQUEST
// ═══════════════════════════════════════════════════════════════════════

/**
 * Costruisce una singola richiesta in formato JSONL per batch
 * 
 * @param {String} customId - ID per ricollegare la risposta
 * @param {String} systemPrompt - System prompt
 * @param {String} userPrompt - User prompt
 * @param {Object} options - {model, maxTokens, temperature, jsonMode}
 * @returns {String} Riga JSON per file JSONL
 */
function buildBatchRequest(customId, systemPrompt, userPrompt, options) {
  options = options || {};
  
  var request = {
    custom_id: customId,
    method: "POST",
    url: "/v1/chat/completions",
    body: {
      model: options.model || BATCH_CONFIG.DEFAULTS.MODEL,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt }
      ],
      max_tokens: options.maxTokens || BATCH_CONFIG.DEFAULTS.MAX_TOKENS,
      temperature: options.temperature !== undefined ? options.temperature : BATCH_CONFIG.DEFAULTS.TEMPERATURE
    }
  };
  
  // JSON mode
  if (options.jsonMode !== false) {
    request.body.response_format = { type: "json_object" };
  }
  
  return JSON.stringify(request);
}

/**
 * Costruisce file JSONL da array di richieste
 * 
 * @param {Array} requests - Array di oggetti {customId, systemPrompt, userPrompt, options}
 * @returns {String} Contenuto file JSONL
 */
function buildBatchJsonl(requests) {
  if (!requests || requests.length === 0) {
    return "";
  }
  
  var lines = requests.map(function(req) {
    return buildBatchRequest(
      req.customId,
      req.systemPrompt,
      req.userPrompt,
      req.options || {}
    );
  });
  
  return lines.join("\n");
}

// ═══════════════════════════════════════════════════════════════════════
// PARSE BATCH RESPONSE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Parse singola risposta da batch output
 * 
 * @param {Object} responseObj - Oggetto risposta dal JSONL di output
 * @returns {Object} {customId, success, content, parsedContent, error, usage}
 */
function parseBatchResponse(responseObj) {
  var result = {
    customId: responseObj.custom_id || null,
    success: false,
    content: null,
    parsedContent: null,
    error: null,
    usage: null
  };
  
  if (!responseObj.response) {
    result.error = "Nessuna risposta";
    return result;
  }
  
  var response = responseObj.response;
  
  // Check errore
  if (response.error) {
    result.error = response.error.message || JSON.stringify(response.error);
    return result;
  }
  
  // Check status code
  if (response.status_code !== 200) {
    result.error = "Status code: " + response.status_code;
    return result;
  }
  
  // Estrai contenuto
  var body = response.body;
  if (body && body.choices && body.choices[0] && body.choices[0].message) {
    result.success = true;
    result.content = body.choices[0].message.content;
    result.usage = body.usage || null;
    
    // Prova a parsare come JSON
    try {
      result.parsedContent = JSON.parse(result.content);
    } catch (e) {
      // Non è JSON, lascia come stringa
      result.parsedContent = result.content;
    }
  } else {
    result.error = "Struttura risposta non valida";
  }
  
  return result;
}

/**
 * Parse tutte le risposte da batch output
 * 
 * @param {Array} lines - Array di oggetti risposta (da downloadBatchFile)
 * @returns {Object} Map customId → parsed response
 */
function parseBatchResponses(lines) {
  var results = {};
  
  if (!lines || lines.length === 0) {
    return results;
  }
  
  lines.forEach(function(line) {
    if (line._parseError) {
      return; // Skip linee non parsabili
    }
    
    var parsed = parseBatchResponse(line);
    if (parsed.customId) {
      results[parsed.customId] = parsed;
    }
  });
  
  return results;
}

// ═══════════════════════════════════════════════════════════════════════
// STIMA COSTI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Stima costo batch
 * 
 * @param {Number} numRequests - Numero richieste
 * @param {String} model - Modello (gpt-4o o gpt-4o-mini)
 * @param {Number} avgInputTokens - Token input medi per richiesta
 * @param {Number} avgOutputTokens - Token output medi per richiesta
 * @returns {Object} {inputCost, outputCost, totalCost}
 */
function stimaCostoBatch(numRequests, model, avgInputTokens, avgOutputTokens) {
  model = model || BATCH_CONFIG.DEFAULTS.MODEL;
  avgInputTokens = avgInputTokens || 500;
  avgOutputTokens = avgOutputTokens || 300;
  
  var costs = BATCH_CONFIG.COSTI[model] || BATCH_CONFIG.COSTI["gpt-4o"];
  
  var totalInputTokens = numRequests * avgInputTokens;
  var totalOutputTokens = numRequests * avgOutputTokens;
  
  var inputCost = (totalInputTokens / 1000000) * costs.input;
  var outputCost = (totalOutputTokens / 1000000) * costs.output;
  
  return {
    inputCost: Math.round(inputCost * 1000) / 1000,
    outputCost: Math.round(outputCost * 1000) / 1000,
    totalCost: Math.round((inputCost + outputCost) * 1000) / 1000,
    model: model,
    numRequests: numRequests
  };
}