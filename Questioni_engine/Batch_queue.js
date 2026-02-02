/**
 * ==========================================================================================
 * BATCH_QUEUE.js - Gestione Coda e Job v1.0.0
 * ==========================================================================================
 * Gestisce:
 * - Accodamento richieste in BATCH_QUEUE
 * - Tracking job in BATCH_JOBS
 * - Invio batch a OpenAI
 * - Polling e recupero risultati
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// INIZIALIZZAZIONE FOGLI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea fogli BATCH_QUEUE e BATCH_JOBS se non esistono
 */
function initBatchSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // BATCH_QUEUE
  var queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
  if (!queueSheet) {
    queueSheet = ss.insertSheet(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    queueSheet.setTabColor("#FFB74D"); // Arancione
    
    var queueHeaders = Object.values(BATCH_CONFIG.COLONNE_QUEUE);
    queueSheet.getRange(1, 1, 1, queueHeaders.length)
      .setValues([queueHeaders])
      .setFontWeight("bold")
      .setBackground("#FFF3E0");
    queueSheet.setFrozenRows(1);
    
    logBatch("Creato foglio BATCH_QUEUE");
  }
  
  // BATCH_JOBS
  var jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
  if (!jobsSheet) {
    jobsSheet = ss.insertSheet(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    jobsSheet.setTabColor("#7E57C2"); // Viola
    
    var jobsHeaders = Object.values(BATCH_CONFIG.COLONNE_JOBS);
    jobsSheet.getRange(1, 1, 1, jobsHeaders.length)
      .setValues([jobsHeaders])
      .setFontWeight("bold")
      .setBackground("#EDE7F6");
    jobsSheet.setFrozenRows(1);
    
    logBatch("Creato foglio BATCH_JOBS");
  }
  
  return { queueSheet: queueSheet, jobsSheet: jobsSheet };
}

// ═══════════════════════════════════════════════════════════════════════
// ACCODAMENTO RICHIESTE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiunge una richiesta alla coda batch
 * 
 * @param {Object} request - {
 *   tipoOperazione: String,
 *   customId: String,
 *   systemPrompt: String,
 *   userPrompt: String,
 *   model: String (optional),
 *   maxTokens: Number (optional),
 *   temperature: Number (optional),
 *   jsonMode: Boolean (optional),
 *   priorita: Number (optional),
 *   metadata: Object (optional)
 * }
 * @returns {Object} {success, requestId, error}
 */
function accodaRichiestaBatch(request) {
  var result = { success: false, requestId: null, error: null };
  
  if (!request.tipoOperazione || !request.customId || !request.systemPrompt || !request.userPrompt) {
    result.error = "Campi obbligatori mancanti: tipoOperazione, customId, systemPrompt, userPrompt";
    return result;
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    
    if (!sheet) {
      initBatchSheets();
      sheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    }
    
    var requestId = generaUUID();
    
    var row = [
      requestId,                                                    // ID_REQUEST
      new Date(),                                                   // TIMESTAMP_CREAZIONE
      request.tipoOperazione,                                       // TIPO_OPERAZIONE
      request.customId,                                             // CUSTOM_ID
      request.systemPrompt,                                         // SYSTEM_PROMPT
      request.userPrompt,                                           // USER_PROMPT
      request.model || BATCH_CONFIG.DEFAULTS.MODEL,                 // MODEL
      request.maxTokens || BATCH_CONFIG.DEFAULTS.MAX_TOKENS,        // MAX_TOKENS
      request.temperature !== undefined ? request.temperature : BATCH_CONFIG.DEFAULTS.TEMPERATURE, // TEMPERATURE
      request.jsonMode !== false ? "SI" : "NO",                     // JSON_MODE
      request.priorita || BATCH_CONFIG.DEFAULTS.PRIORITA_MEDIA,     // PRIORITA
      BATCH_CONFIG.STATUS_QUEUE.PENDING,                            // STATUS
      "",                                                           // BATCH_JOB_ID
      "",                                                           // RISULTATO_JSON
      "",                                                           // TIMESTAMP_COMPLETAMENTO
      "",                                                           // ERRORE
      request.metadata ? JSON.stringify(request.metadata) : ""      // METADATA_JSON
    ];
    
    sheet.appendRow(row);
    
    result.success = true;
    result.requestId = requestId;
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Errore accodamento: " + result.error);
  }
  
  return result;
}

/**
 * Accoda multiple richieste (più efficiente)
 * 
 * @param {Array} requests - Array di oggetti request
 * @returns {Object} {success, count, requestIds, error}
 */
function accodaRichiesteBatch(requests) {
  var result = { success: false, count: 0, requestIds: [], error: null };
  
  if (!requests || requests.length === 0) {
    result.error = "Nessuna richiesta da accodare";
    return result;
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    
    if (!sheet) {
      initBatchSheets();
      sheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    }
    
    var rows = [];
    var now = new Date();
    
    requests.forEach(function(request) {
      if (!request.tipoOperazione || !request.customId || !request.systemPrompt || !request.userPrompt) {
        return; // Skip richieste incomplete
      }
      
      var requestId = generaUUID();
      result.requestIds.push(requestId);
      
      rows.push([
        requestId,
        now,
        request.tipoOperazione,
        request.customId,
        request.systemPrompt,
        request.userPrompt,
        request.model || BATCH_CONFIG.DEFAULTS.MODEL,
        request.maxTokens || BATCH_CONFIG.DEFAULTS.MAX_TOKENS,
        request.temperature !== undefined ? request.temperature : BATCH_CONFIG.DEFAULTS.TEMPERATURE,
        request.jsonMode !== false ? "SI" : "NO",
        request.priorita || BATCH_CONFIG.DEFAULTS.PRIORITA_MEDIA,
        BATCH_CONFIG.STATUS_QUEUE.PENDING,
        "",
        "",
        "",
        "",
        request.metadata ? JSON.stringify(request.metadata) : ""
      ]);
    });
    
    if (rows.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      
      result.success = true;
      result.count = rows.length;
      logBatch("✅ Accodate " + result.count + " richieste");
    }
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Errore accodamento batch: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// INVIO BATCH
// ═══════════════════════════════════════════════════════════════════════

/**
 * Invia richieste pending come batch a OpenAI
 * 
 * @param {Object} options - {
 *   tipoOperazione: String (opzionale, filtra per tipo),
 *   maxRequests: Number (opzionale, default da config),
 *   forza: Boolean (opzionale, ignora MIN_REQUESTS)
 * }
 * @returns {Object} {success, batchJobId, numRequests, error}
 */
function inviaBatchPending(options) {
  var result = { success: false, batchJobId: null, numRequests: 0, error: null };
  
  options = options || {};
  var maxRequests = options.maxRequests || BATCH_CONFIG.DEFAULTS.MAX_REQUESTS_PER_BATCH;
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    var jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    
    if (!queueSheet || !jobsSheet) {
      initBatchSheets();
      queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
      jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    }
    
    // 1. Carica richieste PENDING
    var data = queueSheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    // Mappa colonne
    var colMap = {};
    Object.keys(BATCH_CONFIG.COLONNE_QUEUE).forEach(function(key) {
      colMap[key] = headers.indexOf(BATCH_CONFIG.COLONNE_QUEUE[key]);
    });
    
    // Filtra pending
    var pendingRows = [];
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row[colMap.STATUS] !== BATCH_CONFIG.STATUS_QUEUE.PENDING) continue;
      if (options.tipoOperazione && row[colMap.TIPO_OPERAZIONE] !== options.tipoOperazione) continue;
      
      pendingRows.push({
        rowIndex: i + 2, // +2 per header e 0-based
        data: row,
        colMap: colMap
      });
      
      if (pendingRows.length >= maxRequests) break;
    }
    
    // Check minimo richieste
    if (pendingRows.length === 0) {
      result.error = "Nessuna richiesta pending";
      return result;
    }
    
    if (!options.forza && pendingRows.length < BATCH_CONFIG.DEFAULTS.MIN_REQUESTS_PER_BATCH) {
      result.error = "Richieste insufficienti: " + pendingRows.length + "/" + BATCH_CONFIG.DEFAULTS.MIN_REQUESTS_PER_BATCH;
      return result;
    }
    
    // 2. Costruisci JSONL
    var jsonlRequests = pendingRows.map(function(pr) {
      return {
        customId: pr.data[colMap.CUSTOM_ID],
        systemPrompt: pr.data[colMap.SYSTEM_PROMPT],
        userPrompt: pr.data[colMap.USER_PROMPT],
        options: {
          model: pr.data[colMap.MODEL],
          maxTokens: parseInt(pr.data[colMap.MAX_TOKENS]) || BATCH_CONFIG.DEFAULTS.MAX_TOKENS,
          temperature: parseFloat(pr.data[colMap.TEMPERATURE]) || BATCH_CONFIG.DEFAULTS.TEMPERATURE,
          jsonMode: pr.data[colMap.JSON_MODE] === "SI"
        }
      };
    });
    
    var jsonlContent = buildBatchJsonl(jsonlRequests);
    
    // 3. Upload file
    logBatch("Uploading " + pendingRows.length + " requests...");
    var uploadResult = uploadBatchFile(jsonlContent);
    
    if (!uploadResult.success) {
      result.error = "Upload fallito: " + uploadResult.error;
      return result;
    }
    
    // 4. Crea batch
    var tipoPrevalente = options.tipoOperazione || pendingRows[0].data[colMap.TIPO_OPERAZIONE];
    var createResult = createBatch(uploadResult.fileId, {
      metadata: { tipo: tipoPrevalente, count: pendingRows.length }
    });
    
    if (!createResult.success) {
      result.error = "Creazione batch fallita: " + createResult.error;
      return result;
    }
    
    // 5. Registra job
    var batchJobId = generaBatchJobId();
    var costoStimato = stimaCostoBatch(pendingRows.length, pendingRows[0].data[colMap.MODEL]);
    
    var jobRow = [
      batchJobId,                           // BATCH_JOB_ID
      createResult.batchId,                 // OPENAI_BATCH_ID
      uploadResult.fileId,                  // OPENAI_FILE_ID
      new Date(),                           // TIMESTAMP_INVIO
      "",                                   // TIMESTAMP_COMPLETAMENTO
      tipoPrevalente,                       // TIPO_OPERAZIONE
      pendingRows.length,                   // NUM_REQUESTS
      BATCH_CONFIG.STATUS_JOB.SUBMITTED,    // STATUS
      createResult.status || "validating",  // OPENAI_STATUS
      "",                                   // OUTPUT_FILE_ID
      "",                                   // ERROR_FILE_ID
      0,                                    // COMPLETED_COUNT
      0,                                    // FAILED_COUNT
      "",                                   // ERRORE_MSG
      costoStimato.totalCost,               // COSTO_STIMATO
      ""                                    // NOTE
    ];
    
    jobsSheet.appendRow(jobRow);
    
    // 6. Aggiorna status richieste
    pendingRows.forEach(function(pr) {
      queueSheet.getRange(pr.rowIndex, colMap.STATUS + 1).setValue(BATCH_CONFIG.STATUS_QUEUE.PROCESSING);
      queueSheet.getRange(pr.rowIndex, colMap.BATCH_JOB_ID + 1).setValue(batchJobId);
    });
    
    result.success = true;
    result.batchJobId = batchJobId;
    result.openAiBatchId = createResult.batchId;
    result.numRequests = pendingRows.length;
    
    logBatch("✅ Batch inviato: " + batchJobId + " con " + result.numRequests + " richieste");
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Errore invio batch: " + result.error);
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// POLLING STATUS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Controlla status di tutti i job attivi
 * 
 * @returns {Object} {checked, completed, inProgress, failed}
 */
function pollAllBatchJobs() {
  var result = { checked: 0, completed: [], inProgress: [], failed: [] };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    
    if (!jobsSheet || jobsSheet.getLastRow() <= 1) {
      return result;
    }
    
    var data = jobsSheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    // Mappa colonne
    var colMap = {};
    Object.keys(BATCH_CONFIG.COLONNE_JOBS).forEach(function(key) {
      colMap[key] = headers.indexOf(BATCH_CONFIG.COLONNE_JOBS[key]);
    });
    
    // Check job attivi
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var status = row[colMap.STATUS];
      
      // Skip job già completati o falliti
      if (status === BATCH_CONFIG.STATUS_JOB.COMPLETED ||
          status === BATCH_CONFIG.STATUS_JOB.FAILED ||
          status === BATCH_CONFIG.STATUS_JOB.EXPIRED ||
          status === BATCH_CONFIG.STATUS_JOB.CANCELLED) {
        continue;
      }
      
      var openAiBatchId = row[colMap.OPENAI_BATCH_ID];
      if (!openAiBatchId) continue;
      
      var rowNum = i + 2;
      result.checked++;
      
      // Check status
      var statusResult = checkBatchStatus(openAiBatchId);
      
      if (!statusResult.success) {
        continue; // Errore API, riprova dopo
      }
      
      // Aggiorna foglio
      jobsSheet.getRange(rowNum, colMap.OPENAI_STATUS + 1).setValue(statusResult.status);
      
      if (statusResult.counts) {
        jobsSheet.getRange(rowNum, colMap.COMPLETED_COUNT + 1).setValue(statusResult.counts.completed);
        jobsSheet.getRange(rowNum, colMap.FAILED_COUNT + 1).setValue(statusResult.counts.failed);
      }
      
      // Determina nuovo status interno
      var newStatus = mapOpenAiStatusToInternal(statusResult.status);
      jobsSheet.getRange(rowNum, colMap.STATUS + 1).setValue(newStatus);
      
      if (statusResult.outputFileId) {
        jobsSheet.getRange(rowNum, colMap.OUTPUT_FILE_ID + 1).setValue(statusResult.outputFileId);
      }
      
      if (statusResult.errorFileId) {
        jobsSheet.getRange(rowNum, colMap.ERROR_FILE_ID + 1).setValue(statusResult.errorFileId);
      }
      
      // Categorizza
      var batchJobId = row[colMap.BATCH_JOB_ID];
      
      if (newStatus === BATCH_CONFIG.STATUS_JOB.COMPLETED) {
        jobsSheet.getRange(rowNum, colMap.TIMESTAMP_COMPLETAMENTO + 1).setValue(new Date());
        result.completed.push({
          batchJobId: batchJobId,
          openAiBatchId: openAiBatchId,
          outputFileId: statusResult.outputFileId
        });
      } else if (newStatus === BATCH_CONFIG.STATUS_JOB.FAILED || newStatus === BATCH_CONFIG.STATUS_JOB.EXPIRED) {
        result.failed.push({
          batchJobId: batchJobId,
          openAiBatchId: openAiBatchId,
          status: newStatus
        });
      } else {
        result.inProgress.push({
          batchJobId: batchJobId,
          openAiBatchId: openAiBatchId,
          status: newStatus,
          progress: statusResult.counts
        });
      }
    }
    
    logBatch("📊 Poll: " + result.checked + " checked, " + 
             result.completed.length + " completed, " +
             result.inProgress.length + " in progress, " +
             result.failed.length + " failed");
    
  } catch (e) {
    logBatch("❌ Errore polling: " + e.toString());
  }
  
  return result;
}

/**
 * Mappa status OpenAI a status interno
 */
function mapOpenAiStatusToInternal(openAiStatus) {
  switch (openAiStatus) {
    case "validating":
    case "in_progress":
      return BATCH_CONFIG.STATUS_JOB.IN_PROGRESS;
    case "completed":
      return BATCH_CONFIG.STATUS_JOB.COMPLETED;
    case "failed":
      return BATCH_CONFIG.STATUS_JOB.FAILED;
    case "expired":
      return BATCH_CONFIG.STATUS_JOB.EXPIRED;
    case "cancelled":
    case "cancelling":
      return BATCH_CONFIG.STATUS_JOB.CANCELLED;
    default:
      return BATCH_CONFIG.STATUS_JOB.SUBMITTED;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// RECUPERO RISULTATI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Recupera e processa risultati di un batch completato
 * 
 * @param {String} batchJobId - ID job interno
 * @returns {Object} {success, processed, errors, results}
 */
function recuperaRisultatiBatch(batchJobId) {
  var result = { success: false, processed: 0, errors: 0, results: {}, error: null };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    var queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    
    if (!jobsSheet || !queueSheet) {
      result.error = "Fogli batch non trovati";
      return result;
    }
    
    // 1. Trova job
    var jobsData = jobsSheet.getDataRange().getValues();
    var jobsHeaders = jobsData[0];
    
    var colMapJobs = {};
    Object.keys(BATCH_CONFIG.COLONNE_JOBS).forEach(function(key) {
      colMapJobs[key] = jobsHeaders.indexOf(BATCH_CONFIG.COLONNE_JOBS[key]);
    });
    
    var jobRow = null;
    var jobRowNum = -1;
    
    for (var i = 1; i < jobsData.length; i++) {
      if (jobsData[i][colMapJobs.BATCH_JOB_ID] === batchJobId) {
        jobRow = jobsData[i];
        jobRowNum = i + 1;
        break;
      }
    }
    
    if (!jobRow) {
      result.error = "Job non trovato: " + batchJobId;
      return result;
    }
    
    var outputFileId = jobRow[colMapJobs.OUTPUT_FILE_ID];
    if (!outputFileId) {
      result.error = "Output file non disponibile";
      return result;
    }
    
    // 2. Download risultati
    var downloadResult = downloadBatchFile(outputFileId);
    
    if (!downloadResult.success) {
      result.error = "Download fallito: " + downloadResult.error;
      return result;
    }
    
    // 3. Parse risultati
    var parsedResults = parseBatchResponses(downloadResult.lines);
    
    // 4. Aggiorna BATCH_QUEUE con risultati
    var queueData = queueSheet.getDataRange().getValues();
    var queueHeaders = queueData[0];
    
    var colMapQueue = {};
    Object.keys(BATCH_CONFIG.COLONNE_QUEUE).forEach(function(key) {
      colMapQueue[key] = queueHeaders.indexOf(BATCH_CONFIG.COLONNE_QUEUE[key]);
    });
    
    for (var j = 1; j < queueData.length; j++) {
      var qRow = queueData[j];
      
      if (qRow[colMapQueue.BATCH_JOB_ID] !== batchJobId) continue;
      
      var customId = qRow[colMapQueue.CUSTOM_ID];
      var parsed = parsedResults[customId];
      var qRowNum = j + 1;
      
      if (parsed) {
        result.results[customId] = parsed;
        
        if (parsed.success) {
          queueSheet.getRange(qRowNum, colMapQueue.STATUS + 1).setValue(BATCH_CONFIG.STATUS_QUEUE.COMPLETED);
          queueSheet.getRange(qRowNum, colMapQueue.RISULTATO_JSON + 1).setValue(
            typeof parsed.parsedContent === 'object' ? JSON.stringify(parsed.parsedContent) : parsed.content
          );
          queueSheet.getRange(qRowNum, colMapQueue.TIMESTAMP_COMPLETAMENTO + 1).setValue(new Date());
          result.processed++;
        } else {
          queueSheet.getRange(qRowNum, colMapQueue.STATUS + 1).setValue(BATCH_CONFIG.STATUS_QUEUE.ERROR);
          queueSheet.getRange(qRowNum, colMapQueue.ERRORE + 1).setValue(parsed.error || "Errore sconosciuto");
          result.errors++;
        }
      }
    }
    
    result.success = true;
    logBatch("✅ Recuperati risultati " + batchJobId + ": " + result.processed + " ok, " + result.errors + " errori");
    
  } catch (e) {
    result.error = e.toString();
    logBatch("❌ Errore recupero risultati: " + result.error);
  }
  
  return result;
}

/**
 * Recupera risultati di tutti i batch completati
 * 
 * @returns {Object} {processed, errors, batches}
 */
function recuperaTuttiRisultatiCompletati() {
  var result = { processed: 0, errors: 0, batches: [] };
  
  // Prima poll per aggiornare status
  var pollResult = pollAllBatchJobs();
  
  // Poi recupera risultati dei completati
  pollResult.completed.forEach(function(job) {
    var recuperoResult = recuperaRisultatiBatch(job.batchJobId);
    
    result.batches.push({
      batchJobId: job.batchJobId,
      success: recuperoResult.success,
      processed: recuperoResult.processed,
      errors: recuperoResult.errors
    });
    
    result.processed += recuperoResult.processed;
    result.errors += recuperoResult.errors;
  });
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// STATISTICHE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Statistiche coda e job
 */
function getStatsBatch() {
  var stats = {
    queue: {
      pending: 0,
      processing: 0,
      completed: 0,
      error: 0,
      totale: 0
    },
    jobs: {
      submitted: 0,
      inProgress: 0,
      completed: 0,
      failed: 0,
      totale: 0
    }
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Queue
    var queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    if (queueSheet && queueSheet.getLastRow() > 1) {
      var qData = queueSheet.getDataRange().getValues();
      var qHeaders = qData[0];
      var colStatus = qHeaders.indexOf(BATCH_CONFIG.COLONNE_QUEUE.STATUS);
      
      for (var i = 1; i < qData.length; i++) {
        var status = qData[i][colStatus];
        stats.queue.totale++;
        
        switch (status) {
          case BATCH_CONFIG.STATUS_QUEUE.PENDING: stats.queue.pending++; break;
          case BATCH_CONFIG.STATUS_QUEUE.PROCESSING:
          case BATCH_CONFIG.STATUS_QUEUE.QUEUED: stats.queue.processing++; break;
          case BATCH_CONFIG.STATUS_QUEUE.COMPLETED: stats.queue.completed++; break;
          case BATCH_CONFIG.STATUS_QUEUE.ERROR: stats.queue.error++; break;
        }
      }
    }
    
    // Jobs
    var jobsSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_JOBS);
    if (jobsSheet && jobsSheet.getLastRow() > 1) {
      var jData = jobsSheet.getDataRange().getValues();
      var jHeaders = jData[0];
      var colStatusJ = jHeaders.indexOf(BATCH_CONFIG.COLONNE_JOBS.STATUS);
      
      for (var j = 1; j < jData.length; j++) {
        var statusJ = jData[j][colStatusJ];
        stats.jobs.totale++;
        
        switch (statusJ) {
          case BATCH_CONFIG.STATUS_JOB.SUBMITTED:
          case BATCH_CONFIG.STATUS_JOB.UPLOADING: stats.jobs.submitted++; break;
          case BATCH_CONFIG.STATUS_JOB.IN_PROGRESS: stats.jobs.inProgress++; break;
          case BATCH_CONFIG.STATUS_JOB.COMPLETED: stats.jobs.completed++; break;
          case BATCH_CONFIG.STATUS_JOB.FAILED:
          case BATCH_CONFIG.STATUS_JOB.EXPIRED:
          case BATCH_CONFIG.STATUS_JOB.CANCELLED: stats.jobs.failed++; break;
        }
      }
    }
    
  } catch (e) {
    logBatch("Errore stats: " + e.toString());
  }
  
  return stats;
}