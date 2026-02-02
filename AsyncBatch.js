/**
 * ==========================================================================================
 * ASYNC_BATCH.gs - Gestione OpenAI Batch API (50% sconto)
 * ==========================================================================================
 */

/**
 * FASE 1: Prepara il file JSONL, lo carica e avvia il Batch
 * Da schedulare come Job Manuale o Settimanale
 */
function avviaBatchBacklog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  var maxEmail = 100; // Batch size
  
  // 1. Trova email da analizzare (Backlog)
  // Utilizza status DA_ANALIZZARE o un nuovo status DA_BATCHARE
  var data = sheet.getDataRange().getValues();
  var requests = [];
  var rowsToUpdate = [];
  
  // Mappatura colonne
  var headers = data;
  var colId = headers.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  var colMitt = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colOgg = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  var colCorpo = headers.indexOf(CONFIG.COLONNE_LOG_IN.CORPO);
  var colStatus = headers.indexOf(CONFIG.COLONNE_LOG_IN.STATUS);
  
  // Prompt di sistema (preso da CONFIG o AILayer)
  // [2] Riferimento al prompt Layer 1 esistente
  var systemPrompt = "Sei un analista esperto di email commerciali B2B. " +
    "Analizza l'email e rispondi SOLO in formato JSON valido con: tags, sintesi, scores.";

  for (var i = 1; i < data.length && requests.length < maxEmail; i++) {
    var row = data[i];
    // Filtra per backlog (es. vecchie o status specifico)
    if (row[colStatus] === "DA_ANALIZZARE") { 
      
      var emailId = row[colId];
      var userPrompt = "Analizza questa email:\nMITTENTE: " + row[colMitt] + 
                       "\nOGGETTO: " + row[colOgg] + 
                       "\nCORPO: " + row[colCorpo];

      // Costruzione Oggetto Request per Batch API
      var requestObj = {
        "custom_id": emailId, // FONDAMENTALE per ricollegare la risposta
        "method": "POST",
        "url": "/v1/chat/completions",
        "body": {
          "model": "gpt-4o", // O gpt-4o-mini per risparmiare ancora di più
          "messages": [
            {"role": "system", "content": systemPrompt},
            {"role": "user", "content": userPrompt}
          ],
          "response_format": { "type": "json_object" },
          "max_tokens": 500
        }
      };
      
      requests.push(JSON.stringify(requestObj));
      rowsToUpdate.push(i + 1);
    }
  }
  
  if (requests.length === 0) {
    logSistema("⚠️ Nessuna email trovata per il Batch Backlog");
    return;
  }

  // 2. Crea file JSONL
  var jsonlContent = requests.join("\n");
  var blob = Utilities.newBlob(jsonlContent, "application/jsonl", "batch_input.jsonl");
  
  // 3. Upload File a OpenAI
  var fileId = uploadFileToOpenAI(blob);
  if (!fileId) return; // Errore loggato in helper
  
  // 4. Avvia Batch
  var batchId = startOpenAIBatch(fileId);
  
  if (batchId) {
    // Salva ID Batch e Righe coinvolte nelle Properties o in un foglio tecnico
    var batchData = {
      id: batchId,
      timestamp: new Date(),
      count: requests.length,
      status: "in_progress"
    };
    
    // Aggiorna status righe in "BATCH_IN_CORSO" per non riprocessarle
    rowsToUpdate.forEach(r => {
      sheet.getRange(r, colStatus + 1).setValue("BATCH_IN_CORSO");
    });
    
    // Salviamo l'ID batch per il recupero successivo
    PropertiesService.getScriptProperties().setProperty("ACTIVE_BATCH_ID", batchId);
    logSistema("🚀 Batch avviato: " + batchId + " con " + requests.length + " email.");
  }
}

/**
 * FASE 2: Controlla lo stato e scarica i risultati
 * Da schedulare ogni ora
 */
function controllaScaricaBatch() {
  var batchId = PropertiesService.getScriptProperties().getProperty("ACTIVE_BATCH_ID");
  if (!batchId) return; // Nessun batch attivo
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  // 1. Check Status
  var url = "https://api.openai.com/v1/batches/" + batchId;
  var options = {
    "method": "get",
    "headers": {"Authorization": "Bearer " + apiKey}
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  
  logSistema("🔄 Stato Batch " + batchId + ": " + json.status);
  
  if (json.status === "completed") {
    // 2. Download Risultati
    var outputFileId = json.output_file_id;
    if (outputFileId) {
      processaRisultatiBatch(outputFileId);
      PropertiesService.getScriptProperties().deleteProperty("ACTIVE_BATCH_ID"); // Pulizia
    }
  } else if (json.status === "failed" || json.status === "expired") {
    logSistema("❌ Batch fallito: " + JSON.stringify(json.errors));
    PropertiesService.getScriptProperties().deleteProperty("ACTIVE_BATCH_ID");
    // Qui servirebbe logica per resettare lo status delle email a "DA_ANALIZZARE"
  }
}

/**
 * Helper: Upload file .jsonl (Multipart workaround per GAS)
 */
function uploadFileToOpenAI(blob) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  // OpenAI richiede multipart/form-data. In GAS è automatico se payload ha blob
  var payload = {
    "purpose": "batch",
    "file": blob
  };
  
  var options = {
    "method": "post",
    "headers": {"Authorization": "Bearer " + apiKey},
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/files", options);
  var json = JSON.parse(response.getContentText());
  
  if (json.id) return json.id;
  logSistema("❌ Errore Upload File: " + response.getContentText());
  return null;
}

/**
 * Helper: Start Batch Endpoint
 */
function startOpenAIBatch(inputFileId) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  var payload = {
    "input_file_id": inputFileId,
    "endpoint": "/v1/chat/completions",
    "completion_window": "24h" // Finestra temporale obbligatoria per sconto 50%
  };
  
  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/batches", options);
  var json = JSON.parse(response.getContentText());
  
  if (json.id) return json.id;
  logSistema("❌ Errore Start Batch: " + response.getContentText());
  return null;
}

/**
 * Processa il file dei risultati e aggiorna il foglio
 */
function processaRisultatiBatch(fileId) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  
  // Download contenuto file
  var url = "https://api.openai.com/v1/files/" + fileId + "/content";
  var response = UrlFetchApp.fetch(url, {
    "headers": {"Authorization": "Bearer " + apiKey}
  });
  
  var content = response.getContentText();
  var lines = content.split("\n");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  var data = sheet.getDataRange().getValues();
  
  // Crea mappa ID -> Row Index per velocità
  var idMap = {};
  var colIdIndex = data.indexOf(CONFIG.COLONNE_LOG_IN.ID_EMAIL);
  for (var i = 1; i < data.length; i++) {
    idMap[data[i][colIdIndex]] = i + 1;
  }
  
  logSistema("📥 Elaborazione " + lines.length + " risultati Batch...");
  
  lines.forEach(function(line) {
    if (!line) return;
    try {
      var res = JSON.parse(line);
      var customId = res.custom_id; // ID EMAIL
      var resultBody = res.response.body; // Risultato GPT
      
      var rowNum = idMap[customId];
      
      if (rowNum && resultBody.choices) {
        var contentAI = resultBody.choices.message.content;
        var parsedAI = JSON.parse(contentAI);
        
        // Costruisci oggetto risultato compatibile con il tuo sistema [3]
        var risultatoL1 = {
          tags: parsedAI.tags || [],
          sintesi: parsedAI.sintesi || "",
          scores: parsedAI.scores || {},
          modello: "gpt-4o-batch"
        };
        
        // Scrivi nel foglio usando le tue funzioni esistenti in WRITERS.gs
        scriviRisultatiL1(sheet, rowNum, risultatoL1);
        aggiornaStatus(sheet, rowNum, "ANALIZZATO"); // O "NEEDS_REVIEW"
      }
    } catch (e) {
      logSistema("Errore parsing riga batch: " + e.message);
    }
  });
  
  logSistema("✅ Batch completato e dati importati.");
}