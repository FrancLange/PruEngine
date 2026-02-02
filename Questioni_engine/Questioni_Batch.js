/**
 * ==========================================================================================
 * QUESTIONI_BATCH.js - Batch Processing per Questioni v1.0.0
 * ==========================================================================================
 * Usa l'infrastruttura batch per:
 * - Pre-filtraggio email candidati (locale, no AI)
 * - Clustering email per fornitore (batch AI)
 * - Matching con questioni esistenti (batch AI)
 * - Creazione questioni dai risultati
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// FLOW PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Flow completo batch per questioni (backlog)
 * 
 * FASI:
 * 1. Scan email candidati (locale) → pre-filtra score_problema>70 o urgente>80
 * 2. Raggruppa per fornitore (locale)
 * 3. Accoda richieste clustering (batch)
 * 4. Accoda richieste matching (batch)
 * 5. Invia batch → Aspetta completamento
 * 6. Processa risultati → Crea questioni
 * 
 * @param {Object} params - {
 *   giorniIndietro: Number (default 90),
 *   maxEmail: Number (default 500),
 *   soloAccoda: Boolean (non invia subito)
 * }
 * @returns {Object} Report
 */
function avviaQuestioniBatch(params) {
  params = params || {};
  var giorniIndietro = params.giorniIndietro || 90;
  var maxEmail = params.maxEmail || 500;
  var soloAccoda = params.soloAccoda || false;
  
  var report = {
    fase1_scan: { candidati: 0, filtrati: 0 },
    fase2_gruppi: { fornitori: 0 },
    fase3_clustering: { richieste: 0 },
    fase4_matching: { richieste: 0 },
    fase5_invio: { success: false, batchJobId: null },
    errori: []
  };
  
  logBatch("═══ QUESTIONI BATCH START ═══");
  logBatch("Parametri: " + giorniIndietro + " giorni, max " + maxEmail + " email");
  
  try {
    // ═══════════════════════════════════════════════════════════
    // FASE 1: SCAN EMAIL CANDIDATI
    // ═══════════════════════════════════════════════════════════
    logBatch("📌 FASE 1: Scan email candidati...");
    
    var candidati = scanEmailPerQuestioniBatch({
      giorniIndietro: giorniIndietro,
      maxEmail: maxEmail
    });
    
    report.fase1_scan.candidati = candidati.length;
    logBatch("Trovate " + candidati.length + " email candidati");
    
    if (candidati.length === 0) {
      report.errori.push("Nessuna email candidata trovata");
      return report;
    }
    
    // ═══════════════════════════════════════════════════════════
    // FASE 2: RAGGRUPPA PER FORNITORE
    // ═══════════════════════════════════════════════════════════
    logBatch("📌 FASE 2: Raggruppamento per fornitore...");
    
    var gruppi = raggruppaEmailPerFornitoreBatch(candidati);
    report.fase2_gruppi.fornitori = Object.keys(gruppi).length;
    logBatch("Creati " + report.fase2_gruppi.fornitori + " gruppi");
    
    // ═══════════════════════════════════════════════════════════
    // FASE 3: ACCODA CLUSTERING
    // ═══════════════════════════════════════════════════════════
    logBatch("📌 FASE 3: Accodamento richieste clustering...");
    
    var clusteringRequests = creaRichiesteClusteringBatch(gruppi);
    report.fase3_clustering.richieste = clusteringRequests.length;
    
    if (clusteringRequests.length > 0) {
      var accodaCluster = accodaRichiesteBatch(clusteringRequests);
      if (!accodaCluster.success) {
        report.errori.push("Errore accodamento clustering: " + accodaCluster.error);
      }
    }
    
    logBatch("Accodate " + report.fase3_clustering.richieste + " richieste clustering");
    
    // ═══════════════════════════════════════════════════════════
    // FASE 4: ACCODA MATCHING (se ci sono questioni esistenti)
    // ═══════════════════════════════════════════════════════════
    logBatch("📌 FASE 4: Accodamento richieste matching...");
    
    // Carica questioni esistenti
    var questioniAperte = [];
    if (typeof getQuestioniAperte === 'function') {
      questioniAperte = getQuestioniAperte();
    }
    
    if (questioniAperte.length > 0) {
      var matchingRequests = creaRichiesteMatchingBatch(candidati, questioniAperte);
      report.fase4_matching.richieste = matchingRequests.length;
      
      if (matchingRequests.length > 0) {
        var accodaMatch = accodaRichiesteBatch(matchingRequests);
        if (!accodaMatch.success) {
          report.errori.push("Errore accodamento matching: " + accodaMatch.error);
        }
      }
    } else {
      logBatch("Nessuna questione aperta, skip matching");
    }
    
    logBatch("Accodate " + report.fase4_matching.richieste + " richieste matching");
    
    // ═══════════════════════════════════════════════════════════
    // FASE 5: INVIO BATCH (se non soloAccoda)
    // ═══════════════════════════════════════════════════════════
    if (!soloAccoda) {
      logBatch("📌 FASE 5: Invio batch...");
      
      // Prima clustering
      if (report.fase3_clustering.richieste > 0) {
        var invioCluster = inviaBatchPending({
          tipoOperazione: BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_CLUSTER,
          forza: true // Invia anche se poche richieste
        });
        
        if (invioCluster.success) {
          report.fase5_invio.success = true;
          report.fase5_invio.batchJobId = invioCluster.batchJobId;
          logBatch("✅ Batch clustering inviato: " + invioCluster.batchJobId);
        } else {
          report.errori.push("Invio clustering fallito: " + invioCluster.error);
        }
      }
      
      // Poi matching (se c'erano)
      if (report.fase4_matching.richieste > 0) {
        var invioMatch = inviaBatchPending({
          tipoOperazione: BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_MATCH,
          forza: true
        });
        
        if (invioMatch.success) {
          report.fase5_invio.matchBatchJobId = invioMatch.batchJobId;
          logBatch("✅ Batch matching inviato: " + invioMatch.batchJobId);
        } else {
          report.errori.push("Invio matching fallito: " + invioMatch.error);
        }
      }
    } else {
      logBatch("⏸️ Solo accodamento, invio manuale richiesto");
    }
    
  } catch (e) {
    report.errori.push("Errore critico: " + e.toString());
    logBatch("❌ Errore critico: " + e.toString());
  }
  
  logBatch("═══ QUESTIONI BATCH END ═══");
  logBatch("Report: " + JSON.stringify(report));
  
  return report;
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 1: SCAN EMAIL
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scan email da LOG_IN per candidati questioni
 * Pre-filtra localmente senza AI
 */
function scanEmailPerQuestioniBatch(params) {
  params = params || {};
  var giorniIndietro = params.giorniIndietro || 90;
  var maxEmail = params.maxEmail || 500;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Mappa colonne
  var colMap = {};
  Object.keys(CONFIG.COLONNE_LOG_IN).forEach(function(key) {
    var idx = headers.indexOf(CONFIG.COLONNE_LOG_IN[key]);
    if (idx >= 0) colMap[key] = idx;
  });
  
  // Data limite
  var dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - giorniIndietro);
  
  var candidati = [];
  
  // Soglie da QUESTIONI_CONFIG
  var sogliaProblema = typeof QUESTIONI_CONFIG !== 'undefined' ? 
    QUESTIONI_CONFIG.TRIGGER.SOGLIA_PROBLEMA : 70;
  var sogliaUrgente = typeof QUESTIONI_CONFIG !== 'undefined' ? 
    QUESTIONI_CONFIG.TRIGGER.SOGLIA_URGENTE : 80;
  
  for (var i = 0; i < rows.length && candidati.length < maxEmail; i++) {
    var row = rows[i];
    
    // Skip se già collegata a questione
    if (colMap.ID_QUESTIONE_COLLEGATA && row[colMap.ID_QUESTIONE_COLLEGATA]) {
      continue;
    }
    
    // Check data
    var timestamp = row[colMap.TIMESTAMP];
    if (!timestamp || new Date(timestamp) < dataLimite) {
      continue;
    }
    
    // Check status (solo ANALIZZATO o NEEDS_REVIEW)
    var status = row[colMap.STATUS];
    if (status !== "ANALIZZATO" && status !== "NEEDS_REVIEW") {
      continue;
    }
    
    // Parse scores
    var scoresJson = row[colMap.L3_SCORES_JSON] || row[colMap.L1_SCORES_JSON];
    var scores = {};
    if (scoresJson) {
      try { scores = JSON.parse(scoresJson); } catch(e) {}
    }
    
    // Check soglie
    var scoreProblema = scores.problema || 0;
    var scoreUrgente = scores.urgente || 0;
    
    if (scoreProblema >= sogliaProblema || scoreUrgente >= sogliaUrgente) {
      candidati.push({
        rowNum: i + 2,
        id: row[colMap.ID_EMAIL],
        timestamp: row[colMap.TIMESTAMP],
        mittente: row[colMap.MITTENTE],
        oggetto: row[colMap.OGGETTO],
        corpo: row[colMap.CORPO],
        tags: row[colMap.L3_TAGS_FINALI] || row[colMap.L1_TAGS] || "",
        sintesi: row[colMap.L3_SINTESI_FINALE] || row[colMap.L1_SINTESI] || "",
        scores: scores,
        scoreProblema: scoreProblema,
        scoreUrgente: scoreUrgente
      });
    }
  }
  
  // Ordina per data (più vecchie prima per backlog)
  candidati.sort(function(a, b) {
    return new Date(a.timestamp) - new Date(b.timestamp);
  });
  
  return candidati;
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 2: RAGGRUPPAMENTO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Raggruppa email per fornitore/dominio
 */
function raggruppaEmailPerFornitoreBatch(emails) {
  var gruppi = {};
  
  emails.forEach(function(email) {
    // Determina chiave gruppo
    var chiave = null;
    var fornitore = null;
    
    // Prova lookup fornitore se disponibile
    if (typeof lookupFornitoreByEmail === 'function') {
      fornitore = lookupFornitoreByEmail(email.mittente);
      if (fornitore) {
        chiave = "FOR:" + fornitore.id;
      }
    }
    
    // Fallback su dominio
    if (!chiave) {
      var dominio = estraiDominioBatch(email.mittente);
      chiave = "DOM:" + dominio;
    }
    
    // Aggiungi al gruppo
    if (!gruppi[chiave]) {
      gruppi[chiave] = {
        chiave: chiave,
        fornitore: fornitore,
        emails: []
      };
    }
    
    gruppi[chiave].emails.push(email);
  });
  
  return gruppi;
}

/**
 * Estrae dominio da email
 */
function estraiDominioBatch(email) {
  if (!email) return "unknown";
  var at = email.indexOf("@");
  return at > 0 ? email.substring(at + 1).toLowerCase() : "unknown";
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 3: RICHIESTE CLUSTERING
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea richieste batch per clustering
 */
function creaRichiesteClusteringBatch(gruppi) {
  var requests = [];
  
  // Carica prompt da PROMPTS sheet
  var promptData = null;
  if (typeof caricaPrompt === 'function') {
    promptData = caricaPrompt("QUESTIONI_AI_CLUSTER");
  }
  
  var systemPrompt = promptData ? promptData.testo : getPromptClusterBatchFallback();
  
  Object.keys(gruppi).forEach(function(chiave) {
    var gruppo = gruppi[chiave];
    
    // Skip gruppi con 1 sola email (non serve clustering)
    if (gruppo.emails.length === 1) {
      // Per email singole, crea comunque una "richiesta cluster" semplificata
      // che restituirà un cluster singolo
    }
    
    // Prepara lista email per prompt
    var emailList = gruppo.emails.map(function(e, idx) {
      return (idx + 1) + ". [" + e.id + "] " + e.oggetto + "\n   Sintesi: " + (e.sintesi || "N/D").substring(0, 100);
    }).join("\n\n");
    
    var fornitoreInfo = gruppo.fornitore ? 
      "Fornitore: " + gruppo.fornitore.nome + " (ID: " + gruppo.fornitore.id + ")" :
      "Dominio: " + chiave.replace("DOM:", "");
    
    var userPrompt = "CONTESTO:\n" + fornitoreInfo + "\n\n" +
      "EMAIL DA RAGGRUPPARE (" + gruppo.emails.length + "):\n\n" + emailList + "\n\n" +
      "Analizza e raggruppa le email in questioni/ticket logici.";
    
    requests.push({
      tipoOperazione: BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_CLUSTER,
      customId: chiave,
      systemPrompt: systemPrompt,
      userPrompt: userPrompt,
      model: BATCH_CONFIG.DEFAULTS.MODEL,
      maxTokens: 1000,
      temperature: 0.3,
      jsonMode: true,
      priorita: BATCH_CONFIG.DEFAULTS.PRIORITA_MEDIA,
      metadata: {
        tipo: "clustering",
        fornitoreId: gruppo.fornitore ? gruppo.fornitore.id : null,
        emailCount: gruppo.emails.length,
        emailIds: gruppo.emails.map(function(e) { return e.id; })
      }
    });
  });
  
  return requests;
}

/**
 * Prompt fallback per clustering
 */
function getPromptClusterBatchFallback() {
  return "Sei un assistente per la gestione ticket/questioni B2B.\n\n" +
    "Ti vengono fornite email da uno stesso fornitore. Devi raggrupparle in questioni logiche.\n\n" +
    "Criteri raggruppamento:\n" +
    "- Stesso problema/argomento → stessa questione\n" +
    "- Email di follow-up su stesso tema → stessa questione\n" +
    "- Problemi diversi → questioni separate\n\n" +
    "Rispondi SOLO in JSON:\n" +
    "{\n" +
    '  "clusters": [\n' +
    '    {\n' +
    '      "email_ids": ["ID1", "ID2"],\n' +
    '      "titolo": "Titolo questione breve",\n' +
    '      "descrizione": "Descrizione problema max 100 parole",\n' +
    '      "categoria": "RECLAMO|RITARDO|QUALITA|PAGAMENTO|COMUNICAZIONE|CONTRATTO|ALTRO",\n' +
    '      "urgenza": "ALTA|MEDIA|BASSA"\n' +
    '    }\n' +
    '  ],\n' +
    '  "note": "Eventuali note"\n' +
    '}';
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 4: RICHIESTE MATCHING
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea richieste batch per matching con questioni esistenti
 */
function creaRichiesteMatchingBatch(emails, questioniAperte) {
  var requests = [];
  
  if (!questioniAperte || questioniAperte.length === 0) {
    return requests;
  }
  
  // Carica prompt
  var promptData = null;
  if (typeof caricaPrompt === 'function') {
    promptData = caricaPrompt("QUESTIONI_AI_MATCH");
  }
  
  var systemPrompt = promptData ? promptData.testo : getPromptMatchBatchFallback();
  
  // Prepara lista questioni aperte (comune a tutte le richieste)
  var questioniList = questioniAperte.map(function(q) {
    return "- [" + q.id + "] " + q.titolo + " | " + q.categoria + " | Fornitore: " + (q.fornitore || "N/D");
  }).join("\n");
  
  emails.forEach(function(email) {
    var userPrompt = "QUESTIONI APERTE:\n" + questioniList + "\n\n" +
      "NUOVA EMAIL:\n" +
      "ID: " + email.id + "\n" +
      "Mittente: " + email.mittente + "\n" +
      "Oggetto: " + email.oggetto + "\n" +
      "Sintesi: " + (email.sintesi || "N/D") + "\n" +
      "Tags: " + (email.tags || "N/D") + "\n\n" +
      "L'email corrisponde a una questione esistente?";
    
    requests.push({
      tipoOperazione: BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_MATCH,
      customId: email.id,
      systemPrompt: systemPrompt,
      userPrompt: userPrompt,
      model: BATCH_CONFIG.DEFAULTS.MODEL_MINI, // Mini per matching (più economico)
      maxTokens: 300,
      temperature: 0.2,
      jsonMode: true,
      priorita: BATCH_CONFIG.DEFAULTS.PRIORITA_MEDIA,
      metadata: {
        tipo: "matching",
        mittente: email.mittente,
        scoreProblema: email.scoreProblema,
        scoreUrgente: email.scoreUrgente
      }
    });
  });
  
  return requests;
}

/**
 * Prompt fallback per matching
 */
function getPromptMatchBatchFallback() {
  return "Sei un assistente per la gestione ticket/questioni B2B.\n\n" +
    "Data una lista di questioni aperte e una nuova email, determina se l'email " +
    "corrisponde a una questione esistente.\n\n" +
    "Criteri:\n" +
    "- Stesso fornitore E stesso problema → MATCH\n" +
    "- Follow-up su questione esistente → MATCH\n" +
    "- Nuovo problema → NO MATCH\n\n" +
    "Rispondi SOLO in JSON:\n" +
    "{\n" +
    '  "match": true/false,\n' +
    '  "id_questione": "QST-XXX o null",\n' +
    '  "confidence": 0-100,\n' +
    '  "motivo": "Spiegazione breve"\n' +
    '}';
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 6: PROCESSA RISULTATI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa risultati batch e crea questioni
 * Da chiamare dopo che il batch è completato
 */
function processaRisultatiBatchQuestioni() {
  var report = {
    clustering: { processati: 0, questioni: 0 },
    matching: { processati: 0, collegati: 0, nuovi: 0 },
    errori: []
  };
  
  logBatch("═══ PROCESSA RISULTATI QUESTIONI ═══");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var queueSheet = ss.getSheetByName(BATCH_CONFIG.SHEETS.BATCH_QUEUE);
    
    if (!queueSheet || queueSheet.getLastRow() <= 1) {
      report.errori.push("BATCH_QUEUE vuoto");
      return report;
    }
    
    var data = queueSheet.getDataRange().getValues();
    var headers = data[0];
    
    var colMap = {};
    Object.keys(BATCH_CONFIG.COLONNE_QUEUE).forEach(function(key) {
      colMap[key] = headers.indexOf(BATCH_CONFIG.COLONNE_QUEUE[key]);
    });
    
    // Processa risultati completati
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Solo COMPLETED
      if (row[colMap.STATUS] !== BATCH_CONFIG.STATUS_QUEUE.COMPLETED) continue;
      
      var tipo = row[colMap.TIPO_OPERAZIONE];
      var customId = row[colMap.CUSTOM_ID];
      var risultatoJson = row[colMap.RISULTATO_JSON];
      var metadataJson = row[colMap.METADATA_JSON];
      
      if (!risultatoJson) continue;
      
      var risultato = null;
      var metadata = null;
      
      try {
        risultato = JSON.parse(risultatoJson);
        if (metadataJson) metadata = JSON.parse(metadataJson);
      } catch(e) {
        report.errori.push("Parse error row " + (i+1));
        continue;
      }
      
      // Processa per tipo
      if (tipo === BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_CLUSTER) {
        report.clustering.processati++;
        var qCreate = processaRisultatoClustering(customId, risultato, metadata);
        report.clustering.questioni += qCreate;
        
      } else if (tipo === BATCH_CONFIG.TIPI_OPERAZIONE.QUESTIONI_MATCH) {
        report.matching.processati++;
        var matchResult = processaRisultatoMatching(customId, risultato, metadata);
        if (matchResult.collegato) {
          report.matching.collegati++;
        } else if (matchResult.nuovo) {
          report.matching.nuovi++;
        }
      }
      
      // Marca come processato (aggiungi flag nel metadata o in una colonna dedicata)
      // Per ora lasciamo COMPLETED
    }
    
  } catch (e) {
    report.errori.push("Errore critico: " + e.toString());
    logBatch("❌ Errore: " + e.toString());
  }
  
  logBatch("Report: clustering " + report.clustering.questioni + " questioni, " +
           "matching " + report.matching.collegati + " collegati, " + report.matching.nuovi + " nuovi");
  
  return report;
}

/**
 * Processa singolo risultato clustering
 */
function processaRisultatoClustering(customId, risultato, metadata) {
  var questioniCreate = 0;
  
  if (!risultato.clusters || !Array.isArray(risultato.clusters)) {
    return 0;
  }
  
  var emailIds = metadata ? metadata.emailIds : [];
  var fornitoreId = metadata ? metadata.fornitoreId : null;
  
  risultato.clusters.forEach(function(cluster) {
    if (!cluster.email_ids || cluster.email_ids.length === 0) return;
    
    // Crea questione per ogni cluster
    if (typeof creaQuestione === 'function') {
      var datiQuestione = {
        titolo: cluster.titolo || "Questione da batch",
        descrizione: cluster.descrizione || "",
        categoria: cluster.categoria || "ALTRO",
        priorita: cluster.urgenza === "ALTA" ? "ALTA" : 
                  cluster.urgenza === "BASSA" ? "BASSA" : "MEDIA",
        origine: "BATCH_BACKLOG",
        idFornitore: fornitoreId,
        idEmailOrigine: cluster.email_ids[0] // Prima email
      };
      
      var creaResult = creaQuestione(datiQuestione);
      
      if (creaResult.success) {
        questioniCreate++;
        
        // Collega tutte le email del cluster
        cluster.email_ids.forEach(function(emailId) {
          if (typeof collegaEmailAQuestione === 'function') {
            collegaEmailAQuestione(emailId, creaResult.id, {});
          }
        });
        
        logBatch("Creata questione " + creaResult.id + " con " + cluster.email_ids.length + " email");
      }
    }
  });
  
  return questioniCreate;
}

/**
 * Processa singolo risultato matching
 */
function processaRisultatoMatching(emailId, risultato, metadata) {
  var result = { collegato: false, nuovo: false };
  
  if (risultato.match && risultato.id_questione && risultato.confidence >= 50) {
    // Match trovato
    if (typeof collegaEmailAQuestione === 'function') {
      var collegaResult = collegaEmailAQuestione(emailId, risultato.id_questione, {
        note: "Match batch: " + risultato.motivo + " (conf: " + risultato.confidence + "%)"
      });
      
      if (collegaResult.success) {
        result.collegato = true;
        logBatch("Collegata " + emailId + " → " + risultato.id_questione);
      }
    }
  } else {
    // Nessun match, crea nuova questione
    // (In realtà questo caso dovrebbe essere gestito dal clustering,
    // qui gestiamo solo email che non avevano cluster)
    result.nuovo = true;
  }
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// MENU E UI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Menu: Avvia batch questioni
 */
function menuAvviaQuestioniBatch() {
  var ui = SpreadsheetApp.getUi();
  
  var risposta = ui.alert(
    "📦 Avvia Batch Questioni",
    "Questa operazione:\n\n" +
    "1. Scansiona email con score problema/urgente alto\n" +
    "2. Raggruppa per fornitore\n" +
    "3. Invia richieste AI a OpenAI Batch (sconto 50%)\n" +
    "4. Attendi 1-24h per completamento\n\n" +
    "Continuare?",
    ui.ButtonSet.YES_NO
  );
  
  if (risposta !== ui.Button.YES) return;
  
  // Chiedi parametri
  var giorni = ui.prompt(
    "Giorni Indietro",
    "Quanti giorni di storico analizzare? (default: 90)",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (giorni.getSelectedButton() !== ui.Button.OK) return;
  
  var giorniVal = parseInt(giorni.getResponseText()) || 90;
  
  try {
    var report = avviaQuestioniBatch({
      giorniIndietro: giorniVal,
      maxEmail: 500,
      soloAccoda: false
    });
    
    var msg = "📦 BATCH AVVIATO\n\n" +
      "Email scansite: " + report.fase1_scan.candidati + "\n" +
      "Gruppi fornitore: " + report.fase2_gruppi.fornitori + "\n" +
      "Richieste clustering: " + report.fase3_clustering.richieste + "\n" +
      "Richieste matching: " + report.fase4_matching.richieste + "\n\n";
    
    if (report.fase5_invio.success) {
      msg += "✅ Batch inviato: " + report.fase5_invio.batchJobId + "\n\n" +
        "Attendi 1-24h, poi esegui:\n" +
        "Menu > Batch > Processa Risultati Questioni";
    } else {
      msg += "⚠️ Invio non completato\n" + report.errori.join("\n");
    }
    
    ui.alert("Batch Questioni", msg, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert("❌ Errore", e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu: Processa risultati batch questioni
 */
function menuProcessaRisultatiBatchQuestioni() {
  var ui = SpreadsheetApp.getUi();
  
  // Prima poll per aggiornare status
  var pollResult = pollAllBatchJobs();
  
  if (pollResult.completed.length === 0 && pollResult.inProgress.length > 0) {
    ui.alert(
      "⏳ Batch in Corso",
      "Ci sono " + pollResult.inProgress.length + " batch ancora in elaborazione.\n\n" +
      "Riprova più tardi.",
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Recupera risultati
  var recuperoResult = recuperaTuttiRisultatiCompletati();
  
  // Processa per questioni
  var processResult = processaRisultatiBatchQuestioni();
  
  var msg = "📦 RISULTATI PROCESSATI\n\n" +
    "Clustering:\n" +
    "- Processati: " + processResult.clustering.processati + "\n" +
    "- Questioni create: " + processResult.clustering.questioni + "\n\n" +
    "Matching:\n" +
    "- Processati: " + processResult.matching.processati + "\n" +
    "- Collegati: " + processResult.matching.collegati + "\n" +
    "- Nuovi: " + processResult.matching.nuovi;
  
  if (processResult.errori.length > 0) {
    msg += "\n\n⚠️ Errori: " + processResult.errori.length;
  }
  
  ui.alert("Risultati Batch", msg, ui.ButtonSet.OK);
}

/**
 * Menu: Stato batch
 */
function menuStatoBatch() {
  var ui = SpreadsheetApp.getUi();
  
  var stats = getStatsBatch();
  
  var msg = "📦 STATO BATCH\n\n" +
    "━━━ CODA RICHIESTE ━━━\n" +
    "In attesa: " + stats.queue.pending + "\n" +
    "In elaborazione: " + stats.queue.processing + "\n" +
    "Completate: " + stats.queue.completed + "\n" +
    "Errori: " + stats.queue.error + "\n" +
    "Totale: " + stats.queue.totale + "\n\n" +
    "━━━ JOB ━━━\n" +
    "Inviati: " + stats.jobs.submitted + "\n" +
    "In corso: " + stats.jobs.inProgress + "\n" +
    "Completati: " + stats.jobs.completed + "\n" +
    "Falliti: " + stats.jobs.failed + "\n" +
    "Totale: " + stats.jobs.totale;
  
  ui.alert("Stato Batch", msg, ui.ButtonSet.OK);
}

/**
 * Crea submenu batch
 */
function creaSubmenuBatch() {
  var ui = SpreadsheetApp.getUi();
  
  return ui.createMenu("📦 Batch API")
    .addItem("🚀 Avvia Batch Questioni", "menuAvviaQuestioniBatch")
    .addItem("📥 Processa Risultati", "menuProcessaRisultatiBatchQuestioni")
    .addSeparator()
    .addItem("🔄 Poll Status Job", "pollAllBatchJobs")
    .addItem("📊 Stato Batch", "menuStatoBatch")
    .addSeparator()
    .addItem("🛠️ Inizializza Fogli Batch", "initBatchSheets");
}