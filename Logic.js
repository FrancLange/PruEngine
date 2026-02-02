/**
 * ==========================================================================================
 * LOGIC.gs - Orchestrazione Analisi Email v1.2.0 (Smart Mode)
 * ==========================================================================================
 * Flusso SMART:
 * - Se Claude configurato: L0 → L1 → L2 → L3 (full)
 * - Se Claude NON configurato: L0 → L1 → STOP (fast)
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// FLUSSO PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Analizza email in coda - FLUSSO SMART
 * @param {Number} maxEmail - Max email da processare (default 50)
 * @param {Boolean} forceReanalisi - Forza rianalisi
 * @returns {Object} Risultati per ogni layer
 */
function analizzaEmailInCoda(maxEmail, forceReanalisi) {
  maxEmail = maxEmail || 50;
  forceReanalisi = forceReanalisi || false;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet) {
    logSistema("❌ Errore: Foglio LOG_IN non trovato");
    return { errore: "LOG_IN non trovato" };
  }
  
  var startTime = new Date();
  
  // Check se Claude è configurato
  var claudeKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  var modalita = claudeKey ? "FULL (L0→L1→L2→L3)" : "FAST (L0→L1)";
  
  logSistema("═══════════════════════════════════════════");
  logSistema("🚀 ANALISI EMAIL - START");
  logSistema("📋 Modalità: " + modalita);
  logSistema("═══════════════════════════════════════════");
  
  var risultati = {
    totali: 0,
    modalita: claudeKey ? "FULL" : "FAST",
    l0: { spam: 0, legit: 0, errori: 0 },
    l1: { ok: 0, errori: 0 },
    l2: { ok: 0, errori: 0, skipped: false },
    l3: { ok: 0, needsReview: 0, skipped: false }
  };
  
  // ═══════════════════════════════════════════════════════════
  // FASE 1: LAYER 0 - SPAM FILTER
  // ═══════════════════════════════════════════════════════════
  
  logSistema("🛡️ FASE 1: Layer 0 Spam Filter");
  var l0Result = eseguiLayer0SpamFilter(maxEmail);
  risultati.l0 = l0Result;
  risultati.totali = l0Result.totali;
  
  // ═══════════════════════════════════════════════════════════
  // FASE 2: LAYER 1 - GPT
  // ═══════════════════════════════════════════════════════════
  
  logSistema("🔵 FASE 2: Layer 1 (GPT)");
  try {
    var l1Result = analizzaEmailL1({ max_email: maxEmail });
    risultati.l1.ok = l1Result.ok || 0;
    risultati.l1.errori = l1Result.errori || 0;
    logSistema("🔵 L1 completato: " + risultati.l1.ok + " ok, " + risultati.l1.errori + " errori");
  } catch (e) {
    logSistema("❌ L1 Errore critico: " + e.toString());
  }
  
  // ═══════════════════════════════════════════════════════════
  // DECISIONE: FULL o FAST MODE?
  // ═══════════════════════════════════════════════════════════
  
  if (!claudeKey) {
    // ═══════════════════════════════════════════════════════════
    // FAST MODE: Copia L1 → Output finale, skip L2/L3
    // ═══════════════════════════════════════════════════════════
    
    logSistema("⚡ FAST MODE: Skip L2/L3, finalizza da L1");
    
    try {
      var finalizzate = finalizzaDaL1(sheet, maxEmail);
      risultati.l2.skipped = true;
      risultati.l3.skipped = true;
      risultati.l3.ok = finalizzate;
      logSistema("⚡ Finalizzate " + finalizzate + " email da L1");
    } catch (e) {
      logSistema("❌ Errore finalizzazione: " + e.toString());
    }
    
  } else {
    // ═══════════════════════════════════════════════════════════
    // FULL MODE: L2 (Claude) + L3 (Merge)
    // ═══════════════════════════════════════════════════════════
    
    logSistema("🟣 FASE 3: Layer 2 (Claude)");
    try {
      var l2Result = analizzaEmailL2({ max_email: maxEmail });
      risultati.l2.ok = l2Result.ok || 0;
      risultati.l2.errori = l2Result.errori || 0;
      logSistema("🟣 L2 completato: " + risultati.l2.ok + " ok, " + risultati.l2.errori + " errori");
    } catch (e) {
      logSistema("❌ L2 Errore critico: " + e.toString());
    }
    
    logSistema("🟢 FASE 4: Layer 3 (Merge)");
    try {
      var l3Result = mergeAnalisiL3({});
      risultati.l3.ok = l3Result.ok || 0;
      risultati.l3.needsReview = l3Result.needsReview || 0;
      logSistema("🟢 L3 completato: " + risultati.l3.ok + " ok, " + risultati.l3.needsReview + " needs review");
    } catch (e) {
      logSistema("❌ L3 Errore critico: " + e.toString());
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // REPORT FINALE
  // ═══════════════════════════════════════════════════════════
  
  var durata = Math.round((new Date() - startTime) / 1000);
  
  logSistema("═══════════════════════════════════════════");
  logSistema("✅ ANALISI COMPLETATA in " + durata + "s (" + risultati.modalita + ")");
  logSistema("📊 L0: " + risultati.l0.spam + " spam, " + risultati.l0.legit + " legit");
  logSistema("📊 L1: " + risultati.l1.ok + " ok");
  if (risultati.l2.skipped) {
    logSistema("📊 L2: SKIPPED (no Claude)");
    logSistema("📊 L3: SKIPPED → " + risultati.l3.ok + " finalizzate da L1");
  } else {
    logSistema("📊 L2: " + risultati.l2.ok + " ok");
    logSistema("📊 L3: " + risultati.l3.ok + " ok, " + risultati.l3.needsReview + " review");
  }
  logSistema("═══════════════════════════════════════════");
  
  return risultati;
}

// ═══════════════════════════════════════════════════════════════════════
// FAST MODE: FINALIZZA DA L1
// ═══════════════════════════════════════════════════════════════════════

/**
 * Copia risultati L1 come output finale (skip L2/L3)
 * @param {Sheet} sheet - Foglio LOG_IN
 * @param {Number} maxEmail - Max email
 * @returns {Number} Email finalizzate
 */
function finalizzaDaL1(sheet, maxEmail) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Mappa colonne
  var colMap = {};
  Object.keys(CONFIG.COLONNE_LOG_IN).forEach(function(key) {
    var nome = CONFIG.COLONNE_LOG_IN[key];
    var idx = headers.indexOf(nome);
    if (idx >= 0) colMap[key] = idx;
  });
  
  var count = 0;
  
  for (var i = 0; i < rows.length && count < maxEmail; i++) {
    var row = rows[i];
    var rowNum = i + 2;
    
    // Trova email con L1 completato ma non ancora finalizzate
    var l1Timestamp = row[colMap.L1_TIMESTAMP];
    var l3Timestamp = row[colMap.L3_TIMESTAMP];
    var status = row[colMap.STATUS];
    
    // Se ha L1 ma non L3, e non è già ANALIZZATO
    if (l1Timestamp && !l3Timestamp && status !== "ANALIZZATO" && status !== "SPAM") {
      
      // Copia L1 → L3
      var l1Tags = row[colMap.L1_TAGS] || "";
      var l1Sintesi = row[colMap.L1_SINTESI] || "";
      var l1ScoresJson = row[colMap.L1_SCORES_JSON] || "{}";
      
      // Scrivi risultati finali
      var risultatoFinale = {
        tags: l1Tags ? l1Tags.split(", ") : [],
        sintesi: l1Sintesi,
        scores: {},
        confidence: 95, // Alta confidence per GPT-4o
        divergenza: 0,
        needsReview: false,
        azioniSuggerite: []
      };
      
      // Parse scores e genera azioni
      try {
        risultatoFinale.scores = JSON.parse(l1ScoresJson);
        risultatoFinale.azioniSuggerite = generaAzioniDaScores(risultatoFinale.scores);
      } catch (e) {}
      
      // Scrivi L3
      scriviRisultatiL3(sheet, rowNum, risultatoFinale);
      
      // Aggiorna status
      aggiornaStatus(sheet, rowNum, "ANALIZZATO");
      
      count++;
      logSistema("⚡ Finalizzata: " + (row[colMap.ID_EMAIL] || "ROW-" + rowNum));
    }
  }
  
  return count;
}

/**
 * Genera azioni suggerite dagli scores
 * @param {Object} scores - Scores L1
 * @returns {Array} Azioni suggerite
 */
function generaAzioniDaScores(scores) {
  var azioni = [];
  
  if (scores.problema > 70) azioni.push("CREA_QUESTIONE_PROBLEMA");
  if (scores.fattura > 70) azioni.push("FLAG_FATTURA");
  if (scores.novita > 70) azioni.push("FLAG_NOVITA");
  if (scores.catalogo > 70) azioni.push("FLAG_CATALOGO");
  if (scores.prezzi > 70) azioni.push("FLAG_PREZZI");
  if (scores.risposta_campagna > 70) azioni.push("AGGIORNA_CAMPAGNA");
  if (scores.urgente > 80) azioni.push("ALERT_URGENTE");
  if (scores.ordine > 70) azioni.push("FLAG_ORDINE");
  
  return azioni;
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 1 - GPT CATEGORIZATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Analizza email con GPT-4o (Layer 1)
 * @param {Object} params - {max_email: 10}
 * @returns {Object} {analizzate, ok, errori}
 */
function analizzaEmailL1(params) {
  var MAX_EMAIL = (params && params.max_email) ? params.max_email : 10;
  var startTime = new Date();
  
  logSistema("🔵 L1 START: Analisi GPT-4o");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    logSistema("⚠️ L1: Nessuna email in LOG_IN");
    return { analizzate: 0, ok: 0, errori: 0 };
  }
  
  var emails = caricaEmailDaAnalizzare(sheet, "L1", MAX_EMAIL);
  
  if (emails.length === 0) {
    logSistema("✅ L1: Nessuna email da analizzare");
    return { analizzate: 0, ok: 0, errori: 0 };
  }
  
  logSistema("📧 L1: " + emails.length + " email in coda");
  
  var countOk = 0;
  var countErrore = 0;
  
  emails.forEach(function(email) {
    try {
      aggiornaStatus(sheet, email.rowNum, "IN_ANALISI");
      
      var emailData = {
        mittente: email.mittente,
        oggetto: email.oggetto,
        corpo: email.corpo
      };
      
      var risultato = analisiLayer1(emailData);
      
      if (!risultato.success) {
        throw new Error(risultato.error || "Analisi Layer 1 fallita");
      }
      
      scriviRisultatiL1(sheet, email.rowNum, risultato);
      countOk++;
      logSistema("✅ L1: " + email.id + " - " + (risultato.tags || []).join(", "));
      
    } catch (error) {
      countErrore++;
      scriviErroreL1(sheet, email.rowNum, error.toString());
      logSistema("❌ L1: " + email.id + " - " + error.message);
    }
  });
  
  var durata = Math.round((new Date() - startTime) / 1000);
  logSistema("🔵 L1 completato: " + countOk + " OK, " + countErrore + " errori in " + durata + "s");
  
  return { analizzate: emails.length, ok: countOk, errori: countErrore };
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 2 - CLAUDE VERIFICATION (Solo se configurato)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica email con Claude (Layer 2)
 * @param {Object} params - {max_email: 10}
 * @returns {Object} {verificate, ok, errori}
 */
function analizzaEmailL2(params) {
  var MAX_EMAIL = (params && params.max_email) ? params.max_email : 10;
  var startTime = new Date();
  
  logSistema("🟣 L2 START: Verifica Claude");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    logSistema("⚠️ L2: Nessuna email in LOG_IN");
    return { verificate: 0, ok: 0, errori: 0 };
  }
  
  var emails = caricaEmailDaAnalizzare(sheet, "L2", MAX_EMAIL);
  
  if (emails.length === 0) {
    logSistema("✅ L2: Nessuna email da verificare");
    return { verificate: 0, ok: 0, errori: 0 };
  }
  
  logSistema("📧 L2: " + emails.length + " email in coda");
  
  var countOk = 0;
  var countErrore = 0;
  
  emails.forEach(function(email) {
    try {
      var emailData = {
        mittente: email.mittente,
        oggetto: email.oggetto,
        corpo: email.corpo
      };
      
      var layer1Result = {
        tags: email.l1Tags ? email.l1Tags.split(", ") : [],
        sintesi: email.l1Sintesi || "",
        scores: {}
      };
      
      if (email.l1ScoresJson) {
        try { layer1Result.scores = JSON.parse(email.l1ScoresJson); } catch (e) {}
      }
      
      var risultato = analisiLayer2(emailData, layer1Result);
      
      if (!risultato.success) {
        throw new Error(risultato.error || "Analisi Layer 2 fallita");
      }
      
      scriviRisultatiL2(sheet, email.rowNum, risultato);
      countOk++;
      logSistema("✅ L2: " + email.id + " - Confidence: " + risultato.confidence + "%");
      
    } catch (error) {
      countErrore++;
      scriviErroreL2(sheet, email.rowNum, error.toString());
      logSistema("❌ L2: " + email.id + " - " + error.message);
    }
  });
  
  var durata = Math.round((new Date() - startTime) / 1000);
  logSistema("🟣 L2 completato: " + countOk + " OK, " + countErrore + " errori in " + durata + "s");
  
  return { verificate: emails.length, ok: countOk, errori: countErrore };
}

// ═══════════════════════════════════════════════════════════════════════
// LAYER 3 - MERGE (Solo se Claude configurato)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Merge risultati L1+L2 e calcola confidence finale
 * @param {Object} params - {}
 * @returns {Object} {merged, needsReview, ok}
 */
function mergeAnalisiL3(params) {
  var startTime = new Date();
  
  logSistema("🟢 L3 START: Merge Analisi");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    logSistema("⚠️ L3: Nessuna email in LOG_IN");
    return { merged: 0, needsReview: 0, ok: 0 };
  }
  
  var emails = caricaEmailDaAnalizzare(sheet, "L3", 100);
  
  if (emails.length === 0) {
    logSistema("✅ L3: Nessuna email da mergare");
    return { merged: 0, needsReview: 0, ok: 0 };
  }
  
  logSistema("📧 L3: " + emails.length + " email in coda");
  
  var countOk = 0;
  var countNeedsReview = 0;
  
  emails.forEach(function(email) {
    try {
      var layer1 = {
        success: true,
        tags: email.l1Tags ? email.l1Tags.split(", ") : [],
        sintesi: email.l1Sintesi || "",
        scores: {}
      };
      
      if (email.l1ScoresJson) {
        try { layer1.scores = JSON.parse(email.l1ScoresJson); } catch (e) {}
      }
      
      var layer2 = {
        success: true,
        tags: email.l2Tags ? email.l2Tags.split(", ") : [],
        sintesi: email.l2Sintesi || "",
        scores: {},
        confidence: email.l2Confidence || 50,
        richiestaRetry: email.l2RichiestaRetry === "SI"
      };
      
      if (email.l2ScoresJson) {
        try { layer2.scores = JSON.parse(email.l2ScoresJson); } catch (e) {}
      }
      
      var risultato = mergeAnalisi(layer1, layer2);
      
      if (!risultato.success) {
        throw new Error("Merge fallito");
      }
      
      scriviRisultatiL3(sheet, email.rowNum, risultato);
      
      var statusFinale = risultato.needsReview ? "NEEDS_REVIEW" : "ANALIZZATO";
      aggiornaStatus(sheet, email.rowNum, statusFinale);
      
      if (risultato.needsReview) {
        countNeedsReview++;
        logSistema("⚠️ L3: " + email.id + " - NEEDS_REVIEW (conf: " + risultato.confidence + "%)");
      } else {
        countOk++;
        logSistema("✅ L3: " + email.id + " - OK (conf: " + risultato.confidence + "%)");
      }
      
    } catch (error) {
      logSistema("❌ L3: " + email.id + " - " + error.message);
    }
  });
  
  var durata = Math.round((new Date() - startTime) / 1000);
  logSistema("🟢 L3 completato: " + countOk + " OK, " + countNeedsReview + " NEEDS_REVIEW in " + durata + "s");
  
  return { merged: emails.length, needsReview: countNeedsReview, ok: countOk };
}

/**
 * Helper per arricchire emailData con contesto fornitore
 */
function arricchisciEmailDataConFornitore(emailData, fornitoreInfo) {
  if (!fornitoreInfo || !fornitoreInfo.isFornitoreNoto) {
    return emailData;
  }
  
  // Aggiungi contesto
  emailData.contestoFornitore = fornitoreInfo.contesto;
  emailData.fornitore = fornitoreInfo.fornitore;
  
  return emailData;
}