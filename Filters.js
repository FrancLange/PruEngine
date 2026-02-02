/**
 * ==========================================================================================
 * FILTERS.gs - Filtri Pre-Analisi v1.1.0
 * ==========================================================================================
 * Layer 0: Spam Filter
 * 
 * CHANGELOG v1.1.0:
 * - AGGIUNTA: eseguiLayer0SpamFilter() - funzione orchestratrice batch
 * - FIX: Integrazione completa nel flusso Logic.js
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LAYER 0 - ORCHESTRATORE BATCH (NUOVA FUNZIONE!)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Esegue Layer 0 Spam Filter su batch di email
 * 
 * Questa funzione è chiamata da Logic.js → analizzaEmailInCoda()
 * 
 * FLUSSO:
 * 1. Carica email con status DA_FILTRARE
 * 2. Per ogni email → chiama analisiLayer0SpamFilter()
 * 3. Scrive risultati L0 nel foglio
 * 4. Aggiorna status → SPAM o DA_ANALIZZARE
 * 
 * @param {Number} maxEmail - Max email da processare (default 50)
 * @returns {Object} {totali, spam, legit, errori}
 */
function eseguiLayer0SpamFilter(maxEmail) {
  maxEmail = maxEmail || 50;
  var startTime = new Date();
  
  var risultati = {
    totali: 0,
    spam: 0,
    legit: 0,
    errori: 0
  };
  
  logSistema("🛡️ L0 START: Spam Filter");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    logSistema("⚠️ L0: Nessuna email in LOG_IN");
    return risultati;
  }
  
  // Carica email con status DA_FILTRARE
  var emails = caricaEmailDaAnalizzare(sheet, "L0", maxEmail);
  
  if (emails.length === 0) {
    logSistema("✅ L0: Nessuna email da filtrare (tutte già processate)");
    return risultati;
  }
  
  risultati.totali = emails.length;
  logSistema("📧 L0: " + emails.length + " email in coda filtro spam");
  
  // Processa ogni email
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    
    try {
      // Prepara dati email
      var emailData = {
        mittente: email.mittente || "",
        oggetto: email.oggetto || "",
        corpo: email.corpo || ""
      };
      
      // Esegui analisi L0
      var l0Result;
      var connectorAttivo = false;
      
      // Check se Connector Fornitori è attivo
      try {
        connectorAttivo = typeof isConnectorFornitoriAttivo === 'function' && isConnectorFornitoriAttivo();
      } catch(e) {
        connectorAttivo = false;
      }
      
      if (connectorAttivo && typeof eseguiLayer0SpamFilter_PATCHED === 'function') {
        // Usa versione con integrazione Fornitori (skip L0 per fornitori noti)
        l0Result = eseguiLayer0SpamFilter_PATCHED(emailData);
      } else {
        // Usa versione standard
        l0Result = analisiLayer0SpamFilter(emailData);
      }
      
      // Scrivi risultati nel foglio
      scriviRisultatoL0(sheet, email.rowNum, l0Result);
      
      // Determina nuovo status
      var nuovoStatus;
      if (l0Result.isSpam) {
        nuovoStatus = CONFIG.STATUS_EMAIL.SPAM;
        risultati.spam++;
        logSistema("🚫 L0: " + (email.id || "ROW-" + email.rowNum) + " → SPAM (" + l0Result.confidence + "%) - " + (l0Result.reason || ""));
      } else {
        nuovoStatus = CONFIG.STATUS_EMAIL.DA_ANALIZZARE;
        risultati.legit++;
        
        // Log extra se skip per fornitore noto
        if (l0Result.skipped && l0Result.skipReason) {
          logSistema("✅ L0: " + (email.id || "ROW-" + email.rowNum) + " → SKIP (" + l0Result.skipReason + ")");
        } else {
          logSistema("✅ L0: " + (email.id || "ROW-" + email.rowNum) + " → LEGIT (" + l0Result.confidence + "%)");
        }
      }
      
      // Aggiorna status
      aggiornaStatus(sheet, email.rowNum, nuovoStatus);
      
    } catch (error) {
      risultati.errori++;
      logSistema("❌ L0 Errore su " + (email.id || "ROW-" + email.rowNum) + ": " + error.toString());
      
      // In caso di errore, marca come DA_ANALIZZARE (safe default - meglio analizzare che perdere)
      try {
        aggiornaStatus(sheet, email.rowNum, CONFIG.STATUS_EMAIL.DA_ANALIZZARE);
      } catch(e) {
        // Ignora errore secondario
      }
    }
  }
  
  var durata = Math.round((new Date() - startTime) / 1000);
  logSistema("🛡️ L0 completato in " + durata + "s | Spam: " + risultati.spam + " | Legit: " + risultati.legit + " | Errori: " + risultati.errori);
  
  return risultati;
}


// ═══════════════════════════════════════════════════════════════════════
// LAYER 0 - ANALISI SINGOLA EMAIL
// ═══════════════════════════════════════════════════════════════════════

/**
 * Layer 0 Spam Filter CON integrazione Fornitori
 * 
 * LOGICA:
 * 1. Se connettore attivo → pre-check fornitore
 * 2. Se fornitore noto con status valido → SKIP L0 (non è spam)
 * 3. Se non fornitore noto → esegui L0 normale
 * 
 * @param {Object} email - {mittente, oggetto, corpo}
 * @returns {Object} {isSpam, confidence, reason, skipped, skipReason, fornitoreInfo}
 */
function eseguiLayer0SpamFilter_PATCHED(email) {
  // Pre-check connettore fornitori
  var fornitoreCheck = null;
  
  if (typeof preCheckEmailFornitore === 'function') {
    try {
      fornitoreCheck = preCheckEmailFornitore(email.mittente);
      
      if (fornitoreCheck && fornitoreCheck.isFornitoreNoto && fornitoreCheck.skipL0) {
        // Fornitore conosciuto, skip L0
        logSistema("🔗 L0 SKIP: Fornitore noto → " + fornitoreCheck.fornitore.nome);
        
        return {
          success: true,
          isSpam: false,
          confidence: 100,
          reason: null,
          skipped: true,
          skipReason: "Fornitore conosciuto: " + fornitoreCheck.fornitore.nome,
          fornitoreInfo: fornitoreCheck,
          modello: "SKIP_FORNITORE",
          timestamp: new Date()
        };
      }
    } catch(e) {
      // Se errore nel check fornitori, continua con L0 normale
      logSistema("⚠️ Errore preCheckEmailFornitore: " + e.toString());
    }
  }
  
  // Esegui L0 normale
  var risultato = analisiLayer0SpamFilter({
    mittente: email.mittente,
    oggetto: email.oggetto,
    corpo: email.corpo
  });
  
  // Aggiungi info fornitore se trovato (ma non skip)
  if (fornitoreCheck && fornitoreCheck.isFornitoreNoto) {
    risultato.fornitoreInfo = fornitoreCheck;
  }
  
  return risultato;
}


/**
 * Analisi Layer 0: Classifica email come SPAM o LEGIT
 * 
 * Usa GPT-4o-mini per velocità e costo ridotto
 * 
 * @param {Object} emailData - {mittente, oggetto, corpo}
 * @returns {Object} {success, isSpam, confidence, reason, modello, timestamp}
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
    "- Email da fornitori commerciali\n" +
    "- Conferme ordini, fatture, listini prezzi\n" +
    "- Comunicazioni su prodotti/servizi\n" +
    "- Risposte a campagne marketing nostre\n" +
    "- Problemi/reclami clienti\n" +
    "- Qualsiasi email con contenuto commerciale rilevante\n\n" +
    "IMPORTANTE: In caso di dubbio, classifica come LEGIT.\n" +
    "Meglio analizzare un'email in piu' che perderne una importante.\n\n" +
    "Rispondi SOLO in formato JSON:\n" +
    '{"isSpam": true/false, "confidence": 0-100, "reason": "Motivo breve se spam, altrimenti null"}';

  var corpoPreview = (emailData.corpo || "").substring(0, 500);
  
  var userPrompt = "Classifica questa email:\n\n" +
    "MITTENTE: " + (emailData.mittente || "") + "\n" +
    "OGGETTO: " + (emailData.oggetto || "") + "\n" +
    "CORPO (prime 500 caratteri):\n" +
    corpoPreview + "\n\n" +
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
      
      // Safety: Se confidence bassa su SPAM, considera LEGIT
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
      logSistema("⚠️ L0 Parsing error: " + e.toString());
      // In caso di errore parsing, assume LEGIT (safe default)
      return { 
        success: false, 
        isSpam: false,
        confidence: 0,
        reason: null,
        error: "Parsing fallito: " + e.toString(),
        modello: "ERROR",
        timestamp: new Date()
      };
    }
  }
  
  // Se chiamata API fallisce, assume LEGIT
  return { 
    success: false, 
    isSpam: false,
    confidence: 0,
    reason: null,
    error: result.error,
    modello: "ERROR",
    timestamp: new Date()
  };
}


// ═══════════════════════════════════════════════════════════════════════
// FILTRI FUTURI (Placeholder)
// ═══════════════════════════════════════════════════════════════════════

/**
 * [FUTURO] Filtro rapido per identificare fatture
 * Da implementare quando necessario
 */
function filtroFattura(emailData) {
  // TODO: Implementare
  // Potrebbe controllare:
  // - Parole chiave nell'oggetto (fattura, invoice, pagamento)
  // - Allegati PDF
  // - Mittenti noti per fatture
  return { isFattura: false, confidence: 0 };
}

/**
 * [FUTURO] Filtro rapido per email urgenti
 * Da implementare quando necessario
 */
function filtroUrgente(emailData) {
  // TODO: Implementare
  // Potrebbe controllare:
  // - Parole chiave (urgente, urgent, asap)
  // - Mittenti VIP
  // - Contesto situazionale
  return { isUrgente: false, confidence: 0 };
}