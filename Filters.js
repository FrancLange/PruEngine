/**
 * ==========================================================================================
 * FILTERS.gs - Filtri Pre-Analisi v1.0.0
 * ==========================================================================================
 * Layer 0: Spam Filter
 * (Futuro: Filtro Fatture, Filtro Urgenti, etc.)
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// LAYER 0 - SPAM FILTER
// ═══════════════════════════════════════════════════════════════════════

/**
 * Layer 0 Spam Filter CON integrazione Fornitori
 * 
 * LOGICA:
 * 1. Se connettore attivo → pre-check fornitore
 * 2. Se fornitore noto con status valido → SKIP L0 (non è spam)
 * 3. Se non fornitore noto → esegui L0 normale
 */
function eseguiLayer0SpamFilter_PATCHED(email) {
  // 🆕 PATCH: Pre-check connettore fornitori
  var fornitoreCheck = null;
  
  if (typeof preCheckEmailFornitore === 'function') {
    fornitoreCheck = preCheckEmailFornitore(email.mittente);
    
    if (fornitoreCheck.isFornitoreNoto && fornitoreCheck.skipL0) {
      // Fornitore conosciuto, skip L0
      logSistema("🔗 L0 SKIP: Fornitore noto → " + fornitoreCheck.fornitore.nome);
      
      return {
        isSpam: false,
        confidence: 100,
        reason: null,
        skipped: true,
        skipReason: "Fornitore conosciuto: " + fornitoreCheck.fornitore.nome,
        fornitoreInfo: fornitoreCheck  // 🆕 Passa info al layer successivo
      };
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