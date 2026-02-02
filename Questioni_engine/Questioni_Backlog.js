/**
 * ==========================================================================================
 * QUESTIONI_BACKLOG.js - Primo Popolamento v1.0.0
 * ==========================================================================================
 * Scansione email storiche e creazione questioni da backlog
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// POPOLAMENTO BACKLOG
// ═══════════════════════════════════════════════════════════════════════

/**
 * MAIN: Popola questioni da email backlog
 * 
 * @param {Object} params - {giorni: 90, max_email: 500}
 * @returns {Object} Report
 */
function popolaQuestioniDaBacklog(params) {
  params = params || {};
  var giorni = params.giorni || QUESTIONI_CONFIG.BACKLOG_DEFAULT_GIORNI;
  var maxEmail = params.max_email || QUESTIONI_CONFIG.BACKLOG_MAX_EMAIL;
  
  var startTime = new Date();
  
  logQuestioni("════════════════════════════════════════");
  logQuestioni("📥 BACKLOG START - Ultimi " + giorni + " giorni");
  logQuestioni("════════════════════════════════════════");
  
  var report = {
    email_scansionate: 0,
    email_con_problema: 0,
    fornitori_trovati: 0,
    questioni_create: 0,
    email_collegate: 0,
    dettaglio: []
  };
  
  try {
    // FASE 1: Scan LOG_IN
    logQuestioni("📧 FASE 1: Scan email...");
    var emailProblema = scanEmailPerQuestioni({
      giorni: giorni,
      max: maxEmail
    });
    
    report.email_scansionate = emailProblema.totale_scansionate;
    report.email_con_problema = emailProblema.emails.length;
    
    logQuestioni("Trovate " + emailProblema.emails.length + " email con problemi");
    
    if (emailProblema.emails.length === 0) {
      logQuestioni("✅ Nessuna email da processare");
      return report;
    }
    
    // FASE 2: Raggruppa per fornitore
    logQuestioni("👥 FASE 2: Raggruppamento per fornitore...");
    var gruppi = raggruppaEmailPerFornitore(emailProblema.emails);
    
    report.fornitori_trovati = Object.keys(gruppi).length;
    logQuestioni("Trovati " + report.fornitori_trovati + " fornitori/domini");
    
    // FASE 3-4: Processa ogni gruppo
    logQuestioni("🎫 FASE 3-4: Clustering e creazione questioni...");
    
    Object.keys(gruppi).forEach(function(chiave) {
      var gruppo = gruppi[chiave];
      
      logQuestioni("Processo: " + (gruppo.nome || chiave) + " (" + gruppo.emails.length + " email)");
      
      var risultatoGruppo = processaGruppoFornitore(gruppo);
      
      report.questioni_create += risultatoGruppo.questioni_create;
      report.email_collegate += risultatoGruppo.email_collegate;
      
      report.dettaglio.push({
        fornitore: gruppo.nome || chiave,
        email: gruppo.emails.length,
        questioni: risultatoGruppo.questioni_create
      });
    });
    
    var durata = Math.round((new Date() - startTime) / 1000);
    
    logQuestioni("════════════════════════════════════════");
    logQuestioni("✅ BACKLOG COMPLETATO in " + durata + "s");
    logQuestioni("Questioni create: " + report.questioni_create);
    logQuestioni("Email collegate: " + report.email_collegate);
    logQuestioni("════════════════════════════════════════");
    
  } catch(e) {
    logQuestioni("❌ Errore backlog: " + e.toString());
    throw e;
  }
  
  return report;
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 1: SCAN EMAIL
// ═══════════════════════════════════════════════════════════════════════

/**
 * Scansiona LOG_IN per email con problemi
 */
function scanEmailPerQuestioni(params) {
  var giorni = params.giorni || 90;
  var maxEmail = params.max || 500;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { emails: [], totale_scansionate: 0 };
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
  
  var dataLimite = new Date();
  dataLimite.setDate(dataLimite.getDate() - giorni);
  
  var emails = [];
  var totaleScansionate = 0;
  
  // Ordina dal più vecchio (reverse)
  for (var i = rows.length - 1; i >= 0 && emails.length < maxEmail; i--) {
    var row = rows[i];
    var rowNum = i + 2;
    
    // Skip vuoti
    if (!row[colMap.ID_EMAIL]) continue;
    
    totaleScansionate++;
    
    // Criteri esclusione
    var status = row[colMap.STATUS];
    if (status === CONFIG.STATUS_EMAIL.SPAM) continue;
    
    // Verifica L3 completato
    if (!row[colMap.L3_TIMESTAMP]) continue;
    
    // Verifica data
    var timestamp = row[colMap.TIMESTAMP];
    if (timestamp && new Date(timestamp) < dataLimite) continue;
    
    // TODO: Check se già ha questione collegata (colonna futura)
    // Per ora skip se ha già L3_AZIONI_SUGGERITE processate
    
    // Parse scores
    var scores = {};
    try {
      var scoresJson = row[colMap.L3_SCORES_JSON];
      if (scoresJson) scores = JSON.parse(scoresJson);
    } catch(e) {}
    
    // Verifica soglie
    var problema = parseInt(scores.problema) || 0;
    var urgente = parseInt(scores.urgente) || 0;
    
    if (problema <= QUESTIONI_CONFIG.SOGLIA_PROBLEMA && 
        urgente <= QUESTIONI_CONFIG.SOGLIA_URGENTE) {
      continue;
    }
    
    // Parse tags
    var tags = [];
    var tagsStr = row[colMap.L3_TAGS_FINALI] || row[colMap.L1_TAGS] || "";
    if (tagsStr) tags = tagsStr.split(",").map(function(t) { return t.trim(); });
    
    emails.push({
      rowNum: rowNum,
      id: row[colMap.ID_EMAIL],
      mittente: row[colMap.MITTENTE] || "",
      oggetto: row[colMap.OGGETTO] || "",
      sintesi: row[colMap.L3_SINTESI_FINALE] || row[colMap.L1_SINTESI] || "",
      tags: tags,
      scores: scores,
      data: timestamp ? new Date(timestamp).toLocaleDateString('it-IT') : ""
    });
  }
  
  // Inverti per avere ordine cronologico
  emails.reverse();
  
  return {
    emails: emails,
    totale_scansionate: totaleScansionate
  };
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 2: RAGGRUPPA PER FORNITORE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Raggruppa email per fornitore (o dominio se sconosciuto)
 */
function raggruppaEmailPerFornitore(emails) {
  var gruppi = {};
  
  emails.forEach(function(email) {
    var fornitore = null;
    var chiave = "";
    
    // Lookup fornitore
    if (typeof lookupFornitoreByEmail === 'function') {
      fornitore = lookupFornitoreByEmail(email.mittente);
    }
    
    if (fornitore) {
      chiave = fornitore.id;
      
      if (!gruppi[chiave]) {
        gruppi[chiave] = {
          tipo: "FORNITORE",
          id: fornitore.id,
          nome: fornitore.nome,
          dominio: "",
          emails: []
        };
      }
    } else {
      // Fallback su dominio
      var dominio = estraiDominioEmail(email.mittente);
      chiave = "UNKNOWN:" + dominio;
      
      if (!gruppi[chiave]) {
        gruppi[chiave] = {
          tipo: "DOMINIO",
          id: "",
          nome: "",
          dominio: dominio,
          emails: []
        };
      }
    }
    
    gruppi[chiave].emails.push(email);
  });
  
  return gruppi;
}

// ═══════════════════════════════════════════════════════════════════════
// FASE 3-4: PROCESSA GRUPPO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa un gruppo di email per un fornitore
 */
function processaGruppoFornitore(gruppo) {
  var risultato = {
    questioni_create: 0,
    email_collegate: 0
  };
  
  if (!gruppo.emails || gruppo.emails.length === 0) {
    return risultato;
  }
  
  // Se troppe email, processa a batch
  var emailDaProcessare = gruppo.emails;
  if (emailDaProcessare.length > QUESTIONI_CONFIG.MAX_EMAIL_PER_CLUSTER_CALL) {
    emailDaProcessare = gruppo.emails.slice(0, QUESTIONI_CONFIG.MAX_EMAIL_PER_CLUSTER_CALL);
    logQuestioni("⚠️ Limitato a " + QUESTIONI_CONFIG.MAX_EMAIL_PER_CLUSTER_CALL + " email");
  }
  
  // AI Clustering
  var clusterResult = aiClusterEmailFornitore(emailDaProcessare, {
    id: gruppo.id,
    nome: gruppo.nome || gruppo.dominio
  });
  
  // Crea questioni per ogni cluster
  clusterResult.clusters.forEach(function(cluster) {
    if (!cluster.email_ids || cluster.email_ids.length === 0) return;
    
    // Trova email del cluster per calcolare max scores
    var emailCluster = emailDaProcessare.filter(function(e) {
      return cluster.email_ids.indexOf(e.id) >= 0;
    });
    
    var maxProblema = 0;
    var maxUrgenza = 0;
    var timestampPrimaEmail = null;
    
    emailCluster.forEach(function(e) {
      var p = e.scores ? (parseInt(e.scores.problema) || 0) : 0;
      var u = e.scores ? (parseInt(e.scores.urgente) || 0) : 0;
      if (p > maxProblema) maxProblema = p;
      if (u > maxUrgenza) maxUrgenza = u;
    });
    
    // Crea questione
    try {
      var idQuestione = creaQuestione({
        origine: QUESTIONI_CONFIG.ORIGINI.AUTO_BACKLOG,
        id_fornitore: gruppo.id,
        nome_fornitore: gruppo.nome,
        dominio_email: gruppo.dominio,
        email_collegate: cluster.email_ids,
        titolo: cluster.titolo,
        descrizione: cluster.descrizione,
        categoria: cluster.categoria,
        score_problema: maxProblema,
        score_urgenza: maxUrgenza
      });
      
      risultato.questioni_create++;
      risultato.email_collegate += cluster.email_ids.length;
      
      // TODO: Aggiorna LOG_IN con ID_QUESTIONE_COLLEGATA
      
    } catch(e) {
      logQuestioni("❌ Errore creazione questione: " + e.message);
    }
  });
  
  return risultato;
}

// ═══════════════════════════════════════════════════════════════════════
// PROCESSAMENTO REAL-TIME
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa email recenti per questioni (hook post-analisi)
 */
function processaEmailRecentiPerQuestioni() {
  var risultato = {
    processate: 0,
    questioni_nuove: 0,
    collegate: 0,
    con_flag: 0
  };
  
  // Scan email di oggi
  var emails = scanEmailPerQuestioni({
    giorni: 1,
    max: 50
  });
  
  if (emails.emails.length === 0) {
    return risultato;
  }
  
  emails.emails.forEach(function(email) {
    risultato.processate++;
    
    var esito = processaSingolaEmailPerQuestione(email);
    
    if (esito.azione === "CREATA") {
      risultato.questioni_nuove++;
      if (esito.flag) risultato.con_flag++;
    } else if (esito.azione === "COLLEGATA") {
      risultato.collegate++;
    }
  });
  
  return risultato;
}

/**
 * Processa singola email per questione (real-time)
 */
function processaSingolaEmailPerQuestione(email) {
  var esito = {
    azione: "NESSUNA",
    id_questione: null,
    flag: false
  };
  
  // Lookup fornitore
  var fornitore = null;
  if (typeof lookupFornitoreByEmail === 'function') {
    fornitore = lookupFornitoreByEmail(email.mittente);
  }
  
  var idFornitore = fornitore ? fornitore.id : null;
  var nomeFornitore = fornitore ? fornitore.nome : "";
  var dominio = estraiDominioEmail(email.mittente);
  
  // Cerca questioni aperte
  var questioniAperte = [];
  if (idFornitore) {
    questioniAperte = getQuestioniApertePerFornitore(idFornitore);
  } else if (dominio) {
    questioniAperte = getQuestioniApertePerDominio(dominio);
  }
  
  if (questioniAperte.length === 0) {
    // Nessuna questione aperta → crea nuova
    var id = creaQuestione({
      origine: QUESTIONI_CONFIG.ORIGINI.AUTO_EMAIL,
      id_fornitore: idFornitore,
      nome_fornitore: nomeFornitore,
      dominio_email: dominio,
      email_collegate: [email.id],
      titolo: (email.oggetto || "").substring(0, 60),
      descrizione: email.sintesi || "",
      categoria: determinaCategoriaQuestione(email.tags, email.scores),
      score_problema: email.scores ? email.scores.problema : 0,
      score_urgenza: email.scores ? email.scores.urgente : 0,
      timestamp_email: new Date()
    });
    
    esito.azione = "CREATA";
    esito.id_questione = id;
    return esito;
  }
  
  // AI Matching
  var matchResult = aiMatchQuestioneEsistente(email, questioniAperte);
  
  if (matchResult.match && matchResult.confidence >= QUESTIONI_CONFIG.CONFIDENCE_AUTO_MATCH) {
    // Match confermato → collega
    collegaEmailAQuestione(email.id, matchResult.id_questione, email);
    
    esito.azione = "COLLEGATA";
    esito.id_questione = matchResult.id_questione;
    
  } else if (matchResult.confidence >= QUESTIONI_CONFIG.CONFIDENCE_FLAG_REVIEW) {
    // Match incerto → crea con flag
    var id = creaQuestione({
      origine: QUESTIONI_CONFIG.ORIGINI.AUTO_EMAIL,
      id_fornitore: idFornitore,
      nome_fornitore: nomeFornitore,
      dominio_email: dominio,
      email_collegate: [email.id],
      titolo: (email.oggetto || "").substring(0, 60),
      descrizione: email.sintesi || "",
      categoria: determinaCategoriaQuestione(email.tags, email.scores),
      score_problema: email.scores ? email.scores.problema : 0,
      score_urgenza: email.scores ? email.scores.urgente : 0,
      flag_review: QUESTIONI_CONFIG.FLAGS.POSSIBILE_DUPLICATO
    });
    
    esito.azione = "CREATA";
    esito.id_questione = id;
    esito.flag = true;
    
  } else {
    // No match → crea nuova
    var id = creaQuestione({
      origine: QUESTIONI_CONFIG.ORIGINI.AUTO_EMAIL,
      id_fornitore: idFornitore,
      nome_fornitore: nomeFornitore,
      dominio_email: dominio,
      email_collegate: [email.id],
      titolo: (email.oggetto || "").substring(0, 60),
      descrizione: email.sintesi || "",
      categoria: determinaCategoriaQuestione(email.tags, email.scores),
      score_problema: email.scores ? email.scores.problema : 0,
      score_urgenza: email.scores ? email.scores.urgente : 0
    });
    
    esito.azione = "CREATA";
    esito.id_questione = id;
  }
  
  return esito;
}