/**
 * ==========================================================================================
 * QUESTIONI_CORE.js - Logica Principale v1.0.0 STANDALONE
 * ==========================================================================================
 * Creazione, elaborazione e gestione questioni
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CREAZIONE QUESTIONE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea una nuova questione
 * @param {Object} dati - Dati questione
 * @returns {Object} Questione creata
 */
function creaQuestione(dati) {
  if (!dati.titolo) {
    throw new Error("Titolo obbligatorio");
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUESTIONI_CONFIG.SHEETS.QUESTIONI);
  
  if (!sheet) {
    throw new Error("Foglio QUESTIONI non trovato. Esegui Setup.");
  }
  
  // Genera ID
  var idQuestione = generaIdQuestione();
  var now = new Date();
  
  // Valori default
  var questione = {
    id: idQuestione,
    timestampCreazione: now,
    fonte: dati.fonte || QUESTIONI_CONFIG.FONTI_QUESTIONE.MANUALE,
    
    idFornitore: dati.idFornitore || "",
    nomeFornitore: dati.nomeFornitore || "",
    emailFornitore: dati.emailFornitore || "",
    
    titolo: dati.titolo,
    descrizione: dati.descrizione || "",
    categoria: dati.categoria || QUESTIONI_CONFIG.CATEGORIE_QUESTIONE.ALTRO,
    tags: dati.tags || "",
    
    priorita: dati.priorita || QUESTIONI_CONFIG.PRIORITA.MEDIA,
    urgente: dati.urgente ? "SI" : "NO",
    status: QUESTIONI_CONFIG.STATUS_QUESTIONE.APERTA,
    
    dataScadenza: dati.dataScadenza || "",
    assegnatoA: dati.assegnatoA || "",
    
    emailCollegate: dati.emailCollegate || "",
    numEmail: dati.numEmail || 0,
    
    timestampRisoluzione: "",
    risoluzioneNote: "",
    
    aiConfidence: dati.aiConfidence || 0,
    aiClusterId: dati.aiClusterId || ""
  };
  
  // Prepara riga
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  
  Object.keys(QUESTIONI_CONFIG.COLONNE_QUESTIONI).forEach(function(key) {
    var colName = QUESTIONI_CONFIG.COLONNE_QUESTIONI[key];
    var colIndex = headers.indexOf(colName);
    
    // Mappa key config a key questione object
    var valueKey = key.toLowerCase().replace(/_([a-z])/g, function(m, p1) {
      return p1.toUpperCase();
    });
    
    // Mapping speciale
    var mappingSpeciale = {
      'idQuestione': 'id',
      'timestampCreazione': 'timestampCreazione',
      'idFornitore': 'idFornitore',
      'nomeFornitore': 'nomeFornitore',
      'emailFornitore': 'emailFornitore',
      'emailCollegate': 'emailCollegate',
      'numEmail': 'numEmail',
      'timestampRisoluzione': 'timestampRisoluzione',
      'risoluzioneNote': 'risoluzioneNote',
      'aiConfidence': 'aiConfidence',
      'aiClusterId': 'aiClusterId',
      'dataScadenza': 'dataScadenza',
      'assegnatoA': 'assegnatoA'
    };
    
    var finalKey = mappingSpeciale[valueKey] || valueKey;
    
    if (colIndex >= 0) {
      while (row.length < colIndex) row.push("");
      row[colIndex] = questione[finalKey] !== undefined ? questione[finalKey] : "";
    }
  });
  
  // Scrivi
  sheet.appendRow(row);
  
  logQuestioni("Creata questione " + idQuestione + ": " + questione.titolo);
  
  return questione;
}

// ═══════════════════════════════════════════════════════════════════════
// ELABORAZIONE EMAIL → QUESTIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa email da EMAIL_IN e crea/aggiorna questioni
 * @returns {Object} {emailProcessate, questioniCreate, errori}
 */
function processaEmailInQuestioni() {
  var risultato = {
    emailProcessate: 0,
    questioniCreate: 0,
    questioniAggiornate: 0,
    errori: 0
  };
  
  logQuestioni("═══ PROCESSO EMAIL → QUESTIONI ═══");
  
  // Invalida cache
  CONNECTOR_EMAIL._cache = { emails: null, timestamp: null };
  
  // 1. Carica email da processare
  var emails = caricaEmailDaProcessare();
  
  if (emails.length === 0) {
    logQuestioni("Nessuna email da processare");
    return risultato;
  }
  
  logQuestioni("Email da processare: " + emails.length);
  
  // 2. Processa ogni email singolarmente (semplificato)
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    
    logQuestioni("Processo email " + (i+1) + "/" + emails.length + ": " + email.id);
    
    try {
      // Crea questione da email
      var questione = creaQuestioneDaEmail(email);
      
      if (questione) {
        risultato.questioniCreate++;
        logQuestioni("✅ Creata questione: " + questione.id);
      }
      
      risultato.emailProcessate++;
      
    } catch(e) {
      logQuestioni("❌ Errore email " + email.id + ": " + e.message);
      risultato.errori++;
    }
  }
  
  logQuestioni("═══ COMPLETATO: " + risultato.questioniCreate + " questioni create ═══");
  
  return risultato;
}

/**
 * Processa gruppo email stesso fornitore
 */
function processaGruppoEmail(dominio, emails) {
  var risultato = {
    emailProcessate: 0,
    questioniCreate: 0,
    questioniAggiornate: 0
  };
  
  // Ottieni questioni aperte per questo fornitore
  var questioniAperte = getQuestioniByFornitore(emails[0].mittente);
  questioniAperte = questioniAperte.filter(function(q) {
    return q.status === QUESTIONI_CONFIG.STATUS_QUESTIONE.APERTA ||
           q.status === QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_LAVORAZIONE ||
           q.status === QUESTIONI_CONFIG.STATUS_QUESTIONE.IN_ATTESA_RISPOSTA;
  });
  
  // Se poche email e nessuna questione aperta, crea direttamente senza AI
  if (emails.length <= 2 && questioniAperte.length === 0) {
    emails.forEach(function(email) {
      var questione = creaQuestioneDaEmail(email);
      if (questione) {
        risultato.questioniCreate++;
      }
      risultato.emailProcessate++;
    });
    return risultato;
  }
  
  // Se ci sono questioni aperte, prova matching
  if (questioniAperte.length > 0) {
    emails.forEach(function(email) {
      var matchResult = trovaMatchQuestione(email, questioniAperte);
      
      if (matchResult && matchResult.match && matchResult.confidence >= QUESTIONI_CONFIG.DEFAULTS.MATCH_THRESHOLD) {
        // Collega a questione esistente
        collegaEmailAQuestione(matchResult.idQuestione, email.id);
        risultato.questioniAggiornate++;
      } else {
        // Crea nuova questione
        var questione = creaQuestioneDaEmail(email);
        if (questione) {
          risultato.questioniCreate++;
          // Aggiungi alle aperte per prossimi match
          questioniAperte.push(questione);
        }
      }
      risultato.emailProcessate++;
    });
  } else {
    // Se molte email senza questioni aperte, usa clustering AI
    if (emails.length >= 3) {
      var clusters = clusterizzaEmail(emails);
      
      clusters.forEach(function(cluster) {
        var questione = creaQuestioneDaCluster(cluster, emails);
        if (questione) {
          risultato.questioniCreate++;
        }
        risultato.emailProcessate += cluster.emailIds.length;
      });
    } else {
      // Poche email, crea una questione per ciascuna
      emails.forEach(function(email) {
        var questione = creaQuestioneDaEmail(email);
        if (questione) {
          risultato.questioniCreate++;
        }
        risultato.emailProcessate++;
      });
    }
  }
  
  return risultato;
}

/**
 * Crea questione da singola email
 */
function creaQuestioneDaEmail(email) {
  try {
    var scores = email.scores || {};
    
    var datiQuestione = {
      titolo: generaTitoloQuestione(email),
      descrizione: email.sintesi || email.oggetto,
      fonte: QUESTIONI_CONFIG.FONTI_QUESTIONE.EMAIL,
      emailFornitore: email.mittente,
      nomeFornitore: estraiNomeFornitore(email.mittente),
      categoria: determinaCategoria(email.tags, scores),
      priorita: determinaPriorita(scores),
      urgente: isUrgente(scores),
      tags: email.tags || "",
      emailCollegate: email.id,
      numEmail: 1,
      aiConfidence: email.confidence || 0
    };
    
    var questione = creaQuestione(datiQuestione);
    
    // Aggiorna email locale
    collegaEmailAQuestioneLocale(email.id, questione.id);
    
    return questione;
    
  } catch(e) {
    logQuestioni("Errore creaQuestioneDaEmail: " + e.message);
    return null;
  }
}

/**
 * Crea questione da cluster AI
 */
function creaQuestioneDaCluster(cluster, tutteEmail) {
  try {
    // Trova email del cluster
    var emailCluster = tutteEmail.filter(function(e) {
      return cluster.emailIds.indexOf(e.id) >= 0;
    });
    
    if (emailCluster.length === 0) return null;
    
    // Usa prima email come riferimento
    var primaEmail = emailCluster[0];
    
    var datiQuestione = {
      titolo: cluster.titolo || generaTitoloQuestione(primaEmail),
      descrizione: cluster.descrizione || primaEmail.sintesi,
      fonte: QUESTIONI_CONFIG.FONTI_QUESTIONE.EMAIL,
      emailFornitore: primaEmail.mittente,
      nomeFornitore: estraiNomeFornitore(primaEmail.mittente),
      categoria: cluster.categoria || determinaCategoria(primaEmail.tags, primaEmail.scores),
      priorita: cluster.priorita || determinaPriorita(primaEmail.scores),
      urgente: cluster.urgente || false,
      emailCollegate: cluster.emailIds.join(", "),
      numEmail: cluster.emailIds.length,
      aiClusterId: "CLU-" + cluster.id
    };
    
    var questione = creaQuestione(datiQuestione);
    
    // Aggiorna tutte le email del cluster
    emailCluster.forEach(function(email) {
      collegaEmailAQuestioneLocale(email.id, questione.id);
    });
    
    return questione;
    
  } catch(e) {
    logQuestioni("Errore creaQuestioneDaCluster: " + e.message);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════════════════

/**
 * Genera titolo questione da email
 */
function generaTitoloQuestione(email) {
  var oggetto = email.oggetto || "";
  
  // Rimuovi prefissi comuni
  oggetto = oggetto.replace(/^(Re:|Fwd:|R:|I:)\s*/gi, "");
  
  // Limita lunghezza
  if (oggetto.length > 60) {
    oggetto = oggetto.substring(0, 57) + "...";
  }
  
  return oggetto || "Questione da " + estraiNomeFornitore(email.mittente);
}

/**
 * Estrai nome fornitore da email
 */
function estraiNomeFornitore(email) {
  if (!email) return "Sconosciuto";
  
  var dominio = estraiDominioEmail(email);
  
  // Rimuovi estensione
  var nome = dominio.split(".")[0];
  
  // Capitalizza
  return nome.charAt(0).toUpperCase() + nome.slice(1);
}