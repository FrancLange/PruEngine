/**
 * ==========================================================================================
 * DISCOVERY FORNITORI v2.0.0
 * ==========================================================================================
 * Analizza mittenti sconosciuti in LOG_IN e crea bozze GIA' RAGGRUPPATE per dominio
 * 
 * OUTPUT: Foglio BOZZE_FORNITORE (stesso formato Fornitori_Engine + STATUS_BOZZA)
 * 
 * ESEMPIO RAGGRUPPAMENTO:
 *   marco@domusoleatoscana.it  \
 *   alessio@domusoleatoscana.it } → 1 RIGA: EMAIL_ORDINI=info@... EMAIL_ALTRI=marco@...,alessio@...
 *   info@domusoleatoscana.it   /
 * 
 * WORKFLOW:
 *   1. Esegui: discoveryFornitori()
 *   2. Vai su BOZZE_FORNITORE, rivedi e modifica STATUS_BOZZA → "APPROVATA"
 *   3. Esegui: esportaFornitoriApprovati() → output pronto per copia in FORNITORI master
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CONFIGURAZIONE
// ═══════════════════════════════════════════════════════════════════════

var DISCOVERY = {
  
  // Foglio unico per le bozze
  FOGLIO_BOZZE: "BOZZE_FORNITORE",
  
  // Colonne output = Fornitori_Engine + campi gestione bozze
  COLONNE: [
    // --- CAMPI GESTIONE BOZZE (primi 3) ---
    "STATUS_BOZZA",           // NUOVA, APPROVATA, SCARTATA, IMPORTATA
    "EMAIL_COUNT",            // Totale email ricevute da questo fornitore
    "OGGETTI_ESEMPIO",        // Esempi oggetti email (per capire chi è)
    
    // --- CAMPI FORNITORI_ENGINE (identici) ---
    "ID_FORNITORE",
    "PRIORITA_URGENTE",
    "NOME_AZIENDA",
    "EMAIL_ORDINI",
    "EMAIL_ALTRI",
    "CONTATTO",
    "TELEFONO",
    "PARTITA_IVA",
    "INDIRIZZO",
    "NOTE",
    "METODO_INVIO",
    "TIPO_LISTINO",
    "LINK_LISTINO",
    "MIN_ORDINE",
    "LEAD_TIME_GIORNI",
    "PROMO_SU_RICHIESTA",
    "DATA_PROSSIMA_PROMO",
    "PROMO_ANNUALI_NUM",
    "PROMO_ANNUALI_RICHIESTE",
    "SCONTO_TARGET",
    "RICHIESTA_SCONTO",
    "ESITO_RICHIESTA_SCONTO",
    "TIPO_SCONTO_CONCESSO",
    "SCONTO_PERCENTUALE",
    "DATA_ULTIMO_ORDINE",
    "DATA_ULTIMA_EMAIL",
    "DATA_ULTIMA_ANALISI",
    "STATUS_ULTIMA_AZIONE",
    "EMAIL_ANALIZZATE_COUNT",
    "PERFORMANCE_SCORE",
    "STATUS_FORNITORE",
    "DATA_CREAZIONE"
  ],
  
  // Status bozze
  STATUS: {
    NUOVA: "NUOVA",
    APPROVATA: "APPROVATA",
    SCARTATA: "SCARTATA",
    IMPORTATA: "IMPORTATA"
  },
  
  // Domini da ignorare (email personali)
  DOMINI_BLACKLIST: [
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com",
    "libero.it", "virgilio.it", "alice.it", "tin.it",
    "tiscali.it", "fastwebnet.it", "icloud.com", "me.com",
    "live.com", "msn.com", "pec.it", "legalmail.it", "arubapec.it"
  ],
  
  // Prefissi email prioritari (per scegliere EMAIL_ORDINI)
  PREFISSI_PRIORITA: [
    "ordini", "orders", "commerciale", "sales", "vendite",
    "info", "contatti", "contact", "amministrazione", "admin"
  ]
};

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONE PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * 🚀 ESEGUI QUESTA - Analizza LOG_IN e crea bozze raggruppate per dominio
 */
function discoveryFornitori() {
  var startTime = new Date();
  
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("🔍 DISCOVERY FORNITORI v2.0 - START");
  Logger.log("═══════════════════════════════════════════════════════════════");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Carica email da LOG_IN
  var sheetLogIn = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  if (!sheetLogIn || sheetLogIn.getLastRow() <= 1) {
    Logger.log("❌ LOG_IN vuoto o non trovato");
    return { success: false, error: "LOG_IN non trovato" };
  }
  
  // 2. Estrai tutti i mittenti con statistiche
  Logger.log("📧 Estrazione mittenti da LOG_IN...");
  var mittentiRaw = estraiMittentiDaLogIn(sheetLogIn);
  Logger.log("   Mittenti unici trovati: " + Object.keys(mittentiRaw).length);
  
  // 3. Carica fornitori già noti (per escluderli)
  var fornitoriNoti = caricaFornitoriNoti(ss);
  Logger.log("✅ Fornitori già noti: " + Object.keys(fornitoriNoti).length);
  
  // 4. Filtra e raggruppa per dominio
  Logger.log("🔗 Raggruppamento per dominio...");
  var dominiRaggruppati = raggruppaPerDominio(mittentiRaw, fornitoriNoti);
  Logger.log("   Domini sconosciuti: " + Object.keys(dominiRaggruppati).length);
  
  if (Object.keys(dominiRaggruppati).length === 0) {
    Logger.log("✅ Nessun nuovo fornitore da aggiungere!");
    Logger.log("═══════════════════════════════════════════════════════════════");
    return { success: true, nuovi: 0, aggiornati: 0 };
  }
  
  // 5. Crea/aggiorna foglio BOZZE_FORNITORE
  var sheetBozze = getOrCreateSheetBozze(ss);
  
  // 6. Carica bozze esistenti (per update invece di duplicare)
  var bozzeEsistenti = caricaBozzeEsistenti(sheetBozze);
  Logger.log("📋 Bozze esistenti: " + Object.keys(bozzeEsistenti).length);
  
  // 7. Scrivi nuove bozze / aggiorna esistenti
  var risultato = scriviOAggiornaBozze(sheetBozze, dominiRaggruppati, bozzeEsistenti);
  
  // 8. Report finale
  var durata = Math.round((new Date() - startTime) / 1000);
  
  Logger.log("");
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("✅ DISCOVERY COMPLETATO in " + durata + " secondi");
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("📊 RISULTATI:");
  Logger.log("   🆕 Nuove bozze create: " + risultato.nuovi);
  Logger.log("   📝 Bozze aggiornate: " + risultato.aggiornati);
  Logger.log("");
  Logger.log("👉 PROSSIMI STEP:");
  Logger.log("   1. Vai al foglio: " + DISCOVERY.FOGLIO_BOZZE);
  Logger.log("   2. Rivedi i fornitori e modifica NOME_AZIENDA se necessario");
  Logger.log("   3. Cambia STATUS_BOZZA → 'APPROVATA' per quelli da importare");
  Logger.log("   4. Esegui: esportaFornitoriApprovati()");
  Logger.log("═══════════════════════════════════════════════════════════════");
  
  logSistema("🔍 Discovery: " + risultato.nuovi + " nuovi, " + risultato.aggiornati + " aggiornati");
  
  return { success: true, nuovi: risultato.nuovi, aggiornati: risultato.aggiornati };
}

// ═══════════════════════════════════════════════════════════════════════
// ESTRAZIONE MITTENTI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Estrae tutti i mittenti da LOG_IN con statistiche
 */
function estraiMittentiDaLogIn(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colMittente = headers.indexOf(CONFIG.COLONNE_LOG_IN.MITTENTE);
  var colTimestamp = headers.indexOf(CONFIG.COLONNE_LOG_IN.TIMESTAMP);
  var colOggetto = headers.indexOf(CONFIG.COLONNE_LOG_IN.OGGETTO);
  
  var mittenti = {};
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var email = (row[colMittente] || "").toString().toLowerCase().trim();
    
    if (!email || email === "") continue;
    
    var timestamp = row[colTimestamp];
    var oggetto = (row[colOggetto] || "").toString();
    
    if (!mittenti[email]) {
      mittenti[email] = {
        email: email,
        count: 0,
        primaData: timestamp,
        ultimaData: timestamp,
        oggetti: []
      };
    }
    
    var m = mittenti[email];
    m.count++;
    
    if (timestamp && (!m.primaData || timestamp < m.primaData)) {
      m.primaData = timestamp;
    }
    if (timestamp && (!m.ultimaData || timestamp > m.ultimaData)) {
      m.ultimaData = timestamp;
    }
    
    // Max 3 oggetti esempio
    if (m.oggetti.length < 3 && oggetto) {
      m.oggetti.push(oggetto.substring(0, 60));
    }
  }
  
  return mittenti;
}

// ═══════════════════════════════════════════════════════════════════════
// RAGGRUPPAMENTO PER DOMINIO
// ═══════════════════════════════════════════════════════════════════════

/**
 * Raggruppa mittenti per dominio, escludendo quelli già noti
 */
function raggruppaPerDominio(mittenti, fornitoriNoti) {
  var domini = {};
  
  var emailKeys = Object.keys(mittenti);
  
  for (var i = 0; i < emailKeys.length; i++) {
    var email = emailKeys[i];
    var m = mittenti[email];
    var dominio = estraiDominio(email);
    
    // Skip se no dominio
    if (!dominio) continue;
    
    // Skip se dominio in blacklist
    if (DISCOVERY.DOMINI_BLACKLIST.indexOf(dominio) >= 0) continue;
    
    // Skip se già noto (check email E dominio)
    if (fornitoriNoti[email] || fornitoriNoti[dominio]) continue;
    
    // Crea gruppo dominio se non esiste
    if (!domini[dominio]) {
      domini[dominio] = {
        dominio: dominio,
        emails: [],
        totaleCount: 0,
        primaData: null,
        ultimaData: null,
        oggetti: []
      };
    }
    
    var d = domini[dominio];
    
    // Aggiungi email al gruppo
    d.emails.push({
      email: email,
      count: m.count
    });
    
    // Aggiorna statistiche
    d.totaleCount += m.count;
    
    if (m.primaData && (!d.primaData || m.primaData < d.primaData)) {
      d.primaData = m.primaData;
    }
    if (m.ultimaData && (!d.ultimaData || m.ultimaData > d.ultimaData)) {
      d.ultimaData = m.ultimaData;
    }
    
    // Aggiungi oggetti (max 5 per dominio)
    for (var o = 0; o < m.oggetti.length && d.oggetti.length < 5; o++) {
      d.oggetti.push(m.oggetti[o]);
    }
  }
  
  // Determina email principale per ogni dominio
  var dominiKeys = Object.keys(domini);
  for (var j = 0; j < dominiKeys.length; j++) {
    var dom = domini[dominiKeys[j]];
    var risultatoEmail = determinaEmailPrincipaleESecondarie(dom.emails);
    dom.emailPrincipale = risultatoEmail.principale;
    dom.emailSecondarie = risultatoEmail.secondarie;
  }
  
  return domini;
}

/**
 * Determina email principale (ordini/info/...) e secondarie
 */
function determinaEmailPrincipaleESecondarie(emails) {
  if (!emails || emails.length === 0) {
    return { principale: "", secondarie: "" };
  }
  
  if (emails.length === 1) {
    return { principale: emails[0].email, secondarie: "" };
  }
  
  // Ordina per priorità prefisso, poi per count
  var emailsOrdinata = emails.slice().sort(function(a, b) {
    var prefA = a.email.split("@")[0].toLowerCase();
    var prefB = b.email.split("@")[0].toLowerCase();
    
    var prioA = 999;
    var prioB = 999;
    
    for (var p = 0; p < DISCOVERY.PREFISSI_PRIORITA.length; p++) {
      var pref = DISCOVERY.PREFISSI_PRIORITA[p];
      if (prefA.indexOf(pref) >= 0 && prioA === 999) prioA = p;
      if (prefB.indexOf(pref) >= 0 && prioB === 999) prioB = p;
    }
    
    // Se stessa priorità, quello con più email vince
    if (prioA === prioB) {
      return b.count - a.count;
    }
    
    return prioA - prioB;
  });
  
  var principale = emailsOrdinata[0].email;
  var secondarie = [];
  
  for (var i = 1; i < emailsOrdinata.length; i++) {
    secondarie.push(emailsOrdinata[i].email);
  }
  
  return {
    principale: principale,
    secondarie: secondarie.join(", ")
  };
}

// ═══════════════════════════════════════════════════════════════════════
// GESTIONE FOGLIO BOZZE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea o recupera il foglio BOZZE_FORNITORE
 */
function getOrCreateSheetBozze(ss) {
  var sheet = ss.getSheetByName(DISCOVERY.FOGLIO_BOZZE);
  
  if (!sheet) {
    sheet = ss.insertSheet(DISCOVERY.FOGLIO_BOZZE);
    sheet.setTabColor("#F39C12"); // Arancione
    
    // Header
    sheet.getRange(1, 1, 1, DISCOVERY.COLONNE.length)
      .setValues([DISCOVERY.COLONNE])
      .setFontWeight("bold")
      .setBackground("#FCF3CF");
    
    sheet.setFrozenRows(1);
    
    // Larghezze colonne principali
    sheet.setColumnWidth(1, 100);  // STATUS_BOZZA
    sheet.setColumnWidth(2, 80);   // EMAIL_COUNT
    sheet.setColumnWidth(3, 300);  // OGGETTI_ESEMPIO
    sheet.setColumnWidth(4, 120);  // ID_FORNITORE
    sheet.setColumnWidth(6, 200);  // NOME_AZIENDA
    sheet.setColumnWidth(7, 220);  // EMAIL_ORDINI
    sheet.setColumnWidth(8, 300);  // EMAIL_ALTRI
    
    Logger.log("✅ Creato foglio: " + DISCOVERY.FOGLIO_BOZZE);
  }
  
  return sheet;
}

/**
 * Carica bozze esistenti (per evitare duplicati)
 * Ritorna mappa: dominio → rowNum
 */
function caricaBozzeEsistenti(sheet) {
  var bozze = {};
  
  if (sheet.getLastRow() <= 1) return bozze;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colEmailOrdini = headers.indexOf("EMAIL_ORDINI");
  
  for (var i = 0; i < rows.length; i++) {
    var emailOrdini = (rows[i][colEmailOrdini] || "").toString().toLowerCase().trim();
    if (emailOrdini) {
      var dominio = estraiDominio(emailOrdini);
      if (dominio) {
        bozze[dominio] = i + 2; // Row number (1-based + header)
      }
    }
  }
  
  return bozze;
}

/**
 * Scrivi nuove bozze o aggiorna esistenti
 */
function scriviOAggiornaBozze(sheet, dominiRaggruppati, bozzeEsistenti) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var now = new Date();
  
  var nuovi = 0;
  var aggiornati = 0;
  
  var dominiKeys = Object.keys(dominiRaggruppati);
  
  for (var i = 0; i < dominiKeys.length; i++) {
    var dominio = dominiKeys[i];
    var d = dominiRaggruppati[dominio];
    
    if (bozzeEsistenti[dominio]) {
      // AGGIORNA esistente
      aggiornaBozza(sheet, headers, bozzeEsistenti[dominio], d, now);
      aggiornati++;
      Logger.log("📝 Aggiornato: " + dominio + " (" + d.emails.length + " email)");
    } else {
      // CREA nuova
      creaNuovaBozza(sheet, d, now);
      nuovi++;
      Logger.log("🆕 Nuovo: " + dominio + " → " + d.emailPrincipale + " (" + d.totaleCount + " msg)");
    }
  }
  
  return { nuovi: nuovi, aggiornati: aggiornati };
}

/**
 * Crea nuova riga bozza
 */
function creaNuovaBozza(sheet, d, now) {
  var nomeSuggerito = estraiNomeDaDominio(d.dominio);
  
  // Costruisci riga con tutte le colonne
  var row = [
    DISCOVERY.STATUS.NUOVA,                    // STATUS_BOZZA
    d.totaleCount,                             // EMAIL_COUNT
    d.oggetti.join(" | "),                     // OGGETTI_ESEMPIO
    "",                                        // ID_FORNITORE (vuoto, da assegnare)
    "",                                        // PRIORITA_URGENTE
    nomeSuggerito,                             // NOME_AZIENDA
    d.emailPrincipale,                         // EMAIL_ORDINI
    d.emailSecondarie,                         // EMAIL_ALTRI
    "",                                        // CONTATTO
    "",                                        // TELEFONO
    "",                                        // PARTITA_IVA
    "",                                        // INDIRIZZO
    "Scoperto automaticamente da Email Engine", // NOTE
    "",                                        // METODO_INVIO
    "",                                        // TIPO_LISTINO
    "",                                        // LINK_LISTINO
    "",                                        // MIN_ORDINE
    "",                                        // LEAD_TIME_GIORNI
    "",                                        // PROMO_SU_RICHIESTA
    "",                                        // DATA_PROSSIMA_PROMO
    "",                                        // PROMO_ANNUALI_NUM
    "",                                        // PROMO_ANNUALI_RICHIESTE
    "",                                        // SCONTO_TARGET
    "",                                        // RICHIESTA_SCONTO
    "",                                        // ESITO_RICHIESTA_SCONTO
    "",                                        // TIPO_SCONTO_CONCESSO
    "",                                        // SCONTO_PERCENTUALE
    "",                                        // DATA_ULTIMO_ORDINE
    d.ultimaData,                              // DATA_ULTIMA_EMAIL
    now,                                       // DATA_ULTIMA_ANALISI
    "Nuovo da Discovery",                      // STATUS_ULTIMA_AZIONE
    d.totaleCount,                             // EMAIL_ANALIZZATE_COUNT
    "",                                        // PERFORMANCE_SCORE
    "NUOVO",                                   // STATUS_FORNITORE
    now                                        // DATA_CREAZIONE
  ];
  
  sheet.appendRow(row);
}

/**
 * Aggiorna bozza esistente
 */
function aggiornaBozza(sheet, headers, rowNum, d, now) {
  // Aggiorna solo campi dinamici
  var updates = {
    "EMAIL_COUNT": d.totaleCount,
    "EMAIL_ALTRI": d.emailSecondarie,
    "DATA_ULTIMA_EMAIL": d.ultimaData,
    "DATA_ULTIMA_ANALISI": now,
    "EMAIL_ANALIZZATE_COUNT": d.totaleCount
  };
  
  for (var key in updates) {
    var colIndex = headers.indexOf(key);
    if (colIndex >= 0) {
      sheet.getRange(rowNum, colIndex + 1).setValue(updates[key]);
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CARICAMENTO FORNITORI NOTI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica fornitori già presenti in FORNITORI_SYNC
 * Ritorna mappa: email/dominio → true
 */
function caricaFornitoriNoti(ss) {
  var noti = {};
  
  // Prova FORNITORI_SYNC (Connector)
  var tabSync = "FORNITORI_SYNC";
  if (typeof CONNECTOR_FORNITORI !== "undefined" && CONNECTOR_FORNITORI.SHEETS) {
    tabSync = CONNECTOR_FORNITORI.SHEETS.FORNITORI_SYNC;
  }
  
  var sheet = ss.getSheetByName(tabSync);
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("⚠️ " + tabSync + " non trovato o vuoto");
    return noti;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Trova colonne email
  var colEmailOrdini = headers.indexOf("EMAIL_ORDINI");
  var colEmailAltri = headers.indexOf("EMAIL_ALTRI");
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    
    // Email ordini
    if (colEmailOrdini >= 0) {
      var emailOrd = (row[colEmailOrdini] || "").toString().toLowerCase().trim();
      if (emailOrd) {
        noti[emailOrd] = true;
        var dom = estraiDominio(emailOrd);
        if (dom) noti[dom] = true;
      }
    }
    
    // Email altri
    if (colEmailAltri >= 0) {
      var emailAltri = (row[colEmailAltri] || "").toString().toLowerCase();
      var lista = emailAltri.split(/[,;]/);
      for (var j = 0; j < lista.length; j++) {
        var e = lista[j].trim();
        if (e) {
          noti[e] = true;
          var dom2 = estraiDominio(e);
          if (dom2) noti[dom2] = true;
        }
      }
    }
  }
  
  return noti;
}

// ═══════════════════════════════════════════════════════════════════════
// EXPORT FORNITORI APPROVATI
// ═══════════════════════════════════════════════════════════════════════

/**
 * 🚀 ESEGUI DOPO APPROVAZIONE - Esporta fornitori con STATUS_BOZZA = APPROVATA
 */
function esportaFornitoriApprovati() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DISCOVERY.FOGLIO_BOZZE);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("❌ Foglio " + DISCOVERY.FOGLIO_BOZZE + " non trovato o vuoto");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colStatus = headers.indexOf("STATUS_BOZZA");
  var colNome = headers.indexOf("NOME_AZIENDA");
  var colEmailOrd = headers.indexOf("EMAIL_ORDINI");
  var colEmailAltri = headers.indexOf("EMAIL_ALTRI");
  var colCount = headers.indexOf("EMAIL_COUNT");
  
  // Trova approvati
  var approvati = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][colStatus] === DISCOVERY.STATUS.APPROVATA) {
      approvati.push({
        rowNum: i + 2,
        row: rows[i],
        nome: rows[i][colNome],
        emailOrd: rows[i][colEmailOrd],
        emailAltri: rows[i][colEmailAltri],
        count: rows[i][colCount]
      });
    }
  }
  
  if (approvati.length === 0) {
    Logger.log("═══════════════════════════════════════════════════════════════");
    Logger.log("⚠️ NESSUN FORNITORE CON STATUS 'APPROVATA'");
    Logger.log("═══════════════════════════════════════════════════════════════");
    Logger.log("");
    Logger.log("Per approvare un fornitore:");
    Logger.log("  1. Vai al foglio: " + DISCOVERY.FOGLIO_BOZZE);
    Logger.log("  2. Trova la colonna STATUS_BOZZA");
    Logger.log("  3. Cambia da 'NUOVA' a 'APPROVATA'");
    Logger.log("  4. Esegui di nuovo questa funzione");
    Logger.log("");
    return;
  }
  
  // Header export (colonne Fornitori_Engine, senza i 3 campi discovery)
  var colonneExport = DISCOVERY.COLONNE.slice(3); // Rimuovi primi 3
  
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("📤 EXPORT FORNITORI APPROVATI");
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("");
  Logger.log("Copia le righe qui sotto nel foglio FORNITORI del Master:");
  Logger.log("");
  Logger.log("─────────────────────────────────────────────────────────────────");
  
  // Genera output TSV (tab-separated, facile da incollare)
  for (var j = 0; j < approvati.length; j++) {
    var a = approvati[j];
    var rowData = a.row.slice(3); // Rimuovi primi 3 campi discovery
    
    // Genera ID se vuoto
    if (!rowData[0]) {
      var idNum = (j + 1).toString();
      while (idNum.length < 3) idNum = "0" + idNum;
      rowData[0] = "FOR-NEW-" + idNum;
    }
    
    Logger.log(rowData.join("\t"));
  }
  
  Logger.log("─────────────────────────────────────────────────────────────────");
  Logger.log("");
  Logger.log("📊 Totale: " + approvati.length + " fornitori pronti per import");
  Logger.log("");
  Logger.log("👉 DOPO L'IMPORT:");
  Logger.log("   1. Incolla le righe nel foglio FORNITORI del Master");
  Logger.log("   2. Esegui sync Connector per aggiornare FORNITORI_SYNC");
  Logger.log("   3. Opzionale: esegui marcaImportati() per aggiornare status bozze");
  Logger.log("═══════════════════════════════════════════════════════════════");
  
  logSistema("📤 Export: " + approvati.length + " fornitori approvati");
}

/**
 * Marca i fornitori approvati come IMPORTATI
 */
function marcaImportati() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DISCOVERY.FOGLIO_BOZZE);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("❌ Foglio bozze non trovato");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colStatus = headers.indexOf("STATUS_BOZZA");
  
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][colStatus] === DISCOVERY.STATUS.APPROVATA) {
      sheet.getRange(i + 1, colStatus + 1).setValue(DISCOVERY.STATUS.IMPORTATA);
      count++;
    }
  }
  
  Logger.log("✅ Marcati come IMPORTATI: " + count + " fornitori");
  logSistema("Marcati importati: " + count + " fornitori");
}

// ═══════════════════════════════════════════════════════════════════════
// REPORT
// ═══════════════════════════════════════════════════════════════════════

/**
 * 🚀 Report statistiche bozze
 */
function reportBozze() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DISCOVERY.FOGLIO_BOZZE);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("📋 Nessuna bozza presente");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colStatus = headers.indexOf("STATUS_BOZZA");
  var colCount = headers.indexOf("EMAIL_COUNT");
  
  var stats = {
    totali: rows.length,
    nuove: 0,
    approvate: 0,
    scartate: 0,
    importate: 0,
    emailTotali: 0
  };
  
  for (var i = 0; i < rows.length; i++) {
    var status = rows[i][colStatus];
    var count = parseInt(rows[i][colCount]) || 0;
    
    stats.emailTotali += count;
    
    switch (status) {
      case DISCOVERY.STATUS.NUOVA: stats.nuove++; break;
      case DISCOVERY.STATUS.APPROVATA: stats.approvate++; break;
      case DISCOVERY.STATUS.SCARTATA: stats.scartate++; break;
      case DISCOVERY.STATUS.IMPORTATA: stats.importate++; break;
    }
  }
  
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("📊 REPORT BOZZE FORNITORE");
  Logger.log("═══════════════════════════════════════════════════════════════");
  Logger.log("");
  Logger.log("📋 Totale bozze: " + stats.totali);
  Logger.log("   🆕 Nuove (da valutare): " + stats.nuove);
  Logger.log("   ✅ Approvate (pronte): " + stats.approvate);
  Logger.log("   ❌ Scartate: " + stats.scartate);
  Logger.log("   📥 Importate: " + stats.importate);
  Logger.log("");
  Logger.log("📧 Email totali da fornitori sconosciuti: " + stats.emailTotali);
  Logger.log("═══════════════════════════════════════════════════════════════");
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════

function estraiDominio(email) {
  if (!email) return "";
  email = email.toString().toLowerCase().trim();
  var parts = email.split("@");
  return parts.length > 1 ? parts[1] : "";
}

function estraiNomeDaDominio(dominio) {
  if (!dominio) return "";
  // Rimuovi TLD
  var nome = dominio.replace(/\.(com|it|eu|net|org|co\.uk|de|fr|es|info|biz)$/i, "");
  // Prendi prima parte
  nome = nome.split(".")[0];
  // Capitalizza
  nome = nome.charAt(0).toUpperCase() + nome.slice(1);
  // Sostituisci separatori con spazi
  nome = nome.replace(/[-_]/g, " ");
  return nome;
}

function logSistema(messaggio) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LOG_SISTEMA");
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}