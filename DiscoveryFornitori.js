/**
 * ==========================================================================================
 * DISCOVERY FORNITORI - Identificazione Mittenti Sconosciuti v1.0.0
 * ==========================================================================================
 * Funzione standalone per:
 * - Analizzare mittenti email in LOG_IN
 * - Identificare quelli NON presenti in FORNITORI_SYNC
 * - Creare bozze di nuovi fornitori da validare manualmente
 * 
 * UTILIZZO: Eseguire manualmente da Apps Script Editor
 *   > Seleziona funzione: analizzaMittentiSconosciuti
 *   > Clicca "Esegui"
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// CONFIGURAZIONE
// ═══════════════════════════════════════════════════════════════════════

var DISCOVERY_CONFIG = {
  // Foglio per le bozze
  FOGLIO_BOZZE: "FORNITORI_BOZZE",
  
  // Colonne del foglio bozze
  COLONNE_BOZZE: [
    "ID_BOZZA",
    "EMAIL_MITTENTE",
    "DOMINIO",
    "NOME_ESTRATTO",
    "EMAIL_COUNT",
    "PRIMA_EMAIL_DATA",
    "ULTIMA_EMAIL_DATA",
    "OGGETTI_ESEMPIO",
    "STATUS_BOZZA",      // NUOVA, APPROVATA, SCARTATA, IMPORTATA
    "NOTE",
    "DATA_CREAZIONE",
    "DATA_MODIFICA"
  ],
  
  // Status possibili
  STATUS: {
    NUOVA: "NUOVA",
    APPROVATA: "APPROVATA",
    SCARTATA: "SCARTATA",
    IMPORTATA: "IMPORTATA"
  },
  
  // Domini da ignorare (email personali, generiche)
  DOMINI_BLACKLIST: [
    "gmail.com",
    "yahoo.com",
    "hotmail.com",
    "outlook.com",
    "libero.it",
    "virgilio.it",
    "alice.it",
    "tin.it",
    "tiscali.it",
    "fastwebnet.it",
    "icloud.com",
    "me.com",
    "live.com",
    "msn.com"
  ]
};

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONE PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * 🚀 FUNZIONE PRINCIPALE - Esegui questa da Apps Script
 * Analizza tutti i mittenti in LOG_IN e crea bozze per quelli sconosciuti
 */
function analizzaMittentiSconosciuti() {
  var startTime = new Date();
  
  Logger.log("═══════════════════════════════════════════");
  Logger.log("🔍 DISCOVERY FORNITORI - START");
  Logger.log("═══════════════════════════════════════════");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Carica email da LOG_IN
  var sheetLogIn = ss.getSheetByName(CONFIG.SHEETS.LOG_IN);
  if (!sheetLogIn || sheetLogIn.getLastRow() <= 1) {
    Logger.log("❌ LOG_IN vuoto o non trovato");
    return;
  }
  
  // 2. Estrai tutti i mittenti unici con statistiche
  var mittentiStats = estraiMittentiConStats(sheetLogIn);
  Logger.log("📧 Mittenti unici trovati: " + Object.keys(mittentiStats).length);
  
  // 3. Carica fornitori noti
  var fornitoriNoti = caricaFornitoriNotiDiscovery(ss);
  Logger.log("✅ Fornitori noti in sync: " + Object.keys(fornitoriNoti).length);
  
  // 4. Identifica mittenti sconosciuti
  var mittentiSconosciuti = [];
  var mittentiKeys = Object.keys(mittentiStats);
  
  for (var i = 0; i < mittentiKeys.length; i++) {
    var email = mittentiKeys[i];
    var stats = mittentiStats[email];
    var dominio = estraiDominioDiscovery(email);
    
    // Skip se dominio in blacklist
    if (DISCOVERY_CONFIG.DOMINI_BLACKLIST.indexOf(dominio) >= 0) {
      Logger.log("⏭️ Skip (blacklist): " + email);
      continue;
    }
    
    // Skip se già fornitore noto (check email e dominio)
    if (fornitoriNoti[email] || fornitoriNoti[dominio]) {
      Logger.log("✅ Già noto: " + email);
      continue;
    }
    
    // Mittente sconosciuto!
    mittentiSconosciuti.push({
      email: email,
      dominio: dominio,
      stats: stats
    });
  }
  
  Logger.log("🆕 Mittenti sconosciuti: " + mittentiSconosciuti.length);
  
  if (mittentiSconosciuti.length === 0) {
    Logger.log("✅ Nessun nuovo mittente da aggiungere");
    Logger.log("═══════════════════════════════════════════");
    return;
  }
  
  // 5. Crea/aggiorna foglio bozze
  var sheetBozze = getOrCreateFoglioBozze(ss);
  
  // 6. Carica bozze esistenti (per evitare duplicati)
  var bozzeEsistenti = caricaBozzeEsistenti(sheetBozze);
  Logger.log("📋 Bozze esistenti: " + Object.keys(bozzeEsistenti).length);
  
  // 7. Aggiungi nuove bozze
  var nuoveBozze = 0;
  var bozzeAggiornate = 0;
  
  for (var j = 0; j < mittentiSconosciuti.length; j++) {
    var mittente = mittentiSconosciuti[j];
    
    if (bozzeEsistenti[mittente.email]) {
      aggiornaBozzaEsistente(sheetBozze, bozzeEsistenti[mittente.email], mittente);
      bozzeAggiornate++;
    } else {
      creaNuovaBozza(sheetBozze, mittente);
      nuoveBozze++;
    }
  }
  
  // 8. Report finale
  var durata = Math.round((new Date() - startTime) / 1000);
  
  Logger.log("═══════════════════════════════════════════");
  Logger.log("✅ DISCOVERY COMPLETATO in " + durata + "s");
  Logger.log("📊 Nuove bozze create: " + nuoveBozze);
  Logger.log("📊 Bozze aggiornate: " + bozzeAggiornate);
  Logger.log("📋 Controlla foglio: " + DISCOVERY_CONFIG.FOGLIO_BOZZE);
  Logger.log("═══════════════════════════════════════════");
  
  logSistema("🔍 Discovery Fornitori: " + nuoveBozze + " nuove bozze, " + bozzeAggiornate + " aggiornate");
}

// ═══════════════════════════════════════════════════════════════════════
// ESTRAZIONE MITTENTI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Estrae tutti i mittenti unici con statistiche
 */
function estraiMittentiConStats(sheet) {
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
    var oggetto = row[colOggetto] || "";
    
    if (!mittenti[email]) {
      mittenti[email] = {
        count: 0,
        primaData: timestamp,
        ultimaData: timestamp,
        oggetti: []
      };
    }
    
    mittenti[email].count++;
    
    if (timestamp && timestamp < mittenti[email].primaData) {
      mittenti[email].primaData = timestamp;
    }
    if (timestamp && timestamp > mittenti[email].ultimaData) {
      mittenti[email].ultimaData = timestamp;
    }
    
    if (mittenti[email].oggetti.length < 3 && oggetto) {
      mittenti[email].oggetti.push(oggetto.substring(0, 50));
    }
  }
  
  return mittenti;
}

// ═══════════════════════════════════════════════════════════════════════
// CARICAMENTO DATI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica fornitori noti da FORNITORI_SYNC
 */
function caricaFornitoriNotiDiscovery(ss) {
  var fornitoriNoti = {};
  
  var sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("⚠️ FORNITORI_SYNC vuoto, provo sync...");
    
    if (typeof syncFornitoriCompleto === "function") {
      var syncResult = syncFornitoriCompleto();
      if (!syncResult.success) {
        Logger.log("⚠️ Sync fallito, continuo senza fornitori noti");
        return fornitoriNoti;
      }
      sheet = ss.getSheetByName(SYNC_CONFIG.TAB_SYNC);
    }
    
    if (!sheet || sheet.getLastRow() <= 1) return fornitoriNoti;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colEmail = headers.indexOf("EMAIL_PRINCIPALE");
  var colEmailSec = headers.indexOf("EMAIL_SECONDARIE");
  var colDominio = headers.indexOf("DOMINIO_EMAIL");
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    
    if (colEmail >= 0) {
      var email = (row[colEmail] || "").toString().toLowerCase().trim();
      if (email) fornitoriNoti[email] = true;
    }
    
    if (colEmailSec >= 0) {
      var emailSec = (row[colEmailSec] || "").toString().toLowerCase();
      var secList = emailSec.split(/[,;]/);
      for (var s = 0; s < secList.length; s++) {
        var e = secList[s].trim();
        if (e) fornitoriNoti[e] = true;
      }
    }
    
    if (colDominio >= 0) {
      var dominio = (row[colDominio] || "").toString().toLowerCase().trim();
      if (dominio) fornitoriNoti[dominio] = true;
    }
  }
  
  return fornitoriNoti;
}

/**
 * Carica bozze esistenti
 */
function caricaBozzeEsistenti(sheet) {
  var bozze = {};
  
  if (sheet.getLastRow() <= 1) return bozze;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colEmail = headers.indexOf("EMAIL_MITTENTE");
  
  for (var i = 0; i < rows.length; i++) {
    var email = (rows[i][colEmail] || "").toString().toLowerCase().trim();
    if (email) {
      bozze[email] = i + 2;
    }
  }
  
  return bozze;
}

// ═══════════════════════════════════════════════════════════════════════
// GESTIONE FOGLIO BOZZE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Crea o recupera il foglio FORNITORI_BOZZE
 */
function getOrCreateFoglioBozze(ss) {
  var sheet = ss.getSheetByName(DISCOVERY_CONFIG.FOGLIO_BOZZE);
  
  if (!sheet) {
    sheet = ss.insertSheet(DISCOVERY_CONFIG.FOGLIO_BOZZE);
    sheet.setTabColor("#E67E22");
    
    sheet.getRange(1, 1, 1, DISCOVERY_CONFIG.COLONNE_BOZZE.length)
      .setValues([DISCOVERY_CONFIG.COLONNE_BOZZE])
      .setFontWeight("bold")
      .setBackground("#EFEFEF");
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(8, 300);
    sheet.setColumnWidth(9, 100);
    
    Logger.log("✅ Creato foglio: " + DISCOVERY_CONFIG.FOGLIO_BOZZE);
  }
  
  return sheet;
}

/**
 * Crea nuova bozza fornitore
 */
function creaNuovaBozza(sheet, mittente) {
  var now = new Date();
  var idBozza = "BZ-" + Utilities.formatDate(now, "Europe/Rome", "yyyyMMddHHmmss") + "-" + Math.random().toString(36).substr(2, 4).toUpperCase();
  
  var nomeEstratto = estraiNomeDaDominioDiscovery(mittente.dominio);
  
  var row = [
    idBozza,
    mittente.email,
    mittente.dominio,
    nomeEstratto,
    mittente.stats.count,
    mittente.stats.primaData,
    mittente.stats.ultimaData,
    mittente.stats.oggetti.join(" | "),
    DISCOVERY_CONFIG.STATUS.NUOVA,
    "",
    now,
    now
  ];
  
  sheet.appendRow(row);
  Logger.log("🆕 Nuova bozza: " + mittente.email + " (" + mittente.stats.count + " email)");
}

/**
 * Aggiorna bozza esistente con nuove statistiche
 */
function aggiornaBozzaEsistente(sheet, rowNum, mittente) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colCount = headers.indexOf("EMAIL_COUNT") + 1;
  var colUltimaData = headers.indexOf("ULTIMA_EMAIL_DATA") + 1;
  var colModifica = headers.indexOf("DATA_MODIFICA") + 1;
  
  if (colCount > 0) {
    sheet.getRange(rowNum, colCount).setValue(mittente.stats.count);
  }
  if (colUltimaData > 0) {
    sheet.getRange(rowNum, colUltimaData).setValue(mittente.stats.ultimaData);
  }
  if (colModifica > 0) {
    sheet.getRange(rowNum, colModifica).setValue(new Date());
  }
  
  Logger.log("📝 Aggiornata bozza: " + mittente.email + " (" + mittente.stats.count + " email)");
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════

function estraiDominioDiscovery(email) {
  if (!email || email === "") return "";
  email = email.toString().toLowerCase().trim();
  var parts = email.split("@");
  return parts.length > 1 ? parts[1] : "";
}

function estraiNomeDaDominioDiscovery(dominio) {
  if (!dominio) return "";
  var nome = dominio.split(".")[0];
  nome = nome.charAt(0).toUpperCase() + nome.slice(1);
  nome = nome.replace(/[-_]/g, " ");
  return nome;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI AGGIUNTIVE
// ═══════════════════════════════════════════════════════════════════════

/**
 * 🚀 Report bozze - mostra statistiche
 */
function reportBozzeFornitori() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DISCOVERY_CONFIG.FOGLIO_BOZZE);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("❌ Nessuna bozza presente");
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
      case DISCOVERY_CONFIG.STATUS.NUOVA: stats.nuove++; break;
      case DISCOVERY_CONFIG.STATUS.APPROVATA: stats.approvate++; break;
      case DISCOVERY_CONFIG.STATUS.SCARTATA: stats.scartate++; break;
      case DISCOVERY_CONFIG.STATUS.IMPORTATA: stats.importate++; break;
    }
  }
  
  Logger.log("═══════════════════════════════════════════");
  Logger.log("📊 REPORT BOZZE FORNITORI");
  Logger.log("═══════════════════════════════════════════");
  Logger.log("📋 Totale bozze: " + stats.totali);
  Logger.log("🆕 Nuove (da valutare): " + stats.nuove);
  Logger.log("✅ Approvate: " + stats.approvate);
  Logger.log("❌ Scartate: " + stats.scartate);
  Logger.log("📥 Importate: " + stats.importate);
  Logger.log("📧 Email totali da mittenti sconosciuti: " + stats.emailTotali);
  Logger.log("═══════════════════════════════════════════");
}

/**
 * 🚀 Esporta bozze APPROVATE in formato pronto per Fornitori Engine
 */
function esportaBozzeApprovate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DISCOVERY_CONFIG.FOGLIO_BOZZE);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("❌ Nessuna bozza presente");
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  var colStatus = headers.indexOf("STATUS_BOZZA");
  var colEmail = headers.indexOf("EMAIL_MITTENTE");
  var colDominio = headers.indexOf("DOMINIO");
  var colNome = headers.indexOf("NOME_ESTRATTO");
  
  var approvate = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][colStatus] === DISCOVERY_CONFIG.STATUS.APPROVATA) {
      approvate.push(rows[i]);
    }
  }
  
  if (approvate.length === 0) {
    Logger.log("⚠️ Nessuna bozza con status APPROVATA");
    Logger.log("Cambia lo status delle bozze da importare in 'APPROVATA'");
    return;
  }
  
  Logger.log("═══════════════════════════════════════════");
  Logger.log("📥 EXPORT BOZZE APPROVATE");
  Logger.log("═══════════════════════════════════════════");
  Logger.log("Formato: ID_FORNITORE | NOME_AZIENDA | EMAIL_ORDINI | DOMINIO | STATUS");
  Logger.log("───────────────────────────────────────────");
  
  for (var j = 0; j < approvate.length; j++) {
    var row = approvate[j];
    var idNum = (j + 1).toString();
    while (idNum.length < 3) idNum = "0" + idNum;
    var idFornitore = "FOR-NEW-" + idNum;
    var nome = row[colNome] || "Da definire";
    var email = row[colEmail];
    var dominio = row[colDominio];
    
    Logger.log(idFornitore + " | " + nome + " | " + email + " | " + dominio + " | ATTIVO");
  }
  
  Logger.log("───────────────────────────────────────────");
  Logger.log("Totale: " + approvate.length + " fornitori da importare");
  Logger.log("═══════════════════════════════════════════");
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