/**
 * ==========================================================================================
 * DISPATCHER - ORCHESTRATORE AUTOMAZIONI v1.0.1
 * ==========================================================================================
 * Cuore del sistema: gestisce esecuzione jobs, dipendenze, priorità, timeout e tracking.
 * 
 * CHANGELOG v1.0.1:
 * - Rimossi placeholder JOB-002, JOB-003, JOB-004 (implementati in Logic.gs)
 * - Mantenuti solo healthCheckSistema e syncFileEsterni
 * ==========================================================================================
 */

// ═══════════════════════════════════════════════════════════════════════
// DISPATCHER PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════

/**
 * ENTRY POINT - Chiamato dal trigger orario
 * Esegue tutti i job schedulati rispettando dipendenze e priorità
 */
function eseguiDispatcher() {
  const startTime = new Date();
  const MAX_EXEC_TIME = 5.5 * 60 * 1000; // 5.5 minuti (margine sicurezza)
  
  logSistema("═══ DISPATCHER START ═══");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.AUTOMAZIONI);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      logSistema("⚠️ Nessun job configurato");
      return;
    }
    
    // 1. Carica tutti i job
    const jobs = caricaJobs(sheet);
    
    // 2. Filtra job da eseguire ORA
    const jobsDaEseguire = jobs.filter(job => {
      if (job.attiva !== "SI") return false;
      if (job.esito === "IN_CORSO") return false; // Skip se già in esecuzione
      return deveEseguireOra(job);
    });
    
    // 3. Ordina per priorità (più bassa = prima)
    jobsDaEseguire.sort((a, b) => (a.priorita || 99) - (b.priorita || 99));
    
    logSistema(`📋 Jobs in coda: ${jobsDaEseguire.length} / ${jobs.length} totali`);
    
    // 4. Esegui jobs rispettando tempo massimo
    let jobsEseguiti = 0;
    let jobsOk = 0;
    let jobsErrore = 0;
    let jobsSkip = 0;
    
    for (const job of jobsDaEseguire) {
      // Controllo tempo rimanente
      const elapsed = new Date() - startTime;
      if (elapsed > MAX_EXEC_TIME) {
        logSistema(`⏱️ Timeout dispatcher (${elapsed}ms) - Stop esecuzione`);
        break;
      }
      
      // Esegui job
      const risultato = processaJob(job, sheet, jobs);
      
      jobsEseguiti++;
      if (risultato.esito === "OK") jobsOk++;
      else if (risultato.esito === "ERRORE") jobsErrore++;
      else if (risultato.esito === "SKIP") jobsSkip++;
    }
    
    // 5. Ricalcola tutte le prossime esecuzioni
    ricalcolaTutteProssimeEsecuzioni(sheet);
    
    // 6. Report finale
    const durataTotale = Math.round((new Date() - startTime) / 1000);
    const report = `✅ Dispatcher completato in ${durataTotale}s | ` +
                   `Eseguiti: ${jobsEseguiti} | OK: ${jobsOk} | Errori: ${jobsErrore} | Skip: ${jobsSkip}`;
    
    logSistema(report);
    logSistema("═══ DISPATCHER END ═══");
    
    // Aggiorna SETUP
    aggiornaSetup(CONFIG.KEYS_SETUP.DISPATCHER_ULTIMO_RUN, new Date());
    aggiornaSetup(CONFIG.KEYS_SETUP.DISPATCHER_STATUS, "OK");
    
  } catch (error) {
    const errMsg = `❌ ERRORE CRITICO DISPATCHER: ${error.toString()}`;
    logSistema(errMsg);
    aggiornaSetup(CONFIG.KEYS_SETUP.DISPATCHER_STATUS, "ERRORE: " + error.toString());
    throw error;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CARICAMENTO JOBS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Carica tutti i job dal foglio AUTOMAZIONI
 * @param {Sheet} sheet - Foglio AUTOMAZIONI
 * @returns {Array} Array di oggetti job
 */
function caricaJobs(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const colMap = {};
  Object.entries(CONFIG.COLONNE_AUTOMAZIONI).forEach(([key, nome]) => {
    colMap[key] = headers.indexOf(nome);
  });
  
  return rows.map((row, idx) => ({
    rowNum: idx + 2, // +2 perché header=1 e array 0-based
    id: row[colMap.ID_JOB],
    attiva: row[colMap.ATTIVA],
    nome: row[colMap.NOME],
    categoria: row[colMap.CATEGORIA],
    codice: row[colMap.CODICE_FUNZIONE],
    frequenza: row[colMap.FREQUENZA],
    intervallo: row[colMap.INTERVALLO_ORE],
    ora: row[colMap.ORA_ESECUZIONE],
    dipendeDa: row[colMap.DIPENDE_DA],
    priorita: row[colMap.PRIORITA],
    timeoutSec: row[colMap.TIMEOUT_SEC] || CONFIG.DEFAULTS.TIMEOUT_SEC,
    maxRetry: row[colMap.MAX_RETRY] || CONFIG.DEFAULTS.MAX_RETRY,
    parametri: row[colMap.PARAMETRI_JSON],
    ultimaExec: row[colMap.ULTIMA_EXEC],
    prossimaExec: row[colMap.PROSSIMA_EXEC],
    esito: row[colMap.ESITO],
    durata: row[colMap.DURATA_SEC],
    erroreMsg: row[colMap.ERRORE_MSG],
    retryCount: row[colMap.RETRY_COUNT] || 0,
    execCount: row[colMap.EXEC_COUNT] || 0,
    successRate: row[colMap.SUCCESS_RATE],
    note: row[colMap.NOTE]
  })).filter(job => job.id); // Filtra righe vuote
}

// ═══════════════════════════════════════════════════════════════════════
// LOGICA SCHEDULING
// ═══════════════════════════════════════════════════════════════════════

/**
 * Determina se un job deve eseguire ORA
 * @param {Object} job - Oggetto job
 * @returns {Boolean}
 */
function deveEseguireOra(job) {
  const now = new Date();
  
  // Se ha PROSSIMA_EXEC, usa quella
  if (job.prossimaExec && job.prossimaExec instanceof Date) {
    return now >= job.prossimaExec;
  }
  
  // Altrimenti calcola al volo
  switch (job.frequenza) {
    case CONFIG.FREQUENZE.ORARIA:
      return deveEseguireOraria(job, now);
    
    case CONFIG.FREQUENZE.GIORNALIERA:
      return deveEseguireGiornaliera(job, now);
    
    case CONFIG.FREQUENZE.SETTIMANALE:
      return deveEseguireSettimanale(job, now);
    
    case CONFIG.FREQUENZE.CUSTOM:
      return deveEseguireCustom(job, now);
    
    default:
      logSistema(`⚠️ Frequenza sconosciuta per ${job.id}: ${job.frequenza}`);
      return false;
  }
}

/**
 * Check frequenza ORARIA
 */
function deveEseguireOraria(job, now) {
  if (!job.ultimaExec) return true; // Prima esecuzione
  
  const intervallo = parseInt(job.intervallo) || 1;
  const orePassed = (now - new Date(job.ultimaExec)) / (1000 * 60 * 60);
  
  return orePassed >= intervallo;
}

/**
 * Check frequenza GIORNALIERA
 */
function deveEseguireGiornaliera(job, now) {
  if (!job.ultimaExec) return true;
  
  const ultimaDate = new Date(job.ultimaExec);
  const sameDayUltima = now.toDateString() === ultimaDate.toDateString();
  
  // Se già eseguito oggi, skip
  if (sameDayUltima) return false;
  
  // Se ha ORA_ESECUZIONE specifica, controlla quella
  if (job.ora && job.ora !== "*") {
    const [targetOre, targetMin] = job.ora.split(":").map(Number);
    return now.getHours() >= targetOre && now.getMinutes() >= (targetMin || 0);
  }
  
  return true;
}

/**
 * Check frequenza SETTIMANALE
 */
function deveEseguireSettimanale(job, now) {
  const giorniConfig = (job.ora || "*").toString();
  if (giorniConfig === "*") return deveEseguireGiornaliera(job, now);
  
  const giorniTarget = giorniConfig.split(",").map(Number);
  const dayOfWeek = now.getDay();
  const dayConverted = dayOfWeek === 0 ? 7 : dayOfWeek;
  
  if (!giorniTarget.includes(dayConverted)) return false;
  
  if (job.ultimaExec) {
    const ultimaDate = new Date(job.ultimaExec);
    if (now.toDateString() === ultimaDate.toDateString()) return false;
  }
  
  return true;
}

/**
 * Check frequenza CUSTOM (usa PROSSIMA_EXEC obbligatoria)
 */
function deveEseguireCustom(job, now) {
  if (!job.prossimaExec) {
    logSistema(`⚠️ ${job.id}: CUSTOM richiede PROSSIMA_EXEC`);
    return false;
  }
  
  return now >= new Date(job.prossimaExec);
}

// ═══════════════════════════════════════════════════════════════════════
// ESECUZIONE JOB
// ═══════════════════════════════════════════════════════════════════════

/**
 * Processa un singolo job con controllo dipendenze, timeout, retry
 * @param {Object} job - Job da eseguire
 * @param {Sheet} sheet - Foglio AUTOMAZIONI
 * @param {Array} allJobs - Tutti i job (per check dipendenze)
 * @returns {Object} {esito, durata, errore}
 */
function processaJob(job, sheet, allJobs) {
  const startTime = new Date();
  
  logSistema(`🔵 START ${job.id}: ${job.nome}`);
  
  // 1. Controlla dipendenze
  if (job.dipendeDa) {
    const dipOk = verificaDipendenza(job, allJobs);
    if (!dipOk) {
      aggiornaJob(sheet, job.rowNum, {
        esito: CONFIG.ESITI.WAITING,
        errore: `In attesa di: ${job.dipendeDa}`
      });
      logSistema(`⏸️ ${job.id}: WAITING (dipendenza: ${job.dipendeDa})`);
      return { esito: CONFIG.ESITI.WAITING, durata: 0 };
    }
  }
  
  // 2. Segna IN_CORSO
  aggiornaJob(sheet, job.rowNum, { esito: CONFIG.ESITI.IN_CORSO });
  
  try {
    // 3. Parsing parametri JSON
    let parametri = {};
    if (job.parametri && job.parametri.trim()) {
      try {
        parametri = JSON.parse(job.parametri);
      } catch (e) {
        logSistema(`⚠️ ${job.id}: Parametri JSON invalidi, uso {}`);
      }
    }
    
    // 4. Esegui funzione con timeout
    const risultato = eseguiFunzioneConTimeout(job.codice, parametri, job.timeoutSec);
    
    const durata = Math.round((new Date() - startTime) / 1000);
    
    if (risultato.success) {
      // Successo
      const execCount = (job.execCount || 0) + 1;
      const okCount = Math.round((job.successRate || 0) * (execCount - 1) / 100) + 1;
      const successRate = Math.round((okCount / execCount) * 100);
      
      aggiornaJob(sheet, job.rowNum, {
        ultimaExec: new Date(),
        esito: CONFIG.ESITI.OK,
        durata: durata,
        errore: "",
        retryCount: 0,
        execCount: execCount,
        successRate: successRate
      });
      
      logSistema(`✅ ${job.id}: OK (${durata}s)`);
      return { esito: CONFIG.ESITI.OK, durata };
      
    } else if (risultato.timeout) {
      // Timeout
      aggiornaJob(sheet, job.rowNum, {
        ultimaExec: new Date(),
        esito: CONFIG.ESITI.TIMEOUT,
        durata: durata,
        errore: `Timeout dopo ${job.timeoutSec}s`,
        retryCount: (job.retryCount || 0) + 1
      });
      
      logSistema(`⏱️ ${job.id}: TIMEOUT (${durata}s)`);
      return { esito: CONFIG.ESITI.TIMEOUT, durata };
      
    } else {
      // Errore
      throw new Error(risultato.error);
    }
    
  } catch (error) {
    const durata = Math.round((new Date() - startTime) / 1000);
    const retryCount = (job.retryCount || 0) + 1;
    
    const shouldRetry = retryCount < job.maxRetry;
    const esito = shouldRetry ? CONFIG.ESITI.SKIP : CONFIG.ESITI.ERRORE;
    
    aggiornaJob(sheet, job.rowNum, {
      ultimaExec: new Date(),
      esito: esito,
      durata: durata,
      errore: error.toString().substring(0, 500),
      retryCount: retryCount
    });
    
    const msg = shouldRetry 
      ? `⚠️ ${job.id}: SKIP (retry ${retryCount}/${job.maxRetry}) - ${error.message}`
      : `❌ ${job.id}: ERRORE - ${error.message}`;
    
    logSistema(msg);
    
    return { esito, durata, errore: error.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════
// VERIFICA DIPENDENZE
// ═══════════════════════════════════════════════════════════════════════

/**
 * Verifica se le dipendenze del job sono soddisfatte
 * @param {Object} job - Job da verificare
 * @param {Array} allJobs - Tutti i job
 * @returns {Boolean}
 */
function verificaDipendenza(job, allJobs) {
  const dipIds = job.dipendeDa.split(",").map(id => id.trim());
  
  for (const dipId of dipIds) {
    const dipJob = allJobs.find(j => j.id === dipId);
    
    if (!dipJob) {
      logSistema(`⚠️ ${job.id}: Dipendenza ${dipId} non trovata`);
      return false;
    }
    
    if (dipJob.attiva !== "SI") {
      logSistema(`⚠️ ${job.id}: Dipendenza ${dipId} non attiva`);
      return false;
    }
    
    if (dipJob.esito !== CONFIG.ESITI.OK) {
      logSistema(`⚠️ ${job.id}: Dipendenza ${dipId} non OK (${dipJob.esito})`);
      return false;
    }
    
    if (!dipJob.ultimaExec) {
      logSistema(`⚠️ ${job.id}: Dipendenza ${dipId} mai eseguita`);
      return false;
    }
  }
  
  return true;
}

// ═══════════════════════════════════════════════════════════════════════
// ESECUZIONE FUNZIONI CON TIMEOUT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Esegue una funzione con timeout
 * @param {String} functionName - Nome funzione da eseguire
 * @param {Object} params - Parametri
 * @param {Number} timeoutSec - Timeout in secondi
 * @returns {Object} {success, timeout, error, result}
 */
function eseguiFunzioneConTimeout(functionName, params, timeoutSec) {
  const startTime = new Date();
  
  try {
    // Verifica esistenza funzione
    if (typeof this[functionName] !== 'function') {
      throw new Error(`Funzione '${functionName}' non trovata`);
    }
    
    // Esegui funzione
    const result = this[functionName](params);
    
    const elapsed = (new Date() - startTime) / 1000;
    
    // Check timeout post-esecuzione
    if (elapsed > timeoutSec) {
      return { success: false, timeout: true, error: `Timeout: ${elapsed}s > ${timeoutSec}s` };
    }
    
    return { success: true, result };
    
  } catch (error) {
    const elapsed = (new Date() - startTime) / 1000;
    
    if (elapsed > timeoutSec) {
      return { success: false, timeout: true, error: error.toString() };
    }
    
    return { success: false, error: error.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════════
// AGGIORNAMENTO JOB
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggiorna colonne specifiche di un job
 * @param {Sheet} sheet - Foglio AUTOMAZIONI
 * @param {Number} rowNum - Numero riga (1-based)
 * @param {Object} updates - {colName: valore}
 */
function aggiornaJob(sheet, rowNum, updates) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Object.entries(updates).forEach(([key, value]) => {
    const colName = CONFIG.COLONNE_AUTOMAZIONI[key.toUpperCase()] || key;
    const colIndex = headers.indexOf(colName) + 1;
    
    if (colIndex > 0) {
      sheet.getRange(rowNum, colIndex).setValue(value);
    } else {
      Logger.log(`⚠️ Colonna non trovata: ${colName}`);
    }
  });
}

// ═══════════════════════════════════════════════════════════════════════
// CALCOLO PROSSIME ESECUZIONI
// ═══════════════════════════════════════════════════════════════════════

/**
 * Ricalcola PROSSIMA_EXEC per tutti i job attivi
 * @param {Sheet} sheet - Foglio AUTOMAZIONI
 */
function ricalcolaTutteProssimeEsecuzioni(sheet) {
  const jobs = caricaJobs(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colProssima = headers.indexOf(CONFIG.COLONNE_AUTOMAZIONI.PROSSIMA_EXEC) + 1;
  
  if (colProssima === 0) {
    logSistema("⚠️ Colonna PROSSIMA_EXEC non trovata");
    return;
  }
  
  jobs.forEach(job => {
    if (job.attiva !== "SI") return;
    
    const prossimaExec = calcolaProssimaEsecuzione(job);
    
    if (prossimaExec) {
      sheet.getRange(job.rowNum, colProssima).setValue(prossimaExec);
    }
  });
  
  logSistema("🔄 Ricalcolate tutte le prossime esecuzioni");
}

/**
 * Calcola la prossima esecuzione per un job
 * @param {Object} job - Job
 * @returns {Date|null}
 */
function calcolaProssimaEsecuzione(job) {
  const now = new Date();
  
  switch (job.frequenza) {
    case CONFIG.FREQUENZE.ORARIA:
      const intervallo = parseInt(job.intervallo) || 1;
      return new Date(now.getTime() + intervallo * 60 * 60 * 1000);
    
    case CONFIG.FREQUENZE.GIORNALIERA:
      const domani = new Date(now);
      domani.setDate(domani.getDate() + 1);
      
      if (job.ora && job.ora !== "*") {
        const [ore, min] = job.ora.split(":").map(Number);
        domani.setHours(ore, min || 0, 0, 0);
      } else {
        domani.setHours(now.getHours(), now.getMinutes(), 0, 0);
      }
      
      return domani;
    
    case CONFIG.FREQUENZE.SETTIMANALE:
      const giorniConfig = (job.ora || "*").toString();
      if (giorniConfig === "*") {
        return calcolaProssimaEsecuzione({ ...job, frequenza: CONFIG.FREQUENZE.GIORNALIERA });
      }
      
      const giorniTarget = giorniConfig.split(",").map(Number);
      const prossimoGiorno = trovaProximoGiornoSettimana(now, giorniTarget);
      
      return prossimoGiorno;
    
    case CONFIG.FREQUENZE.CUSTOM:
      return job.prossimaExec ? new Date(job.prossimaExec) : null;
    
    default:
      return null;
  }
}

/**
 * Trova il prossimo giorno della settimana tra quelli target
 * @param {Date} from - Data di partenza
 * @param {Array} targetDays - Array di giorni (1-7, dove 1=Lun)
 * @returns {Date}
 */
function trovaProximoGiornoSettimana(from, targetDays) {
  const result = new Date(from);
  result.setHours(6, 0, 0, 0);
  
  const currentDay = result.getDay() === 0 ? 7 : result.getDay();
  
  for (let i = 1; i <= 7; i++) {
    const testDay = ((currentDay + i - 1) % 7) + 1;
    if (targetDays.includes(testDay)) {
      result.setDate(result.getDate() + i);
      return result;
    }
  }
  
  result.setDate(result.getDate() + 1);
  return result;
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI JOB IMPLEMENTATE
// ═══════════════════════════════════════════════════════════════════════

/**
 * JOB-001: Health Check Sistema
 */
function healthCheckSistema(params) {
  logSistema("🏥 Health Check avviato");
  
  const checks = {
    openai: !!PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"),
    claude: !!PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY"),
    setup: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETUP),
    automazioni: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.AUTOMAZIONI),
    logIn: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOG_IN)
  };
  
  const allOk = Object.values(checks).every(v => v);
  const report = Object.entries(checks).map(([k, v]) => `${v ? "✅" : "❌"} ${k}`).join(", ");
  
  logSistema(`Health Check: ${report}`);
  
  if (!allOk) {
    throw new Error("Health check fallito: " + report);
  }
  
  return { status: "OK", checks };
}

/**
 * JOB-005: Sync File Esterni
 * NOTA: JOB-002, JOB-003, JOB-004 sono implementati in Logic.gs
 */
function syncFileEsterni(params) {
  logSistema("🔄 Sync File Esterni - PLACEHOLDER");
  // TODO: Implementare sync MasterSku, Fornitori
  return { synced: 0, message: "Implementazione futura" };
}

// ═══════════════════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Configura trigger orario per dispatcher
 */
function configuraTriggerOrario() {
  const ui = SpreadsheetApp.getUi();
  
  rimuoviTriggerDispatcher();
  
  const intervallo = parseInt(getSetupValue(CONFIG.KEYS_SETUP.DISPATCHER_INTERVALLO_MIN, 60));
  const ore = Math.floor(intervallo / 60);
  
  if (ore >= 1) {
    ScriptApp.newTrigger('eseguiDispatcher')
      .timeBased()
      .everyHours(ore)
      .create();
  } else {
    ScriptApp.newTrigger('eseguiDispatcher')
      .timeBased()
      .everyMinutes(intervallo)
      .create();
  }
  
  ui.alert(
    '✅ Trigger Configurato!',
    `Dispatcher eseguirà ogni ${intervallo} minuti.\n\n` +
    'Controlla: Estensioni > Apps Script > Trigger',
    ui.ButtonSet.OK
  );
  
  logSistema(`Trigger dispatcher configurato: ogni ${intervallo} min`);
}

/**
 * Rimuovi trigger dispatcher
 */
function rimuoviTriggerDispatcher() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'eseguiDispatcher') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Rimuovi TUTTI i trigger
 */
function rimuoviTuttiTrigger() {
  const ui = SpreadsheetApp.getUi();
  const risposta = ui.alert(
    '⚠️ Conferma',
    'Rimuovere TUTTI i trigger?',
    ui.ButtonSet.YES_NO
  );
  
  if (risposta === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    ui.alert(`✅ Rimossi ${triggers.length} trigger`);
    logSistema(`Rimossi tutti i trigger (${triggers.length})`);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════

function getSetupValue(key, defaultValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SETUP);
    if (!sheet) return defaultValue;
    
    const data = sheet.getDataRange().getValues();
    for (let row of data) {
      if (row[0] === key) {
        return row[1] !== "" ? row[1] : defaultValue;
      }
    }
    return defaultValue;
  } catch (e) {
    return defaultValue;
  }
}

function aggiornaSetup(key, value) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SETUP);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
    sheet.appendRow([key, value]);
  } catch (e) {
    Logger.log(`Errore aggiornaSetup: ${e}`);
  }
}

function logSistema(messaggio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LOG_SISTEMA);
    if (sheet) {
      sheet.appendRow([new Date(), messaggio]);
    }
  } catch(e) {
    Logger.log("Log: " + messaggio);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// FUNZIONI TEST/DEBUG
// ═══════════════════════════════════════════════════════════════════════

/**
 * Test manuale ricalcolo prossime esecuzioni
 */
function ricalcolaProssimeEsecuzioni() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.AUTOMAZIONI);
  
  ricalcolaTutteProssimeEsecuzioni(sheet);
  
  SpreadsheetApp.getUi().alert(
    '✅ Ricalcolo Completato',
    'Controlla colonna PROSSIMA_EXEC nel foglio AUTOMAZIONI',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Test singola automazione
 */
function testSingolaAutomazione() {
  const ui = SpreadsheetApp.getUi();
  const risposta = ui.prompt(
    'Test Automazione',
    'Inserisci ID job (es. JOB-001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (risposta.getSelectedButton() !== ui.Button.OK) return;
  
  const jobId = risposta.getResponseText().trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.AUTOMAZIONI);
  const jobs = caricaJobs(sheet);
  const job = jobs.find(j => j.id === jobId);
  
  if (!job) {
    ui.alert('❌ Job non trovato: ' + jobId);
    return;
  }
  
  ui.alert('🧪 Test in corso...', 'Esecuzione job: ' + job.nome, ui.ButtonSet.OK);
  
  const risultato = processaJob(job, sheet, jobs);
  
  const msg = `Esito: ${risultato.esito}\n` +
              `Durata: ${risultato.durata}s\n` +
              (risultato.errore ? `Errore: ${risultato.errore}` : '');
  
  ui.alert('Test Completato', msg, ui.ButtonSet.OK);
}