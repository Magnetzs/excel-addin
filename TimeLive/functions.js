// ============================================================
// ZAZA TIMBER — Timestamp Add-in v2.0
// Shared Runtime — Pilnīgi automātiska fona darbība
//
// IZMAIŅU ŽURNĀLS:
// v1.5: notikumu rinda (queue) — nepazaudē + ja notiek vienlaicīgi
// v1.6: lapa tiek ņemta no event.worksheetId, watchdog, timeout,
//       periodiska klausītāja "veselības" pārbaude
// v1.7: drošības tīkls — periodiska pilna skenēšana (60s)
// v2.0: batch context.sync() (1 sync uz N rindām), getUsedRange()
//       nevis fiksēts 20000, retry limit ar MAX_RETRIES, queue size
//       limit, drošāka worksheetId apstrāde ar fallback, bez
//       rekursijas processQueue, batch sweep rakstīšana,
//       centralizēta konfigurācija, detalizēta diagnostika.
//       Datums/laiks PALIEK kā teksts (DD.MM.YYYY / HH:MM:SS) —
//       tas ir oriģinālās specifikācijas prasība, netiek mainīts.
// ============================================================

// ============================================================
// KONFIGURĀCIJA
// ============================================================

const CONFIG = {
  SWEEP_INTERVAL_MS: 60000,       // pilna drošības skenēšana
  WATCHDOG_TIMEOUT_MS: 15000,     // piespiedu atbloķēšana, ja apstrāde iesprūst
  HEALTH_CHECK_MS: 30000,         // klausītāja "veselības" pārbaude
  EVENT_TIMEOUT_MS: 10000,        // max laiks vienam Excel.run izsaukumam
  MAX_QUEUE_SIZE: 5000,           // aizsardzība pret nekontrolētu atmiņas pieaugumu
  MAX_RETRIES: 5,                 // max atkārtojumu skaits vienam notikumam
  RETRY_DELAY_MS: 800
};

const SINGLE_TRIGGER_CONFIG = {
  "AUDZESANA":                  { triggerCol: "AJ", dateCol: "AX", timeCol: "AY" },
  "ČETRPUSĪGĀ_ĒVELE_LĪMĒŠANA": { triggerCol: "AJ", dateCol: "AZ", timeCol: "BA" },
  "BIEZUMĒVELE":                { triggerCol: "AJ", dateCol: "AW", timeCol: "AX" },
  "CNC":                        { triggerCol: "AL", dateCol: "BC", timeCol: "BD" }
};

const MULTI_TRIGGER_CONFIG = {
  "Bloki": { triggerCols: ["K", "L"], dateCol: "M", timeCol: "N" }
};

// ============================================================
// GLOBĀLAIS STĀVOKLIS + DIAGNOSTIKA
// ============================================================

let isProcessing = false;
let listenerRegistered = false;
let eventQueue = [];
let processingStartedAt = null;

const stats = {
  startedAt: Date.now(),
  eventsProcessed: 0,
  eventsFixed: 0,        // cik rindās faktiski ierakstīts datums/laiks
  retriesTotal: 0,
  eventsDropped: 0,      // pēc MAX_RETRIES sasniegšanas
  lastSweepAt: null,
  lastSweepFixed: 0,
  lastSuccessAt: null
};

// ============================================================
// WATCHDOG — pret iesprūdušu isProcessing
// ============================================================

setInterval(() => {
  if (isProcessing && processingStartedAt && (Date.now() - processingStartedAt > CONFIG.WATCHDOG_TIMEOUT_MS)) {
    console.warn("ZAZA: ⚠️ Watchdog — apstrāde iesprūdusi >" + (CONFIG.WATCHDOG_TIMEOUT_MS / 1000) + "s, piespiedu atbloķēšana. Rindā: " + eventQueue.length);
    isProcessing = false;
    processingStartedAt = null;
    if (eventQueue.length > 0) processQueue();
  }
}, 3000);

// ============================================================
// INICIALIZĀCIJA
// ============================================================

Office.onReady(async (info) => {
  console.log("ZAZA: Office.onReady fired.");
  if (info.host === Office.HostType.Excel) {
    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
      console.log("ZAZA: ✅ AutoStart ieslēgts.");
    } catch (e) {
      console.warn("ZAZA: setStartupBehavior nav atbalstīts:", e.message);
    }
    await registerGlobalListener();
    // Uzreiz pēc atvēršanas — pilna skenēšana, lai noķertu
    // jau pazudušos gadījumus no iepriekšējās sesijas.
    runFullSweep().catch((e) => console.error("ZAZA: sākotnējā sweep kļūda:", e));
  }
});

async function onWorkbookOpen(event) {
  console.log("ZAZA: WorkbookActivated.");
  await registerGlobalListener();
  if (event && event.completed) event.completed();
}

// Periodiska "veselības" pārbaude — atjauno klausītāju, ja sesija "nogalināta"
setInterval(async () => {
  try {
    await Excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
    });
  } catch (e) {
    console.warn("ZAZA: veselības pārbaude neizdevās, reģistrē klausītāju no jauna:", e.message);
    listenerRegistered = false;
    await registerGlobalListener();
  }
}, CONFIG.HEALTH_CHECK_MS);

// ============================================================
// REĢISTRĒ GLOBĀLO KLAUSĪTĀJU
// ============================================================

async function registerGlobalListener() {
  if (listenerRegistered) return;
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onChanged.add(handleChange);
      await context.sync();
      listenerRegistered = true;
      console.log("ZAZA: ✅ Globālais klausītājs reģistrēts.");
    });
  } catch (e) {
    console.error("ZAZA: Kļūda reģistrējot klausītāju:", e);
  }
}

// ============================================================
// NOTIKUMU UZTVERŠANA — ar queue size limitu
// ============================================================

function handleChange(event) {
  if (event.changeType !== "RangeEdited") return;

  if (eventQueue.length >= CONFIG.MAX_QUEUE_SIZE) {
    console.warn("ZAZA: ⚠️ eventQueue sasniedzis MAX_QUEUE_SIZE (" + CONFIG.MAX_QUEUE_SIZE + "), notikums NETIEK pievienots: " + event.address);
    stats.eventsDropped++;
    return;
  }

  event._retryCount = 0;
  eventQueue.push(event);

  if (isProcessing) {
    console.log("ZAZA: Notikums pievienots rindai (apstrāde jau notiek). Rindā: " + eventQueue.length);
    return;
  }

  processQueue();
}

// ============================================================
// APSTRĀDĀ RINDU — VIENS CIKLS, BEZ REKURSIJAS
// ============================================================

async function processQueue() {
  if (isProcessing) return;
  isProcessing = true;
  processingStartedAt = Date.now();

  try {
    // `while` cikls, kas turpina, kamēr rindā ir notikumi —
    // tas ietver arī notikumus, kas pienāk APSTRĀDES LAIKĀ,
    // tāpēc rekursīvs processQueue() izsaukums beigās nav nepieciešams.
    while (eventQueue.length > 0) {
      const event = eventQueue.shift();
      await handleSingleEvent(event);
    }
  } catch (e) {
    console.error("ZAZA: processQueue kļūda:", e);
  } finally {
    isProcessing = false;
    processingStartedAt = null;
  }
}

// ============================================================
// APSTRĀDĀ VIENU NOTIKUMU — ar retry limitu un drošu worksheetId
// ============================================================

async function handleSingleEvent(event) {
  let address = event.address;
  if (address.includes("!")) address = address.split("!")[1];
  address = address.replace(/\$/g, "");

  const firstCell = address.split(":")[0];
  const colLetter = firstCell.replace(/[0-9]/g, "").toUpperCase();

  try {
    await Promise.race([
      Excel.run(async (context) => {
        const sheet = await resolveSheet(context, event);
        if (!sheet) return; // lapa nav atrasta — izlaiž droši

        sheet.load("name");
        await context.sync();
        const sheetName = sheet.name.trim();

        const singleConfig = SINGLE_TRIGGER_CONFIG[sheetName];
        if (singleConfig && colLetter === singleConfig.triggerCol) {
          await processSingleTrigger(context, sheet, address, singleConfig);
          return;
        }

        const multiConfig = MULTI_TRIGGER_CONFIG[sheetName];
        if (multiConfig && multiConfig.triggerCols.includes(colLetter)) {
          await processMultiTrigger(context, sheet, address, colLetter, multiConfig);
          return;
        }
      }),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error("ZAZA timeout — Excel.run nepabeidzās " + (CONFIG.EVENT_TIMEOUT_MS / 1000) + "s laikā")), CONFIG.EVENT_TIMEOUT_MS)
      )
    ]);

    stats.eventsProcessed++;
    stats.lastSuccessAt = Date.now();

  } catch (e) {
    console.error("ZAZA: handleSingleEvent kļūda (adrese " + address + "):", e.message || e);

    event._retryCount = (event._retryCount || 0) + 1;
    stats.retriesTotal++;

    if (event._retryCount >= CONFIG.MAX_RETRIES) {
      console.error("ZAZA: ❌ Notikums " + address + " sasniedzis MAX_RETRIES (" + CONFIG.MAX_RETRIES + "), tiek izmests no rindas.");
      stats.eventsDropped++;
      return;
    }

    // Mēģina vēlreiz pēc brīža, ja vieta rindā atļauj
    setTimeout(() => {
      if (eventQueue.length >= CONFIG.MAX_QUEUE_SIZE) {
        console.warn("ZAZA: rinda pilna, retry notikums " + address + " tiek izmests.");
        stats.eventsDropped++;
        return;
      }
      eventQueue.push(event);
      if (!isProcessing) processQueue();
    }, CONFIG.RETRY_DELAY_MS);
  }
}

// ============================================================
// DROŠA LAPAS NOSKAIDROŠANA — ar fallback mehānismu
// ============================================================

async function resolveSheet(context, event) {
  // 1. mēģinājums — lapa pēc worksheetId no paša notikuma (precīzākais ceļš)
  if (event && event.worksheetId) {
    try {
      const sheet = context.workbook.worksheets.getItemOrNullObject(event.worksheetId);
      sheet.load("isNullObject,name");
      await context.sync();
      if (!sheet.isNullObject) {
        return sheet;
      }
      console.warn("ZAZA: worksheetId '" + event.worksheetId + "' neeksistē (lapa varētu būt dzēsta/pārvietota), fallback uz aktīvo lapu.");
    } catch (e) {
      console.warn("ZAZA: kļūda iegūstot lapu pēc worksheetId, fallback uz aktīvo lapu:", e.message);
    }
  }

  // 2. mēģinājums — fallback uz aktīvo lapu
  try {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();
    return activeSheet;
  } catch (e) {
    console.error("ZAZA: neizdevās iegūt arī aktīvo lapu — izlaiž notikumu droši:", e.message);
    return null;
  }
}

// ============================================================
// SINGLE TRIGGER apstrāde — BATCH (1 sync uz daudzām rindām)
// ============================================================

async function processSingleTrigger(context, sheet, address, config) {
  const changedRange = sheet.getRange(address);
  changedRange.load(["values", "rowIndex", "rowCount"]);
  await context.sync();

  const rowsToProcess = [];
  for (let i = 0; i < changedRange.rowCount; i++) {
    const val = String(changedRange.values[i][0] ?? "").trim();
    if (val === "+") rowsToProcess.push(changedRange.rowIndex + i + 1);
  }
  if (rowsToProcess.length === 0) return;

  // --- 1. PIEGĀJIENS: ielādē VISAS mērķa šūnas vienā sync ---
  const dateCells = rowsToProcess.map((rowNum) => {
    const c = sheet.getRange(config.dateCol + rowNum);
    c.load("values");
    return c;
  });
  await context.sync();

  // --- Aprēķina datumu/laiku, salīdzina atmiņā, sagatavo rakstāmos ---
  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  const rowsToWrite = [];
  for (let i = 0; i < rowsToProcess.length; i++) {
    const existing = String(dateCells[i].values[0][0] ?? "").trim();
    if (existing !== "" && existing !== "0" && existing !== "false") continue; // jau aizpildīts
    rowsToWrite.push(rowsToProcess[i]);
  }
  if (rowsToWrite.length === 0) return;

  // --- 2. PIEGĀJIENS: raksta VISAS vērtības, VIENS sync ---
  for (const rowNum of rowsToWrite) {
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
  }
  await context.sync();

  stats.eventsFixed += rowsToWrite.length;
  console.log("ZAZA: ✅ " + sheet.name + " — aizpildītas " + rowsToWrite.length + " rinda(s): " + rowsToWrite.join(", ") + " | " + config.dateCol + "=" + dateText + " | " + config.timeCol + "=" + timeText);
}

// ============================================================
// MULTI TRIGGER apstrāde (Bloki) — BATCH
// ============================================================

async function processMultiTrigger(context, sheet, address, changedCol, config) {
  const changedRange = sheet.getRange(address);
  changedRange.load(["values", "rowIndex", "rowCount"]);
  await context.sync();

  const rowsToProcess = [];
  for (let i = 0; i < changedRange.rowCount; i++) {
    const val = String(changedRange.values[i][0] ?? "").trim();
    if (val === "+") rowsToProcess.push(changedRange.rowIndex + i + 1);
  }
  if (rowsToProcess.length === 0) return;

  const dateCells = rowsToProcess.map((rowNum) => {
    const c = sheet.getRange(config.dateCol + rowNum);
    c.load("values");
    return c;
  });
  await context.sync();

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  const rowsToWrite = [];
  for (let i = 0; i < rowsToProcess.length; i++) {
    const existing = String(dateCells[i].values[0][0] ?? "").trim();
    if (existing !== "" && existing !== "0" && existing !== "false") continue;
    rowsToWrite.push(rowsToProcess[i]);
  }
  if (rowsToWrite.length === 0) return;

  for (const rowNum of rowsToWrite) {
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
  }
  await context.sync();

  stats.eventsFixed += rowsToWrite.length;
  console.log("ZAZA: ✅ Bloki — aizpildītas " + rowsToWrite.length + " rinda(s) (kol. " + changedCol + "): " + rowsToWrite.join(", "));
}

// ============================================================
// DROŠĪBAS TĪKLS — PERIODISKA PILNA SKENĒŠANA (getUsedRange, batch)
// ============================================================

setInterval(() => {
  runFullSweep().catch((e) => console.error("ZAZA: runFullSweep kļūda:", e));
}, CONFIG.SWEEP_INTERVAL_MS);

async function runFullSweep() {
  if (isProcessing) {
    console.log("ZAZA: Sweep — apstrāde jau notiek, nogaida nākamo reizi.");
    return;
  }

  let totalFixed = 0;

  try {
    await Excel.run(async (context) => {
      const allSheets = context.workbook.worksheets;
      allSheets.load("items/name");
      await context.sync();

      for (const sheet of allSheets.items) {
        const sheetName = sheet.name.trim();

        const singleConfig = SINGLE_TRIGGER_CONFIG[sheetName];
        if (singleConfig) {
          try {
            totalFixed += await sweepSingleTrigger(context, sheet, singleConfig);
          } catch (e) {
            // Viena lapa neizdodas (piem. protection) — turpina ar pārējām, nevis apstājas.
            console.error("ZAZA: sweepSingleTrigger kļūda lapā '" + sheetName + "', turpina ar pārējām lapām:", e.message || e);
          }
        }

        const multiConfig = MULTI_TRIGGER_CONFIG[sheetName];
        if (multiConfig) {
          try {
            totalFixed += await sweepMultiTrigger(context, sheet, multiConfig);
          } catch (e) {
            console.error("ZAZA: sweepMultiTrigger kļūda lapā '" + sheetName + "', turpina ar pārējām lapām:", e.message || e);
          }
        }
      }
    });
  } catch (e) {
    console.error("ZAZA: runFullSweep — Excel.run kļūda:", e.message || e);
  }

  stats.lastSweepAt = Date.now();
  stats.lastSweepFixed = totalFixed;
}

// Izmanto getUsedRange() — skenē tikai faktiski izmantoto apgabalu,
// nevis fiksētu 20000 rindu, un raksta visu BATCH veidā vienā sync.
async function sweepSingleTrigger(context, sheet, config) {
  let lastRow;
  try {
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("rowCount,isNullObject");
    await context.sync();
    lastRow = usedRange.isNullObject ? 0 : usedRange.rowCount;
  } catch (e) {
    console.warn("ZAZA: getUsedRange neizdevās lapā " + sheet.name + ", izlaiž šo sweep ciklu:", e.message);
    return 0;
  }

  if (lastRow <= 0) return 0; // tukša lapa — nav ko skenēt

  const triggerRange = sheet.getRange(config.triggerCol + "1:" + config.triggerCol + lastRow);
  const dateRange    = sheet.getRange(config.dateCol + "1:" + config.dateCol + lastRow);
  triggerRange.load("values");
  dateRange.load("values");
  await context.sync();

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  // Savāc rindas kuras jāaizpilda ATMIŅĀ, neveido Range objektus cikla iekšā
  const rowsToFix = [];
  for (let i = 0; i < lastRow; i++) {
    const triggerVal = String(triggerRange.values[i][0] ?? "").trim();
    if (triggerVal !== "+") continue;

    const existingDate = String(dateRange.values[i][0] ?? "").trim();
    if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") continue; // NEPĀRRAKSTA

    rowsToFix.push(i + 1);
  }

  if (rowsToFix.length === 0) return 0;

  // Raksta visu BATCH veidā, viens sync
  for (const rowNum of rowsToFix) {
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
  }
  await context.sync();

  console.log("ZAZA: 🛡️ Sweep " + sheet.name + " — izlaboti " + rowsToFix.length + " trūkstoši ieraksti.");
  return rowsToFix.length;
}

async function sweepMultiTrigger(context, sheet, config) {
  let lastRow;
  try {
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    usedRange.load("rowCount,isNullObject");
    await context.sync();
    lastRow = usedRange.isNullObject ? 0 : usedRange.rowCount;
  } catch (e) {
    console.warn("ZAZA: getUsedRange neizdevās lapā " + sheet.name + ", izlaiž šo sweep ciklu:", e.message);
    return 0;
  }

  if (lastRow <= 0) return 0; // tukša lapa — nav ko skenēt

  const dateRange = sheet.getRange(config.dateCol + "1:" + config.dateCol + lastRow);
  dateRange.load("values");

  const triggerRanges = config.triggerCols.map((col) => {
    const r = sheet.getRange(col + "1:" + col + lastRow);
    r.load("values");
    return r;
  });

  await context.sync();

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  const rowsToFix = [];
  for (let i = 0; i < lastRow; i++) {
    const hasPlus = triggerRanges.some((r) => String(r.values[i][0] ?? "").trim() === "+");
    if (!hasPlus) continue;

    const existingDate = String(dateRange.values[i][0] ?? "").trim();
    if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") continue; // NEPĀRRAKSTA

    rowsToFix.push(i + 1);
  }

  if (rowsToFix.length === 0) return 0;

  for (const rowNum of rowsToFix) {
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
  }
  await context.sync();

  console.log("ZAZA: 🛡️ Sweep " + sheet.name + " — izlaboti " + rowsToFix.length + " trūkstoši ieraksti.");
  return rowsToFix.length;
}

// ============================================================
// STATUS / DIAGNOSTIKA
// ============================================================

function showStatus(event) {
  const uptimeMin = Math.round((Date.now() - stats.startedAt) / 60000);
  const lastSweep = stats.lastSweepAt ? Math.round((Date.now() - stats.lastSweepAt) / 1000) + "s atpakaļ" : "vēl nav notikusi";
  const lastSuccess = stats.lastSuccessAt ? Math.round((Date.now() - stats.lastSuccessAt) / 1000) + "s atpakaļ" : "—";

  console.log(
    "ZAZA Timestamp v2.0\n" +
    "  Darbojas: " + uptimeMin + " min\n" +
    "  Klausītājs: " + (listenerRegistered ? "✅ aktīvs" : "❌ neaktīvs") + "\n" +
    "  Rindā: " + eventQueue.length + " | Apstrādā tagad: " + (isProcessing ? "jā" : "nē") + "\n" +
    "  Apstrādāti notikumi: " + stats.eventsProcessed + " | Izlaboti ieraksti: " + stats.eventsFixed + "\n" +
    "  Atkārtojumi (retry): " + stats.retriesTotal + " | Izmesti (max retries): " + stats.eventsDropped + "\n" +
    "  Pēdējā sweep: " + lastSweep + " (izlabots: " + stats.lastSweepFixed + ")\n" +
    "  Pēdējā veiksmīgā apstrāde: " + lastSuccess
  );
  if (event && event.completed) event.completed();
}

// ============================================================
// FORMATĒŠANA — datums/laiks kā TEKSTS (specifikācijas prasība)
// ============================================================

function formatDate(date) {
  return String(date.getDate()).padStart(2, "0") + "." +
         String(date.getMonth() + 1).padStart(2, "0") + "." +
         date.getFullYear();
}

function formatTime(date) {
  return String(date.getHours()).padStart(2, "0") + ":" +
         String(date.getMinutes()).padStart(2, "0") + ":" +
         String(date.getSeconds()).padStart(2, "0");
}
