// ============================================================
// ZAZA TIMBER — Timestamp Add-in v1.7
// Shared Runtime — Pilnīgi automātiska fona darbība
// FIX v1.5: notikumu rinda (queue) — nepazaudē + ja notiek vienlaicīgi
// FIX v1.6: lapa tiek ņemta no event.worksheetId (nevis "aktīvā" lapa),
//           watchdog pret iesprūdušu isProcessing, timeout uz Excel.run,
//           periodiska klausītāja "veselības" pārbaude
// FIX v1.7: drošības tīkls — periodiska pilna skenēšana (60s) kas
//           noķer JEBKURU palaistu "+", nekad nepārraksta esošos datus
// ============================================================

const SINGLE_TRIGGER_CONFIG = {
  "AUDZESANA":                  { triggerCol: "AJ", dateCol: "AX", timeCol: "AY" },
  "ČETRPUSĪGĀ_ĒVELE_LĪMĒŠANA": { triggerCol: "AJ", dateCol: "AZ", timeCol: "BA" },
  "BIEZUMĒVELE":                { triggerCol: "AJ", dateCol: "AW", timeCol: "AX" },
  "CNC":                        { triggerCol: "AL", dateCol: "BC", timeCol: "BD" }
};

const MULTI_TRIGGER_CONFIG = {
  "Bloki": { triggerCols: ["K", "L"], dateCol: "M", timeCol: "N" }
};

let isProcessing = false;
let listenerRegistered = false;
let eventQueue = []; // notikumu rinda — nekas netiek pazaudēts
let processingStartedAt = null;
const WATCHDOG_TIMEOUT_MS = 15000; // ja apstrāde "iesprūst" >15s, atbloķē piespiedu kārtā

// Watchdog — nepārtraukti pārbauda vai isProcessing nav iesprūdis
setInterval(() => {
  if (isProcessing && processingStartedAt && (Date.now() - processingStartedAt > WATCHDOG_TIMEOUT_MS)) {
    console.warn("ZAZA: ⚠️ Watchdog — apstrāde iesprūdusi >15s, piespiedu atbloķēšana. Rindā: " + eventQueue.length);
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
  }
});

async function onWorkbookOpen(event) {
  console.log("ZAZA: WorkbookActivated.");
  await registerGlobalListener();
  if (event && event.completed) event.completed();
}

// Periodiska pārbaude — ja kāda autentifikācijas/token problēma "nogalina"
// klausītāju fonā, šis to konstatē un reģistrē no jauna.
setInterval(async () => {
  try {
    await Excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
    });
  } catch (e) {
    console.warn("ZAZA: periodiskā pārbaude neizdevās, mēģina reģistrēt klausītāju no jauna:", e.message);
    listenerRegistered = false;
    await registerGlobalListener();
  }
}, 30000);

// ============================================================
// DROŠĪBAS TĪKLS — PERIODISKA PILNA SKENĒŠANA
// Noķer JEBKURU "+" kam trūkst datuma/laika, neatkarīgi no tā,
// vai onChanged notikums tika pazaudēts. NEKAD nepārraksta esošos datus.
// ============================================================

const SWEEP_INTERVAL_MS = 60000; // 60 sekundes

setInterval(() => {
  runFullSweep().catch((e) => console.error("ZAZA: runFullSweep kļūda:", e));
}, SWEEP_INTERVAL_MS);

async function runFullSweep() {
  if (isProcessing) {
    console.log("ZAZA: Sweep — apstrāde jau notiek, nogaida.");
    return;
  }

  await Excel.run(async (context) => {
    const allSheets = context.workbook.worksheets;
    allSheets.load("items/name");
    await context.sync();

    for (const sheet of allSheets.items) {
      const sheetName = sheet.name.trim();

      const singleConfig = SINGLE_TRIGGER_CONFIG[sheetName];
      if (singleConfig) {
        await sweepSingleTrigger(context, sheet, singleConfig);
      }

      const multiConfig = MULTI_TRIGGER_CONFIG[sheetName];
      if (multiConfig) {
        await sweepMultiTrigger(context, sheet, multiConfig);
      }
    }
  });
}

// Skenē VISU trigera kolonnu (piem. AJ1:AJ20000) vienā piegājienā,
// salīdzina ar datuma kolonnu, un aizpilda TIKAI tās rindas kur trūkst.
async function sweepSingleTrigger(context, sheet, config) {
  const MAX_ROW = 20000;
  const triggerRange = sheet.getRange(config.triggerCol + "1:" + config.triggerCol + MAX_ROW);
  const dateRange    = sheet.getRange(config.dateCol + "1:" + config.dateCol + MAX_ROW);
  triggerRange.load("values");
  dateRange.load("values");
  await context.sync();

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  let fixedCount = 0;
  for (let i = 0; i < MAX_ROW; i++) {
    const triggerVal = String(triggerRange.values[i][0] ?? "").trim();
    if (triggerVal !== "+") continue;

    const existingDate = String(dateRange.values[i][0] ?? "").trim();
    if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") continue; // jau aizpildīts — NEPĀRRAKSTA

    const rowNum = i + 1;
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
    fixedCount++;
  }

  if (fixedCount > 0) {
    await context.sync();
    console.log("ZAZA: 🛡️ Sweep " + sheet.name + " — izlaboti " + fixedCount + " trūkstoši ieraksti.");
  }
}

async function sweepMultiTrigger(context, sheet, config) {
  const MAX_ROW = 20000;
  const dateRange = sheet.getRange(config.dateCol + "1:" + config.dateCol + MAX_ROW);
  dateRange.load("values");

  const triggerRanges = config.triggerCols.map((col) => {
    const r = sheet.getRange(col + "1:" + col + MAX_ROW);
    r.load("values");
    return { col, range: r };
  });

  await context.sync();

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  let fixedCount = 0;
  for (let i = 0; i < MAX_ROW; i++) {
    const hasPlus = triggerRanges.some(({ range }) => String(range.values[i][0] ?? "").trim() === "+");
    if (!hasPlus) continue;

    const existingDate = String(dateRange.values[i][0] ?? "").trim();
    if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") continue; // jau aizpildīts — NEPĀRRAKSTA

    const rowNum = i + 1;
    sheet.getRange(config.dateCol + rowNum).values = [[dateText]];
    sheet.getRange(config.timeCol + rowNum).values = [[timeText]];
    fixedCount++;
  }

  if (fixedCount > 0) {
    await context.sync();
    console.log("ZAZA: 🛡️ Sweep " + sheet.name + " — izlaboti " + fixedCount + " trūkstoši ieraksti.");
  }
}

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
// GALVENĀ LOĢIKA — NOTIKUMI VIENMĒR TIEK PIEVIENOTI RINDAI
// ============================================================

async function handleChange(event) {
  if (event.changeType !== "RangeEdited") return;

  // VISI notikumi tiek pievienoti rindai — nekas netiek pazaudēts
  eventQueue.push(event);

  // Ja jau notiek apstrāde — jaunais notikums sagaida savu kārtu
  if (isProcessing) {
    console.log("ZAZA: Notikums pievienots rindai (apstrāde jau notiek). Rindā: " + eventQueue.length);
    return;
  }

  await processQueue();
}

// ============================================================
// APSTRĀDĀ RINDU — VIENS NOTIKUMS PĒC OTRA, NEKAS NEPAZŪD
// ============================================================

async function processQueue() {
  if (isProcessing) return; // drošības pārbaude
  isProcessing = true;
  processingStartedAt = Date.now();

  try {
    while (eventQueue.length > 0) {
      const event = eventQueue.shift(); // ņem pirmo no rindas
      await handleSingleEvent(event);
    }
  } catch (e) {
    console.error("ZAZA: processQueue kļūda:", e);
  } finally {
    isProcessing = false;
    processingStartedAt = null;
  }

  // Ja apstrādes laikā pienāca jauni notikumi — apstrādā arī tos
  if (eventQueue.length > 0) {
    await processQueue();
  }
}

// ============================================================
// APSTRĀDĀ VIENU NOTIKUMU
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
        // KRITISKI: izmanto LAPU NO PAŠA NOTIKUMA (worksheetId), nevis "aktīvo" lapu.
        // Ja lietotājs pārslēdzas uz citu lapu kamēr šis notikums gaida rindā,
        // getActiveWorksheet() atgrieztu NEPAREIZO lapu un + paliktu neapstrādāts.
        const sheet = event.worksheetId
          ? context.workbook.worksheets.getItem(event.worksheetId)
          : context.workbook.worksheets.getActiveWorksheet();
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
        setTimeout(() => reject(new Error("ZAZA timeout — Excel.run nepabeidzās 10s laikā")), 10000)
      )
    ]);
  } catch (e) {
    console.error("ZAZA: handleSingleEvent kļūda (adrese " + address + "):", e);
    // Kļūdas gadījumā NEMET notikumu — mēģina vēlreiz pēc brīža
    setTimeout(() => {
      eventQueue.push(event);
      if (!isProcessing) processQueue();
    }, 800);
  }
}

// ============================================================
// SINGLE TRIGGER apstrāde
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

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  for (const rowNum of rowsToProcess) {
    const dateCell = sheet.getRange(config.dateCol + rowNum);
    const timeCell = sheet.getRange(config.timeCol + rowNum);
    dateCell.load("values");
    await context.sync();

    const existing = String(dateCell.values[0][0] ?? "").trim();
    if (existing !== "" && existing !== "0" && existing !== "false") continue;

    dateCell.values = [[dateText]];
    timeCell.values = [[timeText]];
    console.log("ZAZA: ✅ " + sheet.name + " rinda " + rowNum + " → " + config.dateCol + "=" + dateText + " | " + config.timeCol + "=" + timeText);
  }
  await context.sync();
}

// ============================================================
// MULTI TRIGGER apstrāde (Bloki lapa)
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

  const now      = new Date();
  const dateText = formatDate(now);
  const timeText = formatTime(now);

  for (const rowNum of rowsToProcess) {
    const dateCell = sheet.getRange(config.dateCol + rowNum);
    const timeCell = sheet.getRange(config.timeCol + rowNum);
    dateCell.load("values");
    await context.sync();

    const existing = String(dateCell.values[0][0] ?? "").trim();
    if (existing !== "" && existing !== "0" && existing !== "false") continue;

    dateCell.values = [[dateText]];
    timeCell.values = [[timeText]];
    console.log("ZAZA: ✅ Bloki rinda " + rowNum + " (kol. " + changedCol + ") → M=" + dateText + " | N=" + timeText);
  }
  await context.sync();
}

// ============================================================
// STATUS POGA
// ============================================================

function showStatus(event) {
  console.log("ZAZA Timestamp v1.7 — Klausītājs: " + (listenerRegistered ? "✅ aktīvs" : "❌ neaktīvs") +
    " | Rindā: " + eventQueue.length +
    " | Apstrādā: " + (isProcessing ? "jā" : "nē"));
  if (event && event.completed) event.completed();
}

// ============================================================
// FORMATĒŠANA
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
