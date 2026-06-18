// ============================================================
// ZAZA TIMBER — Timestamp Add-in v1.5
// Shared Runtime — Pilnīgi automātiska fona darbība
// FIX: notikumu rinda (queue) — nepazaudē + ja notiek vienlaicīgi
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

  try {
    while (eventQueue.length > 0) {
      const event = eventQueue.shift(); // ņem pirmo no rindas
      await handleSingleEvent(event);
    }
  } catch (e) {
    console.error("ZAZA: processQueue kļūda:", e);
  } finally {
    isProcessing = false;
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
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
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
    });
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
  console.log("ZAZA Timestamp v1.5 — Klausītājs: " + (listenerRegistered ? "✅ aktīvs" : "❌ neaktīvs") + " | Rindā: " + eventQueue.length);
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
