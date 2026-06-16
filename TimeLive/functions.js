// ============================================================
// ZAZA TIMBER — Timestamp Add-in v1.4
// Shared Runtime — Pilnīgi automātiska fona darbība
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

// ============================================================
// INICIALIZĀCIJA — izsaucas automātiski
// ============================================================

Office.onReady(async (info) => {
  console.log("ZAZA: Office.onReady fired.");
  if (info.host === Office.HostType.Excel) {

    // GALVENAIS — liek Excel VIENMĒR ielādēt add-in fonā
    // bez lietotāja darbībām katru reizi atverot Excel
    try {
      await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
      console.log("ZAZA: ✅ AutoStart ieslēgts — turpmāk darbojas automātiski.");
    } catch (e) {
      console.warn("ZAZA: setStartupBehavior nav atbalstīts šajā vidē:", e.message);
    }

    await registerGlobalListener();
  }
});

// WorkbookActivated papildu drošībai (Desktop)
async function onWorkbookOpen(event) {
  console.log("ZAZA: WorkbookActivated.");
  await registerGlobalListener();
  if (event && event.completed) event.completed();
}

// ============================================================
// REĢISTRĒ GLOBĀLO KLAUSĪTĀJU
// ============================================================

async function registerGlobalListener() {
  if (listenerRegistered) {
    console.log("ZAZA: Klausītājs jau aktīvs.");
    return;
  }
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
// GALVENĀ LOĢIKA
// ============================================================

async function handleChange(event) {
  if (isProcessing) return;
  if (event.changeType !== "RangeEdited") return;

  let address = event.address;
  if (address.includes("!")) address = address.split("!")[1];
  address = address.replace(/\$/g, "");

  const firstCell = address.split(":")[0];
  const colLetter = firstCell.replace(/[0-9]/g, "").toUpperCase();

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

  }).catch((e) => {
    isProcessing = false;
    console.error("ZAZA: handleChange kļūda:", e);
  });
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

  isProcessing = true;
  try {
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
  } finally {
    isProcessing = false;
  }
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

  isProcessing = true;
  try {
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
  } finally {
    isProcessing = false;
  }
}

// ============================================================
// STATUS POGA
// ============================================================

function showStatus(event) {
  console.log("ZAZA Timestamp v1.4 — Klausītājs: " + (listenerRegistered ? "✅ aktīvs" : "❌ neaktīvs"));
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
