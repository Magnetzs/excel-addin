// ============================================================
// ZAZA TIMBER — Timestamp Add-in v1.2
// Shared Runtime — fona darbība
// ============================================================

// Konfigurācija lapām ar vienu trigera kolonnu (AJ)
const SINGLE_TRIGGER_CONFIG = {
  "AUDZESANA":                  { triggerCol: "AJ", dateCol: "AX", timeCol: "AY" },
  "ČETRPUSĪGĀ_ĒVELE_LĪMĒŠANA": { triggerCol: "AJ", dateCol: "AZ", timeCol: "BA" },
  "BIEZUMĒVELE":                { triggerCol: "AJ", dateCol: "AW", timeCol: "AX" },
  "CNC":                        { triggerCol: "AL", dateCol: "BC", timeCol: "BD" }
};

// Konfigurācija lapām ar vairākām trigera kolonnām
// Ja JEBKURA no triggerCols satur "+", ieraksta dateCol un timeCol
// Pārrakstīšana NAV atļauta — pirmais laiks paliek
const MULTI_TRIGGER_CONFIG = {
  "Bloki": { triggerCols: ["K", "L"], dateCol: "M", timeCol: "N" }
};

let isProcessing = false;
let listenerRegistered = false;

// ============================================================
// INICIALIZĀCIJA
// ============================================================

Office.onReady(async (info) => {
  console.log("ZAZA: Office.onReady - host:", info.host);
  if (info.host === Office.HostType.Excel) {
    await registerGlobalListener();
  }
});

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
      console.log("ZAZA: ✅ Globālais klausītājs aktīvs.");
    });
  } catch (e) {
    console.error("ZAZA: Kļūda reģistrējot klausītāju:", e);
  }
}

// ============================================================
// GALVENĀ LOĢIKA — NOSAKA KURA LAPA UN KONFIGURĀCIJA
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

    // Pārbauda single trigger lapas (AJ kolonna)
    const singleConfig = SINGLE_TRIGGER_CONFIG[sheetName];
    if (singleConfig && colLetter === singleConfig.triggerCol) {
      await processSingleTrigger(context, sheet, address, singleConfig);
      return;
    }

    // Pārbauda multi trigger lapas (K vai L kolonna)
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
// SINGLE TRIGGER — AJ kolonna (AUDZESANA, CNC, utt.)
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
      timeCell.load("values");
      await context.sync();

      const existing = String(dateCell.values[0][0] ?? "").trim();
      if (existing !== "" && existing !== "0" && existing !== "false") {
        console.log("ZAZA: Rinda " + rowNum + " — datums jau eksistē, izlaižam.");
        continue;
      }

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
// MULTI TRIGGER — K vai L kolonna (Bloki lapa)
// Loģika: ja M jau ir aizpildīts — NEPĀRRAKSTA (pirmais laiks paliek)
// ============================================================

async function processMultiTrigger(context, sheet, address, changedCol, config) {
  const changedRange = sheet.getRange(address);
  changedRange.load(["values", "rowIndex", "rowCount"]);
  await context.sync();

  // Savāc rindas kur izmainītajā kolonnā ir "+"
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
      timeCell.load("values");
      await context.sync();

      // Ja M kolonnā jau ir datums — NEPĀRRAKSTA (pirmais + laiks paliek)
      const existingDate = String(dateCell.values[0][0] ?? "").trim();
      if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") {
        console.log("ZAZA: Bloki rinda " + rowNum + " — datums jau eksistē (pirmais laiks saglabāts), izlaižam.");
        continue;
      }

      // Ieraksta datumu un laiku
      dateCell.values = [[dateText]];
      timeCell.values = [[timeText]];
      console.log("ZAZA: ✅ Bloki rinda " + rowNum + " (trigera kol. " + changedCol + ") → M=" + dateText + " | N=" + timeText);
    }
    await context.sync();
  } finally {
    isProcessing = false;
  }
}

// ============================================================
// STATUS FUNKCIJA (ribbon poga)
// ============================================================

function showStatus(event) {
  console.log("ZAZA Timestamp aktīvs.");
  console.log("Single trigger lapas: " + Object.keys(SINGLE_TRIGGER_CONFIG).join(", "));
  console.log("Multi trigger lapas: " + Object.keys(MULTI_TRIGGER_CONFIG).join(", "));
  if (event && event.completed) event.completed();
}

// ============================================================
// DATUMA UN LAIKA FORMATĒŠANA
// ============================================================

function formatDate(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = date.getFullYear();
  return d + "." + m + "." + y;
}

function formatTime(date) {
  const h  = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const s  = String(date.getSeconds()).padStart(2, "0");
  return h + ":" + mi + ":" + s;
}
