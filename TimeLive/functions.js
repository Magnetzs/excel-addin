// ============================================================
// ZAZA TIMBER — Timestamp Add-in (Shared Runtime)
// Fona darbība — bez UI
// ============================================================

// Lapu konfigurācija: trigera kolonna AJ ir fiksēta visām lapām
// Katrai lapai atšķirīgas mērķa kolonnas datumam un laikam
const SHEET_CONFIG = {
  "AUDZESANA":                   { dateCol: "AX", timeCol: "AY" },
  "ČETRPUSĪGĀ_ĒVELE_LĪMĒŠANA":  { dateCol: "AZ", timeCol: "BA" },
  "BIEZUMĒVELE":                 { dateCol: "AW", timeCol: "AX" },
  "CNC":                         { dateCol: "BC", timeCol: "BD" }
};

const TRIGGER_COL = "AJ";
let isProcessing = false; // Infinite loop aizsardzība

// ============================================================
// AUTO-START — aktivizējas pēc dokumenta atvēršanas
// ============================================================

Office.onReady(async () => {
  await registerGlobalListener();
});

// WorkbookActivated event (Shared Runtime)
async function onWorkbookActivated(event) {
  await registerGlobalListener();
  event.completed();
}

// ============================================================
// REĢISTRĒ GLOBĀLO KLAUSĪTĀJU VISAI DARBGRĀMATAI
// ============================================================

async function registerGlobalListener() {
  try {
    await Excel.run(async (context) => {
      // Globāls klausītājs — visas lapas vienlaicīgi
      context.workbook.worksheets.onChanged.add(handleChange);
      await context.sync();
      console.log("ZAZA Timestamp: globālais klausītājs aktīvs.");
    });
  } catch (e) {
    console.error("ZAZA Timestamp: kļūda reģistrējot klausītāju:", e);
  }
}

// ============================================================
// GALVENĀ LOĢIKA
// ============================================================

async function handleChange(event) {
  // --- INFINITE LOOP AIZSARDZĪBA ---
  // Ja kods pats raksta šūnās, tas izsauc jaunu onChanged.
  // isProcessing flag nodrošina ka mēs ignorējam savus pašu rakstus.
  if (isProcessing) return;

  // Pieņem tikai tiešu lietotāja ievadi
  if (event.changeType !== "RangeEdited") return;

  // Notīra adresi
  let address = event.address;
  if (address.includes("!")) address = address.split("!")[1];
  address = address.replace(/\$/g, "");

  // Pārbauda vai izmaiņa ir tieši AJ kolonnā
  const firstCell  = address.split(":")[0];
  const colLetter  = firstCell.replace(/[0-9]/g, "").toUpperCase();
  if (colLetter !== TRIGGER_COL) return;

  await Excel.run(async (context) => {
    // Iegūst lapas nosaukumu
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();

    const sheetName = sheet.name.trim();
    const config    = SHEET_CONFIG[sheetName];

    // Ignorē lapas kas nav konfigurācijā
    if (!config) return;

    // Ielādē mainītā diapazona vērtības
    const changedRange = sheet.getRange(address);
    changedRange.load(["values", "rowIndex", "rowCount"]);
    await context.sync();

    // Savāc rindas kur vērtība ir tieši "+"
    const rowsToProcess = [];
    for (let i = 0; i < changedRange.rowCount; i++) {
      const raw = changedRange.values[i][0];
      const val = String(raw === null || raw === undefined ? "" : raw).trim();
      if (val === "+") {
        rowsToProcess.push(changedRange.rowIndex + i + 1);
      }
    }

    if (rowsToProcess.length === 0) return;

    // Iegūst pašreizējo datumu un laiku KĀ TEKSTU
    const now         = new Date();
    const dateText    = formatDate(now);  // DD.MM.YYYY
    const timeText    = formatTime(now);  // HH:MM:SS

    // --- RAKSTA VĒRTĪBAS (aizsargāts bloks) ---
    isProcessing = true;
    try {
      for (const rowNum of rowsToProcess) {
        // Pārbauda vai mērķa šūnas jau ir aizpildītas
        const dateCell = sheet.getRange(config.dateCol + rowNum);
        const timeCell = sheet.getRange(config.timeCol + rowNum);
        dateCell.load("values");
        timeCell.load("values");
        await context.sync();

        const existingDate = String(dateCell.values[0][0]).trim();
        // Ja datums jau ir — nelabojam (aizsardzība pret pārrakstīšanu)
        if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") {
          console.log("ZAZA: Datums jau eksistē rindā " + rowNum + ", izlaižam.");
          continue;
        }

        // Ieraksta datumu un laiku kā tekstu
        dateCell.values = [[dateText]];
        timeCell.values = [[timeText]];

        console.log("ZAZA: ✅ " + sheetName + " rinda " + rowNum +
          " → " + config.dateCol + "=" + dateText +
          ", " + config.timeCol + "=" + timeText);
      }
      await context.sync();
    } finally {
      // Vienmēr atbrīvo flag — pat ja kļūda
      isProcessing = false;
    }
  }).catch((e) => {
    isProcessing = false;
    console.error("ZAZA handleChange kļūda:", e);
  });
}

// ============================================================
// DATUMA UN LAIKA FORMATĒŠANA KĀ TEKSTS
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

// ============================================================
// STATUS POGA (redzama ribbon josla — pēc vajadzības)
// ============================================================

async function showStatus(event) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    const sheetName = sheet.name.trim();
    const config    = SHEET_CONFIG[sheetName];
    const msg = config
      ? "✅ ZAZA Timestamp aktīvs\nLapa: " + sheetName + "\nDatums → " + config.dateCol + " | Laiks → " + config.timeCol
      : "⚠️ Lapa '" + sheetName + "' nav konfigurēta.\nKonfigurētās lapas: " + Object.keys(SHEET_CONFIG).join(", ");
    console.log(msg);
  });
  event.completed();
}
