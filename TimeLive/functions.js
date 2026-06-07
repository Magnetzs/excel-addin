// ============================================================
// ZAZA TIMBER — Timestamp Add-in v1.1
// Shared Runtime — fona darbība
// ============================================================

const SHEET_CONFIG = {
  "AUDZESANA":                  { dateCol: "AX", timeCol: "AY" },
  "ČETRPUSĪGĀ_ĒVELE_LĪMĒŠANA": { dateCol: "AZ", timeCol: "BA" },
  "BIEZUMĒVELE":                { dateCol: "AW", timeCol: "AX" },
  "CNC":                        { dateCol: "BC", timeCol: "BD" }
};

const TRIGGER_COL = "AJ";
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
  if (listenerRegistered) {
    console.log("ZAZA: Klausītājs jau reģistrēts.");
    return;
  }
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
// GALVENĀ LOĢIKA
// ============================================================

async function handleChange(event) {
  // Infinite loop aizsardzība
  if (isProcessing) return;

  // Tikai lietotāja ievade
  if (event.changeType !== "RangeEdited") return;

  // Notīra adresi
  let address = event.address;
  if (address.includes("!")) address = address.split("!")[1];
  address = address.replace(/\$/g, "");

  // Pārbauda vai kolonna ir AJ
  const firstCell = address.split(":")[0];
  const colLetter = firstCell.replace(/[0-9]/g, "").toUpperCase();
  if (colLetter !== TRIGGER_COL) return;

  console.log("ZAZA: Izmaiņa AJ kolonnā — adrese: " + address);

  await Excel.run(async (context) => {
    // Iegūst lapas nosaukumu
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();

    const sheetName = sheet.name.trim();
    console.log("ZAZA: Lapa — " + sheetName);

    const config = SHEET_CONFIG[sheetName];
    if (!config) {
      console.log("ZAZA: Lapa nav konfigurēta — ignorē.");
      return;
    }

    // Ielādē mainītās šūnas
    const changedRange = sheet.getRange(address);
    changedRange.load(["values", "rowIndex", "rowCount"]);
    await context.sync();

    // Savāc rindas ar "+"
    const rowsToProcess = [];
    for (let i = 0; i < changedRange.rowCount; i++) {
      const raw = changedRange.values[i][0];
      const val = String(raw === null || raw === undefined ? "" : raw).trim();
      if (val === "+") {
        rowsToProcess.push(changedRange.rowIndex + i + 1);
      }
    }

    if (rowsToProcess.length === 0) {
      console.log("ZAZA: Nav '+' vērtību — ignorē.");
      return;
    }

    console.log("ZAZA: Apstrādā " + rowsToProcess.length + " rind(as): " + rowsToProcess.join(", "));

    // Datums un laiks kā teksts
    const now      = new Date();
    const dateText = formatDate(now);
    const timeText = formatTime(now);

    // Raksta vērtības
    isProcessing = true;
    try {
      for (const rowNum of rowsToProcess) {
        const dateCell = sheet.getRange(config.dateCol + rowNum);
        const timeCell = sheet.getRange(config.timeCol + rowNum);
        dateCell.load("values");
        timeCell.load("values");
        await context.sync();

        // Pārbauda vai jau ir dati
        const existingVal = String(dateCell.values[0][0]).trim();
        if (existingVal !== "" && existingVal !== "0" && existingVal !== "false") {
          console.log("ZAZA: Rinda " + rowNum + " — datums jau eksistē, izlaižam.");
          continue;
        }

        // Ieraksta
        dateCell.values = [[dateText]];
        timeCell.values = [[timeText]];
        console.log("ZAZA: ✅ Rinda " + rowNum + " → " + config.dateCol + "=" + dateText + " | " + config.timeCol + "=" + timeText);
      }
      await context.sync();
    } finally {
      isProcessing = false;
    }

  }).catch((e) => {
    isProcessing = false;
    console.error("ZAZA: handleChange kļūda:", e);
  });
}

// ============================================================
// STATUS FUNKCIJA (ribbon poga)
// ============================================================

function showStatus(event) {
  console.log("ZAZA Timestamp aktīvs. Konfigurētās lapas: " + Object.keys(SHEET_CONFIG).join(", "));
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
