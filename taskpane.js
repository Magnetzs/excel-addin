// ============================================================
// TIMESTAMP UTILITY — Excel Add-in v2.2
// ============================================================

const SETTINGS_KEY_PREFIX = "timestamp_settings_";
let registeredSheets = new Set();
let currentSheetName = "";

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  await loadCurrentSheetSettings();
  document.getElementById("saveBtn").onclick = saveSettings;
  await registerAllSheetListeners();
});

async function loadCurrentSheetSettings() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      currentSheetName = sheet.name;
      document.getElementById("sheetInfo").textContent = "Darblapa: " + currentSheetName;
      const settings = getSheetSettings(currentSheetName);
      if (settings) {
        document.getElementById("triggerCol").value = settings.triggerCol || "";
        document.getElementById("dateCol").value    = settings.dateCol    || "";
        document.getElementById("timeCol").value    = settings.timeCol    || "";
      }
    });
  } catch (e) {
    showStatus("Kļūda ielādējot: " + e.message, "error");
  }
}

async function saveSettings() {
  const triggerCol = document.getElementById("triggerCol").value.toUpperCase().trim();
  const dateCol    = document.getElementById("dateCol").value.toUpperCase().trim();
  const timeCol    = document.getElementById("timeCol").value.toUpperCase().trim();

  if (!triggerCol || !dateCol || !timeCol) {
    showStatus("❌ Lūdzu aizpildi visas 3 kolonnas!", "error"); return;
  }
  if (triggerCol === dateCol || triggerCol === timeCol || dateCol === timeCol) {
    showStatus("❌ Kolonnas nedrīkst sakrist!", "error"); return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      currentSheetName = sheet.name;
      document.getElementById("sheetInfo").textContent = "Darblapa: " + currentSheetName;
    });

    saveSheetSettings(currentSheetName, { triggerCol, dateCol, timeCol });
    await registerAllSheetListeners();
    showStatus("✅ Saglabāts! Ievadi + kolonnā " + triggerCol, "success");
    updateListenerStatus(true);
  } catch (e) {
    showStatus("❌ Kļūda: " + e.message, "error");
  }
}

// --- Storage (localStorage) ---

function getSettingsKey(sheetName) {
  return SETTINGS_KEY_PREFIX + sheetName.replace(/[^a-zA-Z0-9]/g, "_");
}

function saveSheetSettings(sheetName, settings) {
  try {
    localStorage.setItem(getSettingsKey(sheetName), JSON.stringify(settings));
  } catch (e) { console.error("Save error:", e); }
}

function getSheetSettings(sheetName) {
  try {
    const raw = localStorage.getItem(getSettingsKey(sheetName));
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

// --- Klausītāji ---

async function registerAllSheetListeners() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      for (const sheet of sheets.items) {
        if (!registeredSheets.has(sheet.name)) {
          sheet.onChanged.add(handleChange);
          registeredSheets.add(sheet.name);
          console.log("Klausītājs reģistrēts: " + sheet.name);
        }
      }
      await context.sync();
    });
    updateListenerStatus(true);
  } catch (e) {
    console.error("Listener error:", e);
    updateListenerStatus(false);
  }
}

// --- Galvenā loģika ---

async function handleChange(event) {
  if (event.changeType !== "RangeEdited") return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();

    const sheetName = sheet.name;
    const settings = getSheetSettings(sheetName);
    if (!settings) { console.log("Nav iestatījumu: " + sheetName); return; }

    const { triggerCol, dateCol, timeCol } = settings;

    let address = event.address;
    if (address.includes("!")) address = address.split("!")[1];
    address = address.replace(/\$/g, "");

    const firstCell = address.split(":")[0];
    const colLetter = firstCell.replace(/[0-9]/g, "").toUpperCase();

    console.log("Izmaiņa: " + address + " | Kolonna: " + colLetter + " | Trigera: " + triggerCol);

    if (colLetter !== triggerCol) return;

    const changedRange = sheet.getRange(address);
    changedRange.load(["values", "rowIndex", "rowCount"]);
    await context.sync();

    const now = new Date();
    const excelDate = dateToExcelSerial(now);
    const excelTime = timeToExcelSerial(now);

    for (let i = 0; i < changedRange.rowCount; i++) {
      const rawVal = changedRange.values[i][0];
      const cellValue = String(rawVal === null || rawVal === undefined ? "" : rawVal).trim();
      console.log("Rinda " + (changedRange.rowIndex + i + 1) + ": '" + cellValue + "'");
      if (cellValue !== "+") continue;

      const rowNumber  = changedRange.rowIndex + i + 1;
      const dateCell   = sheet.getRange(dateCol + rowNumber);
      const timeCell   = sheet.getRange(timeCol + rowNumber);
      dateCell.load("values");
      timeCell.load("values");
      await context.sync();

      const existingDate = String(dateCell.values[0][0]).trim();
      if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") {
        console.log("Datums jau ir rindā " + rowNumber); continue;
      }

      dateCell.values       = [[excelDate]];
      dateCell.numberFormat = [["dd.mm.yyyy"]];
      timeCell.values       = [[excelTime]];
      timeCell.numberFormat = [["hh:mm"]];
      console.log("✅ Ierakstīts rindā " + rowNumber);
    }

    await context.sync();
  }).catch((e) => console.error("handleChange kļūda:", e));
}

// --- Datuma/laika konvertācija ---

function dateToExcelSerial(date) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const dateUTC = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return (dateUTC - excelEpoch) / (1000 * 60 * 60 * 24);
}

function timeToExcelSerial(date) {
  return (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86400;
}

// --- UI ---

function showStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status " + type;
  if (type === "success") setTimeout(() => { el.className = "status"; }, 5000);
}

function updateListenerStatus(active) {
  const dot  = document.getElementById("listenerDot");
  const text = document.getElementById("listenerText");
  if (active) {
    dot.className = "dot active";
    text.textContent = "Klausītājs aktīvs — " + registeredSheets.size + " darblapa(s)";
  } else {
    dot.className = "dot inactive";
    text.textContent = "Klausītājs nav aktīvs";
  }
}
