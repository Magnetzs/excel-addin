// ============================================================
// TIMESTAMP UTILITY — Excel Add-in
// Katrai darblapai atsevišķi saglabāti iestatījumi
// ============================================================

const SETTINGS_KEY_PREFIX = "timestamp_settings_";
let registeredSheets = new Set();
let currentSheetName = "";

// ============================================================
// INICIALIZĀCIJA
// ============================================================

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  await loadCurrentSheetSettings();
  document.getElementById("saveBtn").onclick = saveSettings;
  await registerAllSheetListeners();
  updateListenerStatus(true);
});

// ============================================================
// IELĀDĒ PAŠREIZĒJĀS DARBA LAPAS IESTATĪJUMUS
// ============================================================

async function loadCurrentSheetSettings() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      currentSheetName = sheet.name;
      document.getElementById("sheetInfo").textContent =
        "Darblapa: " + currentSheetName;

      const settings = getSheetSettings(currentSheetName);
      if (settings) {
        document.getElementById("triggerCol").value = settings.triggerCol || "";
        document.getElementById("dateCol").value = settings.dateCol || "";
        document.getElementById("timeCol").value = settings.timeCol || "";
      } else {
        document.getElementById("triggerCol").value = "";
        document.getElementById("dateCol").value = "";
        document.getElementById("timeCol").value = "";
      }
    });
  } catch (e) {
    showStatus("Kļūda ielādējot iestatījumus: " + e.message, "error");
  }
}

// ============================================================
// SAGLABĀ IESTATĪJUMUS PAŠREIZĒJAI DARBLAPAI
// ============================================================

async function saveSettings() {
  const triggerCol = document.getElementById("triggerCol").value.toUpperCase().trim();
  const dateCol = document.getElementById("dateCol").value.toUpperCase().trim();
  const timeCol = document.getElementById("timeCol").value.toUpperCase().trim();

  if (!triggerCol || !dateCol || !timeCol) {
    showStatus("❌ Lūdzu aizpildi visas 3 kolonnas!", "error");
    return;
  }

  if (triggerCol === dateCol || triggerCol === timeCol || dateCol === timeCol) {
    showStatus("❌ Kolonnas nedrīkst sakrist!", "error");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      currentSheetName = sheet.name;
      const settings = { triggerCol, dateCol, timeCol };
      saveSheetSettings(currentSheetName, settings);

      document.getElementById("sheetInfo").textContent =
        "Darblapa: " + currentSheetName;
    });

    await registerAllSheetListeners();
    showStatus("✅ Iestatījumi saglabāti darblapai!", "success");
    updateListenerStatus(true);

  } catch (e) {
    showStatus("❌ Kļūda: " + e.message, "error");
  }
}

// ============================================================
// IESTATĪJUMU SAGLABĀŠANA (CustomXmlParts vai roamingSettings)
// ============================================================

function getSettingsKey(sheetName) {
  return SETTINGS_KEY_PREFIX + sheetName.replace(/[^a-zA-Z0-9]/g, "_");
}

function saveSheetSettings(sheetName, settings) {
  try {
    const key = getSettingsKey(sheetName);
    Office.context.roamingSettings.set(key, JSON.stringify(settings));
    Office.context.roamingSettings.saveAsync();
  } catch (e) {
    console.error("Kļūda saglabājot iestatījumus:", e);
  }
}

function getSheetSettings(sheetName) {
  try {
    const key = getSettingsKey(sheetName);
    const raw = Office.context.roamingSettings.get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

// ============================================================
// REĢISTRĒ KLAUSĪTĀJUS VISĀM DARBLAPĀM
// ============================================================

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
        }
      }
      await context.sync();
    });
  } catch (e) {
    console.error("Kļūda reģistrējot klausītājus:", e);
    updateListenerStatus(false);
  }
}

// ============================================================
// GALVENĀ LOĢIKA — APSTRĀDĀ IZMAIŅAS
// ============================================================

async function handleChange(event) {
  if (event.changeType !== "RangeEdited") return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();

    const sheetName = sheet.name;
    const settings = getSheetSettings(sheetName);
    if (!settings) return;

    const { triggerCol, dateCol, timeCol } = settings;

    // Iegūst mainītā diapazona adresi
    const address = event.address.includes("!")
      ? event.address.split("!")[1]
      : event.address;

    // Pārbauda vai izmaiņa ir trigera kolonnā
    const changedRange = sheet.getRange(address);
    changedRange.load(["values", "rowIndex", "rowCount", "columnIndex"]);
    await context.sync();

    // Iegūst kolonnas burtu
    const colLetter = address.replace(/[0-9$]/g, "").toUpperCase();
    if (colLetter !== triggerCol) return;

    // Apstrādā KATRU rindu (bulk input atbalsts)
    const now = new Date();
    const excelDate = dateToExcelSerial(now);
    const excelTime = timeToExcelSerial(now);

    for (let i = 0; i < changedRange.rowCount; i++) {
      const cellValue = String(changedRange.values[i][0]).trim();
      if (cellValue !== "+") continue;

      const rowNumber = changedRange.rowIndex + i + 1;

      // Pārbauda vai mērķa šūnas ir tukšas
      const dateCellAddr = dateCol + rowNumber;
      const timeCellAddr = timeCol + rowNumber;

      const dateCell = sheet.getRange(dateCellAddr);
      const timeCell = sheet.getRange(timeCellAddr);
      dateCell.load("values");
      timeCell.load("values");
      await context.sync();

      const dateCellVal = String(dateCell.values[0][0]).trim();
      const timeCellVal = String(timeCell.values[0][0]).trim();

      // Ja šūnās jau ir dati — izlaist šo rindu
      if (dateCellVal !== "" && dateCellVal !== "0") continue;
      if (timeCellVal !== "" && timeCellVal !== "0") continue;

      // Ieraksta datumu un laiku
      dateCell.values = [[excelDate]];
      dateCell.numberFormat = [["dd.mm.yyyy"]];

      timeCell.values = [[excelTime]];
      timeCell.numberFormat = [["hh:mm"]];
    }

    await context.sync();
  }).catch((e) => console.error("handleChange kļūda:", e));
}

// ============================================================
// DATUMA UN LAIKA KONVERTĀCIJA UZ EXCEL SERIAL NUMBER
// ============================================================

function dateToExcelSerial(date) {
  // Excel skaita dienas no 1900-01-01 (sērijas numurs 1)
  const excelEpoch = new Date(1899, 11, 30);
  const diff = date - excelEpoch;
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function timeToExcelSerial(date) {
  // Laiks kā decimāldaļa no 0 (00:00) līdz 1 (24:00)
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = date.getSeconds();
  return (hours * 3600 + minutes * 60 + seconds) / 86400;
}

// ============================================================
// UI PALĪGFUNKCIJAS
// ============================================================

function showStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "status " + type;
  if (type === "success") {
    setTimeout(() => { el.className = "status"; }, 4000);
  }
}

function updateListenerStatus(active) {
  const dot = document.getElementById("listenerDot");
  const text = document.getElementById("listenerText");
  if (active) {
    dot.className = "dot active";
    text.textContent = "Klausītājs aktīvs — " + registeredSheets.size + " darblapa(s)";
  } else {
    dot.className = "dot inactive";
    text.textContent = "Klausītājs nav aktīvs";
  }
}
