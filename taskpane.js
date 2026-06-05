// ============================================================
// TIMESTAMP UTILITY — Excel Add-in v2.3
// Optimizēts bulk operācijām — viens context.sync visam
// ============================================================

const SETTINGS_KEY_PREFIX = "timestamp_settings_";
let registeredSheets = new Set();
let currentSheetName = "";
let pendingChanges = new Map(); // debounce buferis
let debounceTimer = null;
const DEBOUNCE_MS = 600; // gaida kamēr lietotājs beidz rakstīt

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  await loadCurrentSheetSettings();
  document.getElementById("saveBtn").onclick = saveSettings;
  await registerAllSheetListeners();
});

// ============================================================
// IESTATĪJUMU IELĀDE UN SAGLABĀŠANA
// ============================================================

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

// ============================================================
// STORAGE
// ============================================================

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

// ============================================================
// KLAUSĪTĀJI
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
    updateListenerStatus(true);
  } catch (e) {
    console.error("Listener error:", e);
    updateListenerStatus(false);
  }
}

// ============================================================
// GALVENĀ LOĢIKA — DEBOUNCE + BULK
// ============================================================

async function handleChange(event) {
  if (event.changeType !== "RangeEdited") return;

  // Iegūst adresi un darblapas nosaukumu
  let address = event.address;
  if (address.includes("!")) address = address.split("!")[1];
  address = address.replace(/\$/g, "");

  // Iegūst darblapas nosaukumu no event
  let sheetName = "";
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      sheetName = sheet.name;
    });
  } catch (e) { return; }

  const settings = getSheetSettings(sheetName);
  if (!settings) return;

  // Pārbauda vai izmaiņa ir trigera kolonnā
  const firstCell = address.split(":")[0];
  const colLetter = firstCell.replace(/[0-9]/g, "").toUpperCase();
  if (colLetter !== settings.triggerCol) return;

  // Pievieno buferim
  const key = sheetName + "|" + address;
  pendingChanges.set(key, { sheetName, address, settings });

  // Debounce — gaida 600ms pēc pēdējās izmaiņas
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(processPendingChanges, DEBOUNCE_MS);
}

// ============================================================
// APSTRĀDĀ VISAS BUFERĒTĀS IZMAIŅAS VIENĀ REIZĒ
// ============================================================

async function processPendingChanges() {
  if (pendingChanges.size === 0) return;

  const changes = new Map(pendingChanges);
  pendingChanges.clear();

  // Grupē pēc darblapas
  const bySheet = new Map();
  for (const [, change] of changes) {
    if (!bySheet.has(change.sheetName)) {
      bySheet.set(change.sheetName, []);
    }
    bySheet.get(change.sheetName).push(change);
  }

  for (const [sheetName, sheetChanges] of bySheet) {
    await processSheetChanges(sheetName, sheetChanges);
  }
}

async function processSheetChanges(sheetName, changes) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      const settings = getSheetSettings(sheetName);
      if (!settings) return;

      const { triggerCol, dateCol, timeCol } = settings;

      // Ielādē VISAS trigera šūnas uzreiz
      const allRanges = [];
      for (const change of changes) {
        const range = sheet.getRange(change.address);
        range.load(["values", "rowIndex", "rowCount"]);
        allRanges.push({ range, address: change.address });
      }

      await context.sync();

      // Savāc visas rindas kur ir "+"
      const rowsToProcess = [];
      for (const { range } of allRanges) {
        for (let i = 0; i < range.rowCount; i++) {
          const rawVal = range.values[i][0];
          const cellValue = String(rawVal === null || rawVal === undefined ? "" : rawVal).trim();
          if (cellValue !== "+") continue;
          rowsToProcess.push(range.rowIndex + i + 1);
        }
      }

      if (rowsToProcess.length === 0) return;

      // Ielādē VISAS mērķa šūnas uzreiz (viens sync)
      const dateCells = rowsToProcess.map(r => {
        const c = sheet.getRange(dateCol + r);
        c.load("values");
        return c;
      });
      const timeCells = rowsToProcess.map(r => {
        const c = sheet.getRange(timeCol + r);
        c.load("values");
        return c;
      });

      await context.sync();

      // Aprēķina datumu/laiku vienreiz
      const now = new Date();
      const excelDate = dateToExcelSerial(now);
      const excelTime = timeToExcelSerial(now);

      // Ieraksta VISAS vērtības (bez papildu sync)
      let count = 0;
      for (let i = 0; i < rowsToProcess.length; i++) {
        const existingDate = String(dateCells[i].values[0][0]).trim();
        if (existingDate !== "" && existingDate !== "0" && existingDate !== "false") continue;

        dateCells[i].values       = [[excelDate]];
        dateCells[i].numberFormat = [["dd.mm.yyyy"]];
        timeCells[i].values       = [[excelTime]];
        timeCells[i].numberFormat = [["hh:mm"]];
        count++;
      }

      // VIENS sync visām izmaiņām
      await context.sync();
      console.log("✅ Ierakstīts " + count + " rinda(s) darblapā: " + sheetName);

    });
  } catch (e) {
    // Ja Excel vēl editing mode — mēģina vēlreiz pēc 1 sekundes
    if (e.message && e.message.includes("cell-editing mode")) {
      console.log("Excel editing mode, mēģina vēlreiz...");
      setTimeout(() => processSheetChanges(sheetName, changes), 1000);
    } else {
      console.error("processSheetChanges kļūda:", e);
    }
  }
}

// ============================================================
// DATUMA UN LAIKA KONVERTĀCIJA
// ============================================================

function dateToExcelSerial(date) {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const dateUTC = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return (dateUTC - excelEpoch) / (1000 * 60 * 60 * 24);
}

function timeToExcelSerial(date) {
  return (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86400;
}

// ============================================================
// UI
// ============================================================

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
