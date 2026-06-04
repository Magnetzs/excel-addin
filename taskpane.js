let settings = {
  triggerColumn: "",
  resultColumn: ""
};

let handlerRegistered = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    loadSettings();
    document.getElementById("saveButton").onclick = saveSettings;
    registerChangeHandler();
  }
});

function saveSettings() {
  const triggerCol = document.getElementById("triggerColumn").value.toUpperCase().trim();
  const resultCol = document.getElementById("resultColumn").value.toUpperCase().trim();

  if (!triggerCol || !resultCol) {
    document.getElementById("status").innerText = "Ludzu ievadi abas kolonnas!";
    document.getElementById("status").style.color = "red";
    setTimeout(() => { document.getElementById("status").innerText = ""; }, 3000);
    return;
  }

  settings.triggerColumn = triggerCol;
  settings.resultColumn = resultCol;

  try {
    const roaming = Office.context.roamingSettings;
    roaming.set("triggerColumn", triggerCol);
    roaming.set("resultColumn", resultCol);
    roaming.saveAsync(() => {});
  } catch (e) {}

  document.getElementById("status").style.color = "green";
  document.getElementById("status").innerText = "Iestatijumi saglabati!";
  setTimeout(() => { document.getElementById("status").innerText = ""; }, 3000);
}

function loadSettings() {
  try {
    const roaming = Office.context.roamingSettings;
    const savedTrigger = roaming.get("triggerColumn");
    const savedResult = roaming.get("resultColumn");
    if (savedTrigger) { settings.triggerColumn = savedTrigger; document.getElementById("triggerColumn").value = savedTrigger; }
    if (savedResult) { settings.resultColumn = savedResult; document.getElementById("resultColumn").value = savedResult; }
  } catch (e) {}
}

function registerChangeHandler() {
  if (handlerRegistered) return;
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(handleChange);
    await context.sync();
    handlerRegistered = true;
  }).catch((e) => console.error("Handler error:", e));
}

async function handleChange(event) {
  if (!settings.triggerColumn || !settings.resultColumn) return;
  if (event.changeType !== "RangeEdited") return;

  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let changedRange = sheet.getRange(event.address);
    changedRange.load(["values", "rowIndex"]);
    await context.sync();

    const address = event.address.includes("!") ? event.address.split("!")[1] : event.address;
    const currentColumnLetter = address.replace(/[0-9]/g, "").toUpperCase();

    if (currentColumnLetter === settings.triggerColumn) {
      const enteredValue = changedRange.values[0][0];
      if (enteredValue === "+") {
        const rowNumber = changedRange.rowIndex + 1;
        const resultAddress = settings.resultColumn + rowNumber;
        sheet.getRange(resultAddress).values = [[123]];
        await context.sync();
      }
    }
  }).catch((e) => console.error("handleChange error:", e));
}
