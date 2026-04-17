function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dispatch Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Robust Dispatch ID generator that ignores trailing blank rows.
 */
function generateDispatchID(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "DIS-0001"; // only header

  const count = lastRow - 1;
  const colA = sheet.getRange(2, 1, count, 1).getValues().map(r => r[0]);

  for (let i = colA.length - 1; i >= 0; i--) {
    const val = String(colA[i] ?? "").trim();
    if (val && val.startsWith("DIS-")) {
      const parts = val.split("-");
      const n = parseInt(parts[1], 10);
      const nextNum = isNaN(n) ? 1 : (n + 1);
      return "DIS-" + String(nextNum).padStart(4, "0");
    }
  }
  return "DIS-0001";
}

/* =====================================================
   🛠 ADD ONLY – GLOBAL UNIQUE DISPATCH ID GENERATOR
   Ensures 100% uniqueness even on same millisecond submits
===================================================== */
function generateUltraUniqueDispatchID(sheet) { // 🛠 ADD ONLY
  const lock = LockService.getScriptLock();       // 🛠 ADD ONLY
  lock.waitLock(30000);                           // 🛠 ADD ONLY

  try {                                           // 🛠 ADD ONLY
    const baseID = generateDispatchID(sheet);     // 🛠 ADD ONLY
    const ts = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyMMddHHmmssSSS"
    );                                            // 🛠 ADD ONLY

    // Example: DIS-0023-250129103015123
    return baseID + "-" + ts;                     // 🛠 ADD ONLY
  } finally {                                     // 🛠 ADD ONLY
    lock.releaseLock();                           // 🛠 ADD ONLY
  }
}

function submitDispatch(form) {
  if (!form || !form.items || form.items.length === 0)
    throw new Error('No items submitted.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dispatch responses');

  if (!sheet) {
    sheet = ss.insertSheet('Dispatch responses');
    sheet.appendRow([
      'Dispatch ID','Date Time','Party Details','Ship To',
      'Item Name','Item Description','Item Quantity','Unit','Remark'
    ]);
  }

  /* ===============================================
     🛠 ADD ONLY – AUTO DATE TIME FALLBACK
     Manual allowed, blank → auto current datetime
  =============================================== */
  const finalDateTime = form.dateTime
    ? form.dateTime
    : new Date(); // 🛠 ADD ONLY

  /* ===============================================
     🛠 ADD ONLY – ULTRA UNIQUE DISPATCH ID
  =============================================== */
  const dispatchID = generateUltraUniqueDispatchID(sheet); // 🛠 ADD ONLY

  const rows = [];
  form.items.forEach((it, index) => {
    if (!finalDateTime || !form.shipTo)
      throw new Error('Missing parent fields.');

    if (!it.name || !it.quantity || !it.unit)
      throw new Error('Each item needs name, quantity, and unit.');

    const partyDetailsValue = form.partyDetails || '';

    rows.push([
      dispatchID,
      finalDateTime,          // 🛠 ADD ONLY (replaced safely)
      partyDetailsValue,
      form.shipTo,
      it.name,
      it.description || '',
      it.quantity,
      it.unit,
      it.remark || ''
    ]);
  });

  // -------------------------------------------------------
  // NEW: preserve formulas in columns after I (col 9).
  // -------------------------------------------------------
  const lastRow = sheet.getLastRow();
  let writeStart;

  if (lastRow > 1) {
    const nextRowValues = sheet.getRange(lastRow + 1, 1, 1, 9).getValues()[0];
    const hasContentInAtoI = nextRowValues.some(cell => String(cell).trim() !== '');

    if (!hasContentInAtoI) {
      writeStart = lastRow + 2;
    } else {
      sheet.getRange(lastRow + 1, 1, 1, 9).clearContent();
      writeStart = lastRow + 2;
    }
  } else {
    writeStart = 2;
  }

  sheet.getRange(writeStart, 1, rows.length, 9).setValues(rows);

  return {
    status: 'success',
    dispatchID: dispatchID
  };
}
