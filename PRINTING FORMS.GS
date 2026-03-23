/*************************************************
 * INWARD PRINTING FORM – UPDATED VERSION
 * SAME SHEET SCRIPT | SHEET NAME = INWARD
 *************************************************/

function doGet() {
  return HtmlService.createTemplateFromFile('form')
    .evaluate()
    .setTitle('Printing Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ================================
   FETCH ITEM NAMES (Sheet: Items)
================================ */
function getItems(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Items");
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 1)
           .getValues()
           .flat()
           .filter(String);
}

/*
 FINAL COLUMN ORDER (Sheet: INWARD)
 Timestamp | Date | Shift | Item | Rolls | Weight | Wastage | Total Briquette
*/
function _getSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("INWARD");

  if (!sh) {
    sh = ss.insertSheet("INWARD");
  }

  const headers = [
    "Timestamp (Date & Time)",
    "Date",
    "Shift",
    "Item",
    "Rolls",
    "Weight (kg)",
    "Wastage (kg)",
    "Total Briquette Use (kg)"
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sh.getRange("A:A").setNumberFormat("dd-MM-yyyy HH:mm:ss");
  sh.getRange("B:B").setNumberFormat("dd-MM-yyyy");

  return sh;
}

/* ================================
   ✅ NEW HELPER (ADDED – NOTHING REMOVED)
   ACTUAL LAST DATA ROW FINDER
================================ */
function getLastDataRow(sh, col) {
  const values = sh.getRange(2, col, sh.getMaxRows()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      return i + 2; // actual last filled row
    }
  }
  return 1; // only header exists
}

/* ================================
   SAVE FORM DATA (BULLETPROOF)
================================ */
function saveFormData(rows, briq, shift, entryDate){
  try {
    if (!entryDate) throw new Error("Date missing");
    if (!shift) throw new Error("Shift missing");
    if (!Array.isArray(rows)) throw new Error("Invalid rows data");

    const errors = [];
    const validRows = [];

    rows.forEach((r, i) => {
      const rowNo = i + 1;

      const item = String(r.item || "").trim();
      const rolls = String(r.rolls || "").trim();
      const weight = String(r.weight || "").trim();
      const wastage = String(r.wastage || "").trim();

      // ✅ completely empty row → ignore
      if (!item && !rolls && !weight && !wastage) return;

      // ❌ partial row → block everything
      if (!item || !rolls || !weight) {
        if (!item) errors.push(`Row ${rowNo}: Item is required`);
        if (!rolls) errors.push(`Row ${rowNo}: Rolls is required`);
        if (!weight) errors.push(`Row ${rowNo}: Weight is required`);
        return;
      }

      // ✅ fully valid row
      validRows.push({
        item,
        rolls,
        weight,
        wastage: wastage ? parseFloat(wastage) : 0
      });
    });

    if (validRows.length === 0) {
      errors.push("At least ONE complete row is required");
    }

    if (errors.length > 0) {
      throw new Error(errors.join("\n"));
    }

    // ---------------- SAVE AFTER 100% VALIDATION ----------------

    const sh = _getSheet();
    const entryD = new Date(entryDate);
    const now = new Date();

    const finalRows = validRows.map(r => [
      now,
      entryD,
      shift,
      r.item,
      r.rolls,
      r.weight,
      r.wastage,
      briq
    ]);

    // ✅ FIXED INSERT POSITION (ACTUAL LAST DATA ROW)
    const lastDataRow = getLastDataRow(sh, 1); // Column A (Timestamp)

    sh.getRange(
      lastDataRow + 1,
      1,
      finalRows.length,
      finalRows[0].length
    ).setValues(finalRows);

    return { status: "OK", rows: finalRows.length };

  } catch (err) {
    return { status: "ERROR", message: err.message };
  }
}
