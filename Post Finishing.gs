function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("POST FINISHING FORM");
}

function submitPreFinishing(data) {
  if (!data || !Array.isArray(data.items)) {
    throw new Error("Invalid form submission.");
  }

  const errors = [];
  const validRows = [];

  data.items.forEach((row, index) => {
    const rowNumber = index + 1;

    const item = String(row.item || "").trim();
    const rolls = String(row.rolls || "").trim();
    const weight = String(row.weight || "").trim();
    const wastage = String(row.wastage || "").trim(); // ✅ NEW

    const isCompletelyEmpty =
      !item && !rolls && !weight && !wastage;

    // ignore fully empty rows
    if (isCompletelyEmpty) return;

    // partial row = error
    if (!item || !rolls || !weight || !wastage) {
      if (!item) errors.push(`Row ${rowNumber}: Item is required`);
      if (!rolls) errors.push(`Row ${rowNumber}: Rolls is required`);
      if (!weight) errors.push(`Row ${rowNumber}: Weight is required`);
      if (!wastage) errors.push(`Row ${rowNumber}: Wastage is required`);
      return;
    }

    validRows.push({
      item,
      rolls,
      weight,
      wastage
    });
  });

  if (validRows.length === 0) {
    errors.push("At least ONE complete row is required.");
  }

  if (errors.length > 0) {
    throw new Error(errors.join("\n"));
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Report") || ss.insertSheet("Report");

  // ✅ Header (LAST COLUMN = Wastage)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp",
      "Date",
      "Shift",
      "Item",
      "Rolls",
      "Weight",
      "Wastage"
    ]);
  }

  const tz = Session.getScriptTimeZone();
  const entryDate = new Date(data.date);
  const onlyDate = Utilities.formatDate(entryDate, tz, "dd/MM/yyyy");
  const now = new Date();

  // ✅ Save rows
  validRows.forEach(r => {
    sheet.appendRow([
      now,
      onlyDate,
      data.shift,
      r.item,
      r.rolls,
      r.weight,
      r.wastage
    ]);
  });

  return "Saved Successfully";
}
