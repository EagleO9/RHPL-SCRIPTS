function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dispatch Tools')
    .addItem('Process Dispatches', 'processDispatches')
    .addToUi();
}

function processDispatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcName = 'Dispatch responses';
  const destName = 'COMPLETE DISPATCH';

  const src = ss.getSheetByName(srcName);
  if (!src) {
    SpreadsheetApp.getUi().alert(`Sheet "${srcName}" not found.`);
    return;
  }

  let dest = ss.getSheetByName(destName);
  if (!dest) dest = ss.insertSheet(destName);

  const lastRow = src.getLastRow();
  const neededCols = 13;
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows to process.');
    return;
  }

  const srcData = src.getRange(1, 1, lastRow, neededCols).getValues();
  const header = srcData[0].slice(0, 12);

  if (dest.getLastRow() === 0) {
    dest.getRange(1, 1, 1, header.length).setValues([header]);
  }

  const destLast = dest.getLastRow();
  const destVals = dest.getRange(1, 1, destLast, 12).getValues();

  const existingMap = new Map();
  const existingRowsSet = new Set(); // 🔥 full row duplicate control

  // Existing data map + row set
  for (let i = 1; i < destVals.length; i++) {
    const dID = String(destVals[i][0] ?? '').trim();
    const itemName = String(destVals[i][4] ?? '').trim();
    const kVal = parseFloat(destVals[i][10]) || 0;

    if (dID && itemName) {
      existingMap.set(dID + '|' + itemName, { kValue: kVal });
    }

    // 🔥 store full row for duplicate prevention
    existingRowsSet.add(destVals[i].join('|'));
  }

  const toAppend = [];
  const toDelete = [];

  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const rowNum = i + 1;

    const colL = String(row[11] ?? '').trim();
    if (colL === '') continue;

    const colM = String(row[12] ?? '').trim().toUpperCase();
    if (!['PENDING', 'DISPATCH', 'EXTRA DISPATCH'].includes(colM)) continue;

    const dispatchID = String(row[0] ?? '').trim();
    const itemName = String(row[4] ?? '').trim();
    const srcK = parseFloat(row[10]) || 0;

    const key = dispatchID + '|' + itemName;
    const newRow = row.slice(0, 12);

    // ✅ Duplicate + Zero + Map Fix
    if (existingMap.has(key)) {
      const existingK = parseFloat(existingMap.get(key).kValue) || 0;
      const newK = Math.abs(existingK - srcK);

      if (newK === 0) continue; // ❌ skip zero

      newRow[10] = newK;
      existingMap.set(key, { kValue: newK });

    } else {
      if (srcK === 0) continue; // ❌ skip zero

      newRow[10] = srcK;
      existingMap.set(key, { kValue: srcK });
    }

    // 🔥 FULL ROW DUPLICATE CHECK
    const rowCheckStr = newRow.join('|');
    if (existingRowsSet.has(rowCheckStr)) continue;

    existingRowsSet.add(rowCheckStr);

    toAppend.push(newRow);

    if (colM === 'DISPATCH' || colM === 'EXTRA DISPATCH') {
      toDelete.push(rowNum);
    }
  }

  if (toAppend.length > 0) {
    const destStart = dest.getLastRow() + 1;
    dest.getRange(destStart, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }

  if (toDelete.length > 0) {
    toDelete.sort((a, b) => b - a);
    for (const r of toDelete) {
      src.deleteRow(r);
    }
  }

  SpreadsheetApp.getUi().alert(
    `✅ Processing complete.\n` +
    `${toAppend.length} rows copied to "${destName}".\n` +
    `${toDelete.length} rows deleted from "${srcName}".`
  );
}
