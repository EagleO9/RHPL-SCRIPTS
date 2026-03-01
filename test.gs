/**
 * Automatically moves rows to an Archive sheet when status is 'Completed'.
 * @param {Object} e The event object from the onEdit trigger.
 */
function onEdit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // Configuration: Adjust these to match your sheet
    const TARGET_SHEET_NAME = "Main";
    const ARCHIVE_SHEET_NAME = "Archive";
    const STATUS_COLUMN = 4; // Column D
    const SUCCESS_CRITERIA = "Completed";

    // 1. Validation: Ensure the edit happened in the right place
    if (sheet.getName() !== TARGET_SHEET_NAME || range.getColumn() !== STATUS_COLUMN || e.value !== SUCCESS_CRITERIA) {
      return;
    }

    // 2. Safety: User confirmation for data movement
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Archive Row?', 'Do you want to move this row to the Archive?', ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      const archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
      const row = range.getRow();
      const numColumns = sheet.getLastColumn();
      const rowData = sheet.getRange(row, 1, 1, numColumns).getValues();

      
      // 3. Execution: Move data and delete original row
      archiveSheet.appendRow(rowData[0]);
      sheet.deleteRow(row);
      
      ss.toast("Row successfully archived.", "Automation Success", 3);
    }
  } catch (error) {
    Logger.log(`Error in onEdit: ${error.toString()}`);
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}