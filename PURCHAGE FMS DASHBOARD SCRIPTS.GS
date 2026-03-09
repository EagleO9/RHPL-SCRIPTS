function onOpen() {
 var ui = SpreadsheetApp.getUi();
 ui.createMenu('BMP Formulas')
     .addItem('⚙️ Initial Setup (Run First Time)', 'initialSetup')
     .addSeparator()
     .addSubMenu(ui.createMenu('📊 FMS to Looker Conversion')
         .addItem('1️⃣ Setup Looker Studio Sheets', 'createSetup')
         .addItem('2️⃣ Create Form with Prefilled URL', 'createFormWithPrefilledURL')
         .addItem('3️⃣ Convert FMS to Looker Format', 'convertFMStoDBWithDoers'))
     .addSeparator()
     .addItem('📥 Assisted ImportRange', 'importRangeFormula')
     .addSeparator()
     .addSubMenu(ui.createMenu('📐 FMS Formulas')
         .addItem('TAT with Working Hours', 'plannedwwh')
         .addItem('TAT in Days', 'plannedindays')
         .addItem('T-x Formula', 'plannedlead')
         .addItem('Specific Time', 'specificTime')
         .addItem('Show Planned Only When Status is NO', 'tatifno')
         .addItem('Show Planned Only When Status is YES', 'tatifyes')
         .addItem('Set Actual Time', 'actualTime')
         .addItem('Time Delay Formula', 'timeDelay')
         .addItem('Vlookup Wizard', 'vlookupFormula'))
     .addToUi();
}








function findAndEditFormSheet() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var ui = SpreadsheetApp.getUi();
 var sheets = ss.getSheets();
 var formResponseSheet = null;
 for (var i = 0; i < sheets.length; i++) {
   var sheet = sheets[i];
   if (sheet.isSheetHidden() || sheet.getName() === 'form_config') {
     continue;
   }
   if (sheet.getLastColumn() >= 3) {
     var headers = sheet.getRange(1, 1, 1, 3).getValues()[0];
     if (headers[0] === 'Timestamp' &&
         headers[1] === 'Unique Key' &&
         headers[2] === 'Step Code') {
       formResponseSheet = sheet;
       break;
     }
   }
 }
 if (!formResponseSheet) {
   ui.alert('Error: No form response sheet found with the expected headers (Timestamp, Unique Key, Step Code).');
   return null;
 }
 Logger.log('Found form response sheet: ' + formResponseSheet.getName());
 formResponseSheet.moveColumns(formResponseSheet.getRange("A:A"), 3);
 formResponseSheet.moveColumns(formResponseSheet.getRange("C:C"), 2);
 Logger.log('Columns rearranged successfully');
 var existingDataSheet = ss.getSheetByName('Data');
 if (existingDataSheet && existingDataSheet.getSheetId() !== formResponseSheet.getSheetId()) {
   var timestamp = new Date().getTime();
   existingDataSheet.setName('Data_old_' + timestamp);
   Logger.log('Existing "Data" sheet renamed to: Data_old_' + timestamp);
 }
 formResponseSheet.setName('Data');
 Logger.log('Sheet renamed to: Data');
 var configSheet = ss.getSheetByName('form_config');
 if (!configSheet) {
   ui.alert('Warning: form_config sheet not found. Sheet has been renamed but cannot retrieve prefilled URL.');
   return null;
 }
 var lastRow = configSheet.getLastRow();
 if (lastRow < 2) {
   ui.alert('Warning: No form configuration found in form_config sheet.');
   return null;
 }
 var prefilledUrl = configSheet.getRange(lastRow, 2).getValue();
 Logger.log('Prefilled URL retrieved: ' + prefilledUrl);
 ui.alert(
   'Form Sheet Updated Successfully!\n\n' +
   'Sheet renamed to: Data\n' +
   'Column order: Unique Key, Step Code, Timestamp\n\n' +
   'Prefilled URL retrieved from form_config.'
 );
 return prefilledUrl;
}








function createFormWithPrefilledURL() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var ui = SpreadsheetApp.getUi();
 var response = ui.prompt(
   'Create New Form',
   'Please enter the name for the new form:',
   ui.ButtonSet.OK_CANCEL
 );
 if (response.getSelectedButton() != ui.Button.OK) {
   ui.alert('Form creation cancelled.');
   return;
 }
 var formName = response.getResponseText().trim();
 if (!formName || formName === '') {
   ui.alert('Error: Form name cannot be empty. Please try again.');
   return;
 }
 var form = FormApp.create(formName);
 var uniqueKeyItem = form.addTextItem();
 uniqueKeyItem.setTitle('Unique Key')
              .setRequired(true);
 var stepCodeItem = form.addTextItem();
 stepCodeItem.setTitle('Step Code')
             .setRequired(true);
 form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
 form.setAcceptingResponses(true);
 var prefilledResponse = form.createResponse()
   .withItemResponse(uniqueKeyItem.createResponse('u123u'))
   .withItemResponse(stepCodeItem.createResponse('k123k'))
   .toPrefilledUrl();
 var formUrl = form.getPublishedUrl();
 var configSheet = ss.getSheetByName('form_config');
 if (!configSheet) {
   configSheet = ss.insertSheet('form_config');
   configSheet.getRange(1, 1, 1, 2).setValues([['Form Name', 'Prefilled URL']]);
   configSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
 }
 configSheet.hideSheet();
 var lastRow = configSheet.getLastRow();
 configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[
   formName,
   prefilledResponse
 ]]);
 configSheet.autoResizeColumns(1, 2);
 Logger.log('Form created: ' + formName);
 Logger.log('Form URL: ' + formUrl);
 Logger.log('Prefilled URL: ' + prefilledResponse);
 ui.alert(
   'Form Created Successfully!\n\n' +
   'Form Name: ' + formName + '\n\n' +
   'Configuration saved in hidden sheet "form_config"\n\n' +
   'Next Step: Run "3️⃣ Convert FMS to DB Format" to complete the setup.'
 );
 return {
   formUrl: formUrl,
   prefilledUrl: prefilledResponse,
   formName: formName
 };
}








// Combined initial setup function
function initialSetup() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var openingTime = ui.prompt('Enter Opening Time (e.g., 9 for 9 AM)', ui.ButtonSet.OK_CANCEL);
 if (openingTime.getSelectedButton() !== ui.Button.OK) return;
 var closingTime = ui.prompt('Enter Closing Time (e.g., 18 for 6 PM)', ui.ButtonSet.OK_CANCEL);
 if (closingTime.getSelectedButton() !== ui.Button.OK) return;
 // Enable iterative calculation
 ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);
 ss.setIterativeCalculationEnabled(true);
 ss.setMaxIterativeCalculationCycles(1);
 ss.setIterativeCalculationConvergenceThreshold(0.05);
 var sheet = ss.getActiveSheet();
 sheet.getRange('A1').setFormula('=now()');
 sheet.getRange('C1').setValue(parseFloat(openingTime.getResponseText()));
 sheet.getRange('D1').setValue(parseFloat(closingTime.getResponseText()));
 sheet.getRange('B1').setFormula('=D1/24-C1/24');
 sheet.hideRows(1);
 ui.alert('✅ Setup Complete!', 'Your spreadsheet is now configured with working hours.', ui.ButtonSet.OK);
}








// ============================================================
// FMS to DB conversion — now with per-step Doer Column support
// ============================================================
function convertFMStoDBWithDoers() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var ui = SpreadsheetApp.getUi();
 // Ensure setup sheets exist (7-column Steps Directory, 10-column DB_Format)
 createSetup();
 var sheet = ss.getSheetByName("Steps Directory");
 var lastRow = sheet.getLastRow() - 1;
 if (lastRow < 1) {
   ui.alert('⚠️ No Data Found', 'Please add steps to the "Steps Directory" sheet first.', ui.ButtonSet.OK);
   return;
 }
 // Now 7 columns: Step, How, Step Code, Planned/Actual Range, Unique Key Range, Doer Column, Query
 var data = sheet.getRange(2, 1, lastRow, 7).getValues();
 var arr = [];
 var stepNames = [];
 for (var i = 0; i < data.length; i++) {
   var row = data[i];
   var step     = row[0]; // Step name
   var how      = row[1]; // How
   var keyrange = row[4]; // Unique Key Range
   var range    = row[3]; // Planned/Actual Range (2 cols)
   var doerCol  = row[5]; // Doer Column range (single col, may be empty)
 
   stepNames.push([step]);
 
   // Query column is now column 7
   var toChange = sheet.getRange(i + 2, 7);
 
   if (doerCol && doerCol.toString().trim() !== '') {
     // ── Doer column present ──
     // Data source: {keyrange, range, doerCol} → Col1=Key, Col2=Planned, Col3=Actual, Col4=DoerName
     // Output: Key, Planned, 'Step', 'How', DoerName → 5 columns
     toChange.setValue(
       'select Col1,Col2,\'' + step + '\',\'' + how + '\',Col4 ' +
       'where Col3 is null and Col1 is not null and Col2 is not null ' +
       'label \'' + step + '\' \'\',\'' + how + '\' \'\''
     );
   
     var formula = "ifna(query({" + keyrange + "," + range + "," + doerCol + "},'Steps Directory'!" + toChange.getA1Notation() + "),{\"\",\"\",\"\",\"\",\"\"})";
   } else {
     // ── No doer column ──
     // FIX: Use ' ' (space) instead of '' (empty string) because
     // Google Sheets QUERY cannot handle '' as a select expression.
     // The space is later handled by TRIM() in the Final Doer formulas.
     toChange.setValue(
       'select Col1,Col2,\'' + step + '\',\'' + how + '\',\' \' ' +
       'where Col3 is null and Col1 is not null and Col2 is not null ' +
       'label \'' + step + '\' \'\',\'' + how + '\' \'\',\' \' \'\''
     );
   
     var formula = "ifna(query({" + keyrange + "," + range + "},'Steps Directory'!" + toChange.getA1Notation() + "),{\"\",\"\",\"\",\"\",\"\"})";
   }
 
   if (i < data.length - 1) {
     formula += ";";
   }
   arr.push(formula);
 }
 // Build and set the main query formula in DB_Format A2
 var stringToSet = "=sort({" + arr.join("") + "},1,1)";
 ss.getSheetByName("DB_Format").getRange("A2").setFormula(stringToSet);
 // Create/refresh Doer Emails sheet
 createDoerEmailsSheet(stepNames);
 // Add Doer Email, Final Doer Email, Final Doer Name columns
 addDoerColumnsToDBFormat();
 // Get prefilled form URL
 var prefilledLink = findAndEditFormSheet();
 if (!prefilledLink) {
   ui.alert('⚠️ Error', 'Could not retrieve prefilled URL. Please ensure you have run "2️⃣ Create Form with Prefilled URL" first.', ui.ButtonSet.OK);
   return;
 }
 // Link formula in column F (was E)
 // $A = Unique Key, $C = Step → vlookup to Step Code
 var finalLink = prefilledLink
   .replace("u123u", "\"&$A$2:$A&\"")
   .replace("k123k", "\"&vlookup($C$2:$C,'Steps Directory'!$A2:$C,3,false)))");
 var arrayFormula = "=ArrayFormula(if(isblank($A$2:$A),\"\",\"" + finalLink;
 // For PC formula in column G, referencing K1 (=now())
 var pcFormula = "=ArrayFormula(if(isblank($B$2:$B),\"\",if($B$2:$B<$K$1,\"Yes\",\"No\")))";
 ss.getSheetByName("DB_Format").getRange("F2").setFormula(arrayFormula);
 ss.getSheetByName("DB_Format").getRange("G2").setFormula(pcFormula);
 ui.alert(
   '✅ Conversion Complete!',
   'DB_Format sheet has been updated with:\n' +
   '• Query results (with per-step Doer Name)\n' +
   '• Form links\n' +
   '• Doer Email (step-based)\n' +
   '• Final Doer Email & Final Doer Name (priority: FMS doer > step default)\n\n' +
   'Your FMS to DB conversion is complete!',
   ui.ButtonSet.OK
 );
}








// Create Doer Emails sheet (unchanged structure)
function createDoerEmailsSheet(stepNames) {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var doerSheet = ss.getSheetByName("Doer Emails");
 if (doerSheet != null) {
   doerSheet.clear();
 } else {
   doerSheet = ss.insertSheet();
   doerSheet.setName("Doer Emails");
 }
 doerSheet.appendRow(["Step Name", "Doer Email", "Doer Name"]);
 doerSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
 if (stepNames.length > 0) {
   doerSheet.getRange(2, 1, stepNames.length, 1).setValues(stepNames);
 }
 doerSheet.setColumnWidth(1, 200);
 doerSheet.setColumnWidth(2, 250);
 doerSheet.setColumnWidth(3, 200);
 doerSheet.setFrozenRows(1);
}








// ============================================================
// Add Doer columns to DB_Format — FIXED with TRIM() to handle
// space placeholders from no-doer query results
// ============================================================
function addDoerColumnsToDBFormat() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var dbSheet = ss.getSheetByName("DB_Format");
 // Set headers for columns H, I, J
 dbSheet.getRange("H1").setValue("Doer Email");
 dbSheet.getRange("I1").setValue("Final Doer Email");
 dbSheet.getRange("J1").setValue("Final Doer Name");
 // H: Step-based Doer Email (vlookup Step → Doer Email)
 var doerEmailFormula =
   '=ArrayFormula(if(isblank($C$2:$C),"",iferror(vlookup($C$2:$C,\'Doer Emails\'!$A:$B,2,false),"")))';
 dbSheet.getRange("H2").setFormula(doerEmailFormula);
 // I: Final Doer Email
 //    FIX: Use TRIM() on col E so that space placeholders from no-doer queries
 //    are treated as empty, correctly falling back to step-based email (col H)
 var finalDoerEmailFormula =
   '=ArrayFormula(if(isblank($A$2:$A),"",' +
     'if(TRIM($E$2:$E)<>"",' +
       'iferror(vlookup(TRIM($E$2:$E),{\'Doer Emails\'!$C:$C,\'Doer Emails\'!$B:$B},2,false),""),' +
       'H$2:H' +
     ')' +
   '))';
 dbSheet.getRange("I2").setFormula(finalDoerEmailFormula);
 // J: Final Doer Name
 //    FIX: Use TRIM() on col E so that space placeholders are treated as empty,
 //    correctly falling back to step-based Doer Name from Doer Emails
 var finalDoerNameFormula =
   '=ArrayFormula(if(isblank($A$2:$A),"",' +
     'if(TRIM($E$2:$E)<>"",' +
       'TRIM($E$2:$E),' +
       'iferror(vlookup($C$2:$C,\'Doer Emails\'!$A:$C,3,false),"")' +
     ')' +
   '))';
 dbSheet.getRange("J2").setFormula(finalDoerNameFormula);
 // Format all headers
 dbSheet.getRange("A1:J1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
}








// ============================================================
// createSetup — updated for 7-col Steps Directory & 10-col DB_Format
// ============================================================
function createSetup() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 // Create DB_Format sheet
 var datastudiosheet = ss.getSheetByName("DB_Format");
 if (datastudiosheet == null) {
   datastudiosheet = ss.insertSheet();
   datastudiosheet.setName("DB_Format");
   datastudiosheet.appendRow([
     "Unique Key", "Planned", "Step", "How", "Doer Name",
     "Link", "For PC", "Doer Email", "Final Doer Email", "Final Doer Name"
   ]);
   datastudiosheet.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
   datastudiosheet.getRange('F:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
   // =now() in K1 (hidden utility cell for For PC comparison)
   datastudiosheet.getRange('K1').setFormula("=now()");
   datastudiosheet.getRange("A1:J1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
 }
 // Create Steps Directory sheet — now 7 columns
 var stepsDirectory = ss.getSheetByName("Steps Directory");
 if (stepsDirectory == null) {
   stepsDirectory = ss.insertSheet();
   stepsDirectory.setName("Steps Directory");
   stepsDirectory.appendRow([
     "Step", "How", "Step Code",
     "Planned/Actual Range", "Unique Key Range",
     "Doer Column",   // ← NEW column
     "Query"
   ]);
   stepsDirectory.getRange("A1:G1").setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
   stepsDirectory.setFrozenRows(1);
 }
}








// Simplified prompt helper
function getPromptValue(ui, message) {
 var response = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
 if (response.getSelectedButton() !== ui.Button.OK) return null;
 return response.getResponseText();
}








function importRangeFormula() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var donorSheet = getPromptValue(ui, 'Enter URL of the source spreadsheet:');
 if (!donorSheet) return;
 var tabName = getPromptValue(ui, 'Enter the sheet/tab name:');
 if (!tabName) return;
 var rangeName = getPromptValue(ui, 'Enter the range (e.g., A1:Z100):');
 if (!rangeName) return;
 ss.getCurrentCell().setFormula('=importrange("' + donorSheet + '","' + tabName + '!' + rangeName + '")');
}








function vlookupFormula() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var donorColumn = getPromptValue(ui, 'Enter Dependent Cell Column (e.g., A):');
 if (!donorColumn) return;
 donorColumn = donorColumn.replace(/\$/g, '');
 var tabColumn = getPromptValue(ui, 'Enter Unique Key Cell Column (e.g., B):');
 if (!tabColumn) return;
 tabColumn = tabColumn.replace(/\$/g, '');
 var rangeName = getPromptValue(ui, 'Enter Step Code Cell (e.g., $G$1):');
 if (!rangeName) return;
 var sheet = ss.getActiveSheet();
 var cell = sheet.getActiveCell();
 var column = cell.getColumn();
 var startRow = cell.getRow();
 var lastRow = sheet.getLastRow();
 var columnValues = sheet.getRange(startRow, column, lastRow - startRow + 1).getValues();
 var formulas = [];
 for (var i = 0; i < columnValues.length; i++) {
   var row = startRow + i;
   if (!columnValues[i][0]) {
     formulas.push([
       '=IF(ISBLANK(' + donorColumn + row + '), "", IFNA(VLOOKUP(' + tabColumn + row + '&' + rangeName + ', INDEX({Data!$A:$A&Data!$B:$B, Data!$C:$C}), 2, FALSE), ""))'
     ]);
   } else {
     formulas.push([columnValues[i][0]]);
   }
 }
 sheet.getRange(startRow, column, formulas.length, 1).setFormulas(formulas);
}








// Helper function to apply formula to range
function applyFormulaToRange(formula) {
 var ss = SpreadsheetApp.getActive();
 ss.getCurrentCell().setFormula(formula);
 var currentCell = ss.getCurrentCell();
 ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
 currentCell.activateAsCurrentCell();
 currentCell = ss.getCurrentCell();
 ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
 currentCell.activateAsCurrentCell();
 ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}








function plannedwwh() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var ofs = ss.getRange("C1").getValue();
 var ofe = ss.getRange("D1").getValue();
 var fromDate = getPromptValue(ui, 'Enter Date Cell (e.g., A2):');
 if (!fromDate) return;
 var tatInHours = getPromptValue(ui, 'Enter TAT Cell (e.g., G$5):');
 if (!tatInHours) return;
 var formula = '=if(' + fromDate + ',if(and(hour(' + fromDate + '+' + tatInHours + ')>' + ofs + ',(hour(' + fromDate + '+' + tatInHours + ')<' + ofe + ')),' + fromDate + '+' + tatInHours + ',workday.intl(int(' + fromDate + '),1,"0000001")+hour(' + fromDate + '+' + tatInHours + '-$B$1)/24+minute(' + fromDate + ')/1440),"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function plannedindays() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Date Cell (e.g., A2):');
 if (!fromDate) return;
 var tatInHours = getPromptValue(ui, 'Enter TAT Cell (e.g., G$5):');
 if (!tatInHours) return;
 var formula = '=if(' + fromDate + ',WORKDAY.INTL(' + fromDate + ',' + tatInHours + ',"0000001")+hour(' + fromDate + ')/24+MINUTE(' + fromDate + ')/1440,"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function plannedlead() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Date Cell:');
 if (!fromDate) return;
 var leadtime = getPromptValue(ui, 'Enter Lead Time Cell:');
 if (!leadtime) return;
 var tatInHours = getPromptValue(ui, 'Enter Number of Days Before Lead Time:');
 if (!tatInHours) return;
 var formula = '=if(' + leadtime + ',' + fromDate + '+' + leadtime + '-' + tatInHours + ',"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function specificTime() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Date Cell:');
 if (!fromDate) return;
 var leadtime = getPromptValue(ui, 'Enter Number of Days After Previous Planned (0 for same day):');
 if (!leadtime) return;
 var tatInHours = getPromptValue(ui, 'Enter Time of Day (hour/24 format):');
 if (!tatInHours) return;
 var formula = '=if(' + fromDate + ',workday.intl(int(' + fromDate + '),' + leadtime + ',"0000001",)+' + tatInHours + ',"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function actualTime() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var leadtime = getPromptValue(ui, 'Enter Status Cell:');
 if (!leadtime) return;
 var currcella1 = ss.getCurrentCell().getA1Notation();
 var formula = '=if(' + currcella1 + ',' + currcella1 + ',if(' + leadtime + '<>"",$A$1,""))';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function timeDelay() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Planned Cell:');
 if (!fromDate) return;
 var leadtime = getPromptValue(ui, 'Enter Actual Cell:');
 if (!leadtime) return;
 var formula = '=if(' + fromDate + ',if(' + leadtime + '<>"",if(' + leadtime + '>' + fromDate + ',' + leadtime + '-' + fromDate + ',""),$A$1-' + fromDate + '),"")';
 ss.getCurrentCell().setFormula(formula);
 ss.getActiveRangeList().setNumberFormat('[h]:mm:ss');
 var spreadsheet = SpreadsheetApp.getActive();
 var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
 conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
   .setRanges([spreadsheet.getActiveRange()])
   .whenFormulaSatisfied('=if(' + leadtime + ',if(' + leadtime + '>' + fromDate + ',1,0),0)')
   .setBackground('#F4C7C3')
   .build());
 conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
   .setRanges([spreadsheet.getActiveRange()])
   .whenFormulaSatisfied('=if(' + leadtime + ',0,if(' + fromDate + '<$A$1,1,0))')
   .setBackground('#FCE8B2')
   .build());
 spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
 var currentCell = ss.getCurrentCell();
 ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
 currentCell.activateAsCurrentCell();
 currentCell = ss.getCurrentCell();
 ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
 currentCell.activateAsCurrentCell();
 ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}








function tatifno() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Status Cell:');
 if (!fromDate) return;
 var formulaincell = ss.getCurrentCell().getFormula().substr(1);
 var formula = '=if(' + fromDate + '="No",' + formulaincell + ',"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}








function tatifyes() {
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var fromDate = getPromptValue(ui, 'Enter Status Cell:');
 if (!fromDate) return;
 var formulaincell = ss.getCurrentCell().getFormula().substr(1);
 var formula = '=if(' + fromDate + '="Yes",' + formulaincell + ',"")';
 ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
 applyFormulaToRange(formula);
}




