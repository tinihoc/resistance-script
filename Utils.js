/**
 * Gets sheet GID by sheet name.
 */
function getSheetGidByName(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  return sheet ? sheet.getSheetId() : null;
}

/**
 * Returns the name of the current sheet and refreshes when the input cell changes.
 * @customfunction
 */
function SHEETNAME_DYNAMIC(dummy) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Generic Spreadsheet Utilities & Dynamic Formulas
 * File: Utils.gs
 */

/**
 * Gets sheet GID by sheet name.
 * @customfunction
 */
function getSheetGidByName(name) {
  if (!name) return "";
  const sheet = SpreadsheetApp.getActive().getSheetByName(String(name).trim());
  return sheet ? sheet.getSheetId() : "";
}

/**
 * Returns the name of the current sheet and refreshes when the input cell changes.
 * @customfunction
 */
function SHEETNAME_DYNAMIC(dummy) {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/**
 * Custom function used by draggable Column C formula.
 * Returns the target date cell value, honoring the unscheduled bypass rule.
 * @customfunction
 */
function GET_EXERCISE_DATE(sheetName, dummy) {
  if (!sheetName) return "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(String(sheetName).trim());
  if (!sheet) return "";

  const sections = getExerciseSections(sheet);
  if (sections.length === 0) return "";

  const latest = sections[sections.length - 1];
  let reflected = latest;
  if (latest.statusVal === 'unscheduled' && sections.length > 1) {
    reflected = sections[sections.length - 2];
  }

  return reflected.dateVal;
}

/**
 * Custom function used by draggable Column C formula.
 * Returns the row number that the hyperlink should jump to (the latest section).
 * @customfunction
 */
function GET_EXERCISE_ROW(sheetName, dummy) {
  if (!sheetName) return 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(String(sheetName).trim());
  if (!sheet) return 1;

  const sections = getExerciseSections(sheet);
  if (sections.length === 0) return 1;

  // Hyperlink always jumps to the newest section (even if unscheduled)
  return sections[sections.length - 1].dateRow;
}

/**
 * Restores the draggable dynamic formula down Column C of Dashboard.
 */
/**
 * Restores the draggable dynamic formulas down Column C and Column D of Dashboard.
 */
function restoreDraggableFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName("dashboard");
  if (!dash) return;
  const lastRow = dash.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const formulasC = [];
  const formulasD = [];

  for (let r = 2; r <= lastRow; r++) {
    // Draggable hyperlink + date formula in Column C
    formulasC.push([
      `=IF(ISBLANK(B${r}), "", HYPERLINK("#gid=" & getSheetGidByName(B${r}) & "&range=A" & GET_EXERCISE_ROW(B${r}), GET_EXERCISE_DATE(B${r})))`
    ]);
    // Optional draggable status formula in Column D if you want both draggable:
    // formulasD.push([`=IF(ISBLANK(B${r}), "", GET_EXERCISE_STATUS(B${r}))`]);
  }

  dash.getRange(2, 3, numRows, 1).setFormulas(formulasC);
  SpreadsheetApp.flush();
  syncDashboardColors();
  ss.toast("Dashboard dynamic formulas restored!", "Success", 4);
}

/**
 * Custom function to get latest status if dragging Column D.
 * @customfunction
 */
function GET_EXERCISE_STATUS(sheetName) {
  if (!sheetName) return "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(String(sheetName).trim());
  if (!sheet) return "";
  const sections = getExerciseSections(sheet);
  if (sections.length === 0) return "";
  const latest = sections[sections.length - 1];
  let reflected = latest;
  if (latest.statusVal === 'unscheduled' && sections.length > 1) {
    reflected = sections[sections.length - 2];
  }
  return reflected.statusVal;
}