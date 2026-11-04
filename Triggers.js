/**
 * Desktop top-level menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏋️ Workout')
    .addItem('Select Exercises...', 'showWorkoutDialog')
    .addItem('Sync Dashboard Colors', 'syncDashboardColors')
    .addSeparator()
    .addItem('↩️ Undo / Reverse Workout...', 'showUndoDialog')
    .addToUi();
}

/**
 * Master onEdit handler:
 * - Duplicate / remove workout block upon ticking/unticking Date checkbox.
 * - Manual override in Dashboard Column D writes to the target sheet (respecting unscheduled bypass).
 * - Manual override in Exercise Sheet Column A reflects to Dashboard Column D.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const activeSheet = range.getSheet();
  const sheetName = activeSheet.getName();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const ss = e.source;

  // -------------------------------------------------------------
  // TRIGGER: Workout Section Checkbox Toggle
  // -------------------------------------------------------------
  if (sheetName.toLowerCase() !== "dashboard") {
    const block = getBlockAroundRow(activeSheet, editedRow);
    if (block && editedRow === block.startRow + 1 && editedCol === block.endCol) {
      const isChecked = (e.value === "TRUE" || range.getValue() === true);
      if (isChecked) {
        range.insertCheckboxes();
        range.setValue(true);
        const targetStartRow = block.endRow + 1;
        
        activeSheet.insertRowsAfter(block.endRow, block.numRows);
        const sourceRange = activeSheet.getRange(block.startRow, block.startCol, block.numRows, block.numCols);
        const targetRange = activeSheet.getRange(targetStartRow, block.startCol);
        sourceRange.copyTo(targetRange);

        const newCheckboxCell = activeSheet.getRange(targetStartRow + 1, block.endCol);
        newCheckboxCell.clearDataValidations();
        newCheckboxCell.insertCheckboxes();
        newCheckboxCell.setValue(false);

        const newDateCell = activeSheet.getRange(targetStartRow + 1, block.startCol);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        newDateCell.setValue(today);
        newDateCell.setNumberFormat('dd/mm/yy');

        const newStatusCell = activeSheet.getRange(targetStartRow + 2, block.startCol);
        newStatusCell.setValue("unscheduled");
        newStatusCell.setFontColor(STATUS_COLORS["unscheduled"] || "#7c02ff");
        newDateCell.setFontColor(STATUS_COLORS["unscheduled"] || "#7c02ff");

        SpreadsheetApp.flush();
        activeSheet.setActiveRange(newDateCell);
        ss.toast("Workout duplicated below!", "Success", 3);

        // Update dashboard so it retains next-to-last section or reflects properly
        syncDashboardColors();
        return;
      } else {
        range.insertCheckboxes();
        range.setValue(false);
        const nextBlockStart = block.endRow + 1;
        const nextBlock = getBlockAroundRow(activeSheet, nextBlockStart + 1);
        if (nextBlock && nextBlock.startRow === nextBlockStart) {
          activeSheet.deleteRows(nextBlock.startRow, nextBlock.numRows);
          SpreadsheetApp.flush();
          ss.toast("Subsequent workout block removed.", "Info", 3);
          syncDashboardColors();
          return;
        }
      }
    }
  }

  // -------------------------------------------------------------
  // CASE 1: Manual Edit on 'dashboard' (Column D: Status dropdown)
  // -------------------------------------------------------------
  if (sheetName === "dashboard" && editedCol === 4 && editedRow >= 2) {
    const newStatus = String(e.value || "").trim().toLowerCase();
    const exerciseName = String(activeSheet.getRange(editedRow, 2).getValue()).trim();
    if (!exerciseName) return;

    const targetSheet = ss.getSheetByName(exerciseName);
    if (!targetSheet) return;

    const sections = getExerciseSections(targetSheet);
    if (sections.length === 0) return;

    const latestSection = sections[sections.length - 1];
    let targetSection = latestSection;

    // If latest section is unscheduled and a prior section exists, edit next-to-last section
    if (latestSection.statusVal === 'unscheduled' && sections.length > 1) {
      targetSection = sections[sections.length - 2];
    }

    // Write new status to the target section in exercise sheet
    targetSheet.getRange(targetSection.statusRow, 1).setValue(e.value);

    const targetColor = STATUS_COLORS[newStatus] || "#000000";
    targetSheet.getRange(targetSection.dateRow, 1).setFontColor(targetColor);
    targetSheet.getRange(targetSection.statusRow, 1).setFontColor(targetColor);

    // Apply color to Dashboard Date and Status cells
    const dashLinkCell = activeSheet.getRange(editedRow, 3);
    dashLinkCell.setFontColor(targetColor);
    dashLinkCell.setFontLine("none");
    range.setFontColor(targetColor);
    return;
  }

  // -------------------------------------------------------------
  // CASE 2: Manual Edit on an individual Exercise Sheet (Column A)
  // -------------------------------------------------------------
  if (sheetName.toLowerCase() !== "dashboard" && editedCol === 1) {
    syncDashboardColors();
    return;
  }
}