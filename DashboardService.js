/**
 * Dashboard & Workout Block Service
 * File: DashboardService.gs
 */

function getBlockAroundRow(sheet, row) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const searchLimit = 50;

  let topMarkerRow = -1;
  let bottomMarkerRow = -1;
  let bottomMarkerCol = -1;

  const minRow = Math.max(1, row - searchLimit);
  for (let r = row; r >= minRow; r--) {
    const rowValues = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const foundColIndex = rowValues.findIndex(val => val.trim() === "__");
    if (foundColIndex !== -1) {
      topMarkerRow = r;
      break;
    }
  }

  const maxRow = Math.min(lastRow, row + searchLimit);
  for (let r = row; r <= maxRow; r++) {
    const rowValues = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const foundColIndex = rowValues.findIndex(val => val.trim() === "__");
    if (foundColIndex !== -1 && r > topMarkerRow) {
      bottomMarkerRow = r;
      bottomMarkerCol = foundColIndex + 1;
      break;
    }
  }

  if (topMarkerRow === -1 || bottomMarkerRow === -1) {
    return null;
  }

  return {
    startRow: topMarkerRow,
    startCol: 1,
    endRow: bottomMarkerRow,
    endCol: bottomMarkerCol > 0 ? bottomMarkerCol : lastCol,
    numRows: bottomMarkerRow - topMarkerRow + 1,
    numCols: bottomMarkerCol > 0 ? bottomMarkerCol : lastCol
  };
}

function getExerciseSections(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const colA = sheet.getRange(1, 1, lastRow, 1).getValues().flat();
  const sections = [];

  for (let r = 0; r < colA.length - 1; r++) {
    const val = colA[r];
    const nextVal = String(colA[r + 1] || '').trim().toLowerCase();

    const isDate = (val instanceof Date && !isNaN(val)) || (typeof val === 'number' && val > 0);
    const isStatus = ['completed', 'planned', 'missed', 'unscheduled'].includes(nextVal);

    if (isDate && isStatus) {
      sections.push({
        dateRow: r + 1,
        statusRow: r + 2,
        dateVal: val,
        statusVal: nextVal
      });
    }
  }

  return sections;
}

function getLatestDateRow(sheet) {
  const sections = getExerciseSections(sheet);
  if (sections.length === 0) return -1;
  return sections[sections.length - 1].dateRow;
}

function columnToLetter(colIndex) {
  let temp, letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIndex = Math.floor((colIndex - temp - 1) / 26);
  }
  return letter;
}

/**
 * Adjusts relative row coordinates in cell formulas when simulating a cloned block.
 */
function shiftFormulaRows(formula, rowDelta) {
  if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) return formula;
  return formula.replace(/(\$?[A-Za-z]+)(\$?)(\d+)/g, function(match, colPart, dollar, rowNum) {
    if (dollar === '$') {
      return match;
    }
    const shifted = parseInt(rowNum, 10) + rowDelta;
    return colPart + shifted;
  });
}

/**
 * Extracts section data and formulas for dynamic client calculation.
 * If the latest section is completed, creates a virtual planned section with row offsets shifted.
 */
function getSectionForModification(sheetName, targetDateStr) {
  if (sheetName.toLowerCase() === 'treadmill') {
    return { error: 'Treadmill exercise is not currently supported for custom modification.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { error: `Sheet "${sheetName}" not found.` };
  }

  const sections = getExerciseSections(sheet);
  if (sections.length === 0) {
    return { error: `Sheet "${sheetName}" has no existing workout blocks.` };
  }

  const latestSection = sections[sections.length - 1];
  const hasUncompletedTail = ['planned', 'missed', 'unscheduled'].includes(latestSection.statusVal);

  let editSection = latestSection;
  let baselineSection = null;

  if (hasUncompletedTail) {
    editSection = latestSection;
    for (let i = sections.length - 2; i >= 0; i--) {
      if (sections[i].statusVal === 'completed') {
        baselineSection = sections[i];
        break;
      }
    }
  } else {
    // Latest section is completed -> it becomes both the template and baseline
    editSection = latestSection;
    baselineSection = latestSection;
  }

  const editBlock = getBlockAroundRow(sheet, editSection.dateRow);
  if (!editBlock) {
    return { error: `Could not identify block bounds for "${sheetName}".` };
  }

  const range = sheet.getRange(editBlock.startRow, editBlock.startCol, editBlock.numRows, editBlock.numCols);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const displays = range.getDisplayValues();

  // If virtualizing a new block, the start row shifts right below the current block
  const isVirtual = !hasUncompletedTail;
  const virtualRowDelta = isVirtual ? editBlock.numRows : 0;
  const effectiveStartRow = editBlock.startRow + virtualRowDelta;

  // External cell lookup map
  const allRange = sheet.getDataRange();
  const allValues = allRange.getValues();
  const allDisplays = allRange.getDisplayValues();
  const externalLookup = {};

  for (let r = 0; r < allValues.length; r++) {
    for (let c = 0; c < allValues[r].length; c++) {
      const a1 = columnToLetter(c + 1) + (r + 1);
      const v = allValues[r][c];
      const d = allDisplays[r][c];
      externalLookup[a1] = (v instanceof Date) ? d : v;
    }
  }

  let baseDisplays = null;
  if (baselineSection && (baselineSection.dateRow !== editSection.dateRow || isVirtual)) {
    const baseBlock = getBlockAroundRow(sheet, baselineSection.dateRow);
    if (baseBlock) {
      baseDisplays = sheet.getRange(baseBlock.startRow, baseBlock.startCol, baseBlock.numRows, baseBlock.numCols).getDisplayValues();
    }
  }

  const cells = [];
  for (let r = 0; r < editBlock.numRows; r++) {
    const rowCells = [];
    const absRow = effectiveStartRow + r;

    for (let c = 0; c < editBlock.numCols; c++) {
      const absCol = editBlock.startCol + c;
      const cellRef = columnToLetter(absCol) + absRow;
      let val = values[r][c];
      let form = formulas[r][c];
      let disp = displays[r][c];

      // Shift formula row indices if virtual block
      if (isVirtual && form) {
        form = shiftFormulaRows(form, virtualRowDelta);
      }

      // Virtualize Date row
      if (isVirtual && r === 1 && c === 0) {
        val = targetDateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        disp = val;
      }

      // Virtualize Status row
      if (isVirtual && r === 2 && c === 0) {
        val = "planned";
        disp = "planned";
      }

      let isFunctional = false;

      // 1. Formula cells
      if (form && form.trim().startsWith('=')) {
        isFunctional = true;
      }
      // 2. Section markers
      else if (disp.trim() === '__') {
        isFunctional = true;
      }
      // 3. Top boundary row
      else if (r === 0) {
        isFunctional = true;
      }
      // 4. Date & Checkbox row
      else if (r === 1 && (c === 0 || c === 1 || c === editBlock.numCols - 1)) {
        isFunctional = true;
      }
      // 5. Status dropdown row
      else if (r === 2 && c === 0) {
        isFunctional = true;
      }
      // 6. Set numbers
      else if (c === 1 && /^\d+$/.test(disp.trim()) && r >= 3 && r < editBlock.numRows - 1) {
        isFunctional = true;
      }

      let isDiff = false;
      if (!isVirtual && baseDisplays && r < baseDisplays.length && c < baseDisplays[0].length) {
        const baseDisp = baseDisplays[r][c];
        if (disp.trim() !== baseDisp.trim() && !isFunctional) {
          isDiff = true;
        }
      }

      rowCells.push({
        rOffset: r,
        cOffset: c,
        absRow: absRow,
        absCol: absCol,
        cellRef: cellRef,
        val: val instanceof Date ? disp : val,
        baseVal: val instanceof Date ? disp : val,
        display: disp,
        formula: form || '',
        isFunctional: isFunctional,
        isDiff: isDiff
      });
    }
    cells.push(rowCells);
  }

  return {
    success: true,
    sheetName: sheetName,
    startRow: effectiveStartRow,
    startCol: editBlock.startCol,
    numRows: editBlock.numRows,
    numCols: editBlock.numCols,
    cells: cells,
    externalLookup: externalLookup,
    hasUncompletedTail: hasUncompletedTail,
    sectionStatus: isVirtual ? "planned" : editSection.statusVal
  };
}

function syncDashboardColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("dashboard");
  if (!summarySheet) return;

  const lastRow = summarySheet.getLastRow();
  if (lastRow < 2) return;
  const numRows = lastRow - 1;

  const exerciseNames = summarySheet.getRange(2, 2, numRows, 1).getValues();
  const formulasC = [];
  const colorsC = [];
  const valuesD = [];
  const colorsD = [];

  for (let i = 0; i < numRows; i++) {
    const rowNum = i + 2;
    const exerciseName = String(exerciseNames[i][0] || '').trim();
    if (!exerciseName) {
      formulasC.push([""]);
      colorsC.push(["#000000"]);
      valuesD.push([""]);
      colorsD.push(["#000000"]);
      continue;
    }

    const targetSheet = ss.getSheetByName(exerciseName);
    if (!targetSheet) {
      formulasC.push([summarySheet.getRange(rowNum, 3).getFormula() || ""]);
      colorsC.push(["#000000"]);
      valuesD.push([summarySheet.getRange(rowNum, 4).getValue() || ""]);
      colorsD.push(["#000000"]);
      continue;
    }

    const sections = getExerciseSections(targetSheet);
    if (sections.length === 0) {
      formulasC.push([summarySheet.getRange(rowNum, 3).getFormula() || ""]);
      colorsC.push(["#000000"]);
      valuesD.push([summarySheet.getRange(rowNum, 4).getValue() || ""]);
      colorsD.push(["#000000"]);
      continue;
    }

    const latestSection = sections[sections.length - 1];
    let reflectedSection = latestSection;
    if (latestSection.statusVal === 'unscheduled' && sections.length > 1) {
      reflectedSection = sections[sections.length - 2];
    }

    const targetColor = STATUS_COLORS[reflectedSection.statusVal] || "#000000";

    // Dynamic formula using INDIRECT so manual sheet entries mirror automatically
    formulasC.push([
      `=LET(name, B${rowNum}, range, INDIRECT("'" & name & "'!A:A"), maxVal, MAX(range), IF(maxVal = 0, "", HYPERLINK("#gid=" & getSheetGidByName(name) & "&range=A" & MATCH(maxVal, range, 0), TEXT(maxVal, "yyyy-mm-dd"))))`
    ]);
    colorsC.push([targetColor]);
    valuesD.push([reflectedSection.statusVal]);
    colorsD.push([targetColor]);

    targetSheet.getRange(reflectedSection.dateRow, 1).setFontColor(targetColor);
    targetSheet.getRange(reflectedSection.statusRow, 1).setFontColor(targetColor);
  }

  const rangeC = summarySheet.getRange(2, 3, numRows, 1);
  rangeC.setFormulas(formulasC);
  rangeC.setFontColors(colorsC);
  rangeC.setFontLine("none");

  const rangeD = summarySheet.getRange(2, 4, numRows, 1);
  rangeD.setValues(valuesD);
  rangeD.setFontColors(colorsD);

  SpreadsheetApp.flush();
}