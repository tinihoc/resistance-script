/**
 * Daily Exercise Selector
 * File: DailyExerciseSelector.gs
 */

const MAX_SNAPSHOTS = 5;
const PROP_SNAPSHOT_KEY = 'WORKOUT_SNAPSHOTS_STACK';

function showWorkoutDialog() {
  // Starts with clean compact defaults; Dialog.html resizes dynamically based on content
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Today's Exercises");
}

function showUndoDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UndoDialog')
    .setWidth(480)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, "Undo Previous Confirmation");
}

function runSelection(dayType, dateStr, targetMuscles, excludedExercises = [], excludedMuscles = []) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('dashboard') || ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { error: 'No data found in dashboard sheet.' };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const exExSet = new Set((excludedExercises || []).map(e => String(e).trim().toLowerCase()));
  const exMuSet = new Set((excludedMuscles || []).map(m => String(m).trim().toLowerCase()));

  const lowerBodyMuscles = ['calves', 'hammies', 'quads', 'inner thighs', 'glutes', 'lower back'];
  const lowerSet = new Set(lowerBodyMuscles);

  let rawExercises = [];
  for (let i = 0; i < values.length; i++) {
    const rowIdx = i + 2;
    const category = String(values[i][0] || '').trim().toLowerCase();
    const name = String(values[i][1] || '').trim();
    const rawDate = values[i][2];
    const status = String(values[i][3] || '').trim().toLowerCase();
    const rawMuscles = String(values[i][4] || '').trim();

    if (!name) continue;

    let parsedDateStr = '';
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      parsedDateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (rawDate) {
      parsedDateStr = String(rawDate).trim();
    }

    const muscles = rawMuscles
      .split(',')
      .map(m => m.trim().toLowerCase())
      .filter(m => m.length > 0);

    rawExercises.push({
      rowIndex: rowIdx,
      category: category,
      name: name,
      lastSession: parsedDateStr,
      status: status,
      muscles: muscles
    });
  }

  let dayBeforeStr = '';
  if (dateStr) {
    const parts = dateStr.split('-');
    if (parts.length === 3) {
      const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
      d.setDate(d.getDate() - 1);
      dayBeforeStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  const lastTrainedLowerMuscles = new Set();
  if (dayBeforeStr) {
    for (const ex of rawExercises) {
      if (ex.status === 'completed' && ex.lastSession === dayBeforeStr) {
        for (const m of ex.muscles) {
          if (lowerSet.has(m)) {
            lastTrainedLowerMuscles.add(m);
          }
        }
      }
    }
  }

  let exercises = rawExercises.filter(e => {
    if (exExSet.has(e.name.toLowerCase())) return false;
    if (e.muscles.some(m => exMuSet.has(m))) return false;
    if (dateStr && e.lastSession && e.lastSession > dateStr) return false;
    return true;
  });

  exercises.sort((a, b) => {
    const isMissedA = a.status === 'missed' ? 0 : 1;
    const isMissedB = b.status === 'missed' ? 0 : 1;
    if (isMissedA !== isMissedB) return isMissedA - isMissedB;

    const dateA = a.lastSession || '9999-99-99';
    const dateB = b.lastSession || '9999-99-99';
    if (dateA !== dateB) return dateA.localeCompare(dateB);

    return b.rowIndex - a.rowIndex;
  });

  function attemptSelection(allowedCategories, allowMuscleOverlap, allowBackConflict, allowMultipleLower, allowLastLowerConflict) {
    const selected = [];
    const targetedMuscles = new Set();
    let hasLowerBody = false;
    let hasLowerBack = false;
    let hasUpperBack = false;

    for (const ex of exercises) {
      if (!allowedCategories.includes(ex.category)) continue;

      const isLower = ex.muscles.some(m => lowerSet.has(m));
      if (!allowMultipleLower && isLower && hasLowerBody) {
        continue;
      }

      if (!allowLastLowerConflict && isLower) {
        const conflictsWithLast = ex.muscles.some(m => lowerSet.has(m) && lastTrainedLowerMuscles.has(m));
        if (conflictsWithLast) {
          continue;
        }
      }

      const newMuscles = ex.muscles.filter(m => !targetedMuscles.has(m));
      if (!allowMuscleOverlap) {
        if (newMuscles.length !== ex.muscles.length) continue;
      }

      if (targetedMuscles.size + newMuscles.length > targetMuscles) {
        continue;
      }

      const isLb = ex.muscles.includes('lower back');
      const isUb = ex.muscles.includes('upper back');
      if (!allowBackConflict && dayType === 'heavy') {
        if (isLb && hasUpperBack) continue;
        if (isUb && hasLowerBack) continue;
      }

      selected.push(ex);
      newMuscles.forEach(m => targetedMuscles.add(m));
      if (isLower) hasLowerBody = true;
      if (isLb) hasLowerBack = true;
      if (isUb) hasUpperBack = true;

      if (targetedMuscles.size === targetMuscles) break;
    }

    return {
      selected: selected,
      targetedMuscles: Array.from(targetedMuscles),
      count: targetedMuscles.size
    };
  }

  const baseCategories = dayType === 'heavy' ? ['heavy', 'moderate'] : ['light', 'moderate'];
  const allCategories = ['heavy', 'moderate', 'light'];

  const stages = [
    { name: 'Strict rules', cats: baseCategories, overlap: false, backConflict: false, multiLower: false, lastLowerConflict: false },
    { name: 'Relaxed: Allowed muscle overlap', cats: baseCategories, overlap: true, backConflict: false, multiLower: false, lastLowerConflict: false },
    { name: 'Relaxed: Allowed lower muscles from day before', cats: baseCategories, overlap: true, backConflict: false, multiLower: false, lastLowerConflict: true },
    { name: 'Relaxed: Allowed multiple lower body & back overlap', cats: baseCategories, overlap: true, backConflict: true, multiLower: true, lastLowerConflict: true },
    { name: 'Relaxed: Allowed all categories pool', cats: allCategories, overlap: true, backConflict: true, multiLower: true, lastLowerConflict: true }
  ];

  let result = null;
  let relaxedNote = '';

  for (let i = 0; i < stages.length; i++) {
    const stage = stages[i];
    result = attemptSelection(stage.cats, stage.overlap, stage.backConflict, stage.multiLower, stage.lastLowerConflict);
    if (result.count === targetMuscles) {
      if (i > 0) relaxedNote = stage.name;
      break;
    }
  }

  return {
    selected: result.selected,
    muscles: result.targetedMuscles,
    target: targetMuscles,
    relaxedNote: relaxedNote,
    isPartial: result.count < targetMuscles
  };
}

function getSnapshotsList() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(PROP_SNAPSHOT_KEY);
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch (e) {
    return [];
  }
}

function pushSnapshot(snapshot) {
  const props = PropertiesService.getDocumentProperties();
  let list = getSnapshotsList();
  list.unshift(snapshot);
  if (list.length > MAX_SNAPSHOTS) {
    list = list.slice(0, MAX_SNAPSHOTS);
  }
  props.setProperty(PROP_SNAPSHOT_KEY, JSON.stringify(list));
}

function removeSnapshotById(id) {
  const props = PropertiesService.getDocumentProperties();
  let list = getSnapshotsList().filter(s => s.id !== id);
  props.setProperty(PROP_SNAPSHOT_KEY, JSON.stringify(list));
}

function confirmWorkoutSelection(selectedExercises, dateStr, customEditsByEx = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName("dashboard");
  if (!dashboardSheet) throw new Error("Dashboard sheet not found.");

  const tz = Session.getScriptTimeZone();
  const parts = dateStr.split('-');
  const targetDateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  const timeFormatted = Utilities.formatDate(new Date(), tz, "hh:mm a");

  const namesSummary = selectedExercises.map(e => e.name).slice(0, 3).join(', ') + (selectedExercises.length > 3 ? '...' : '');
  const snapshotTitle = `${dateStr}   ${timeFormatted} (${selectedExercises.length} ex: ${namesSummary})`;

  const snapshot = {
    id: Utilities.getUuid(),
    title: snapshotTitle,
    timestamp: new Date().toISOString(),
    exercises: []
  };

  for (const item of selectedExercises) {
    const exName = item.name;
    const targetSheet = ss.getSheetByName(exName);
    if (!targetSheet) {
      throw new Error(`Exercise sheet "${exName}" not found.`);
    }

    const maxDateRow = getLatestDateRow(targetSheet);
    if (maxDateRow === -1) {
      throw new Error(`Sheet "${exName}" has no existing workout blocks.`);
    }

    const currentStatus = String(targetSheet.getRange(maxDateRow + 1, 1).getValue()).trim().toLowerCase();
    const currentDateVal = targetSheet.getRange(maxDateRow, 1).getValue();
    const customEdits = customEditsByEx[exName] || [];

    if (currentStatus === "completed") {
      const block = getBlockAroundRow(targetSheet, maxDateRow);
      if (!block) throw new Error(`Could not find bounding markers for "${exName}".`);

      const targetStartRow = block.endRow + 1;
      targetSheet.insertRowsAfter(block.endRow, block.numRows);
      
      const sourceRange = targetSheet.getRange(block.startRow, block.startCol, block.numRows, block.numCols);
      const targetRange = targetSheet.getRange(targetStartRow, block.startCol);
      sourceRange.copyTo(targetRange);

      const prevCheckbox = targetSheet.getRange(block.startRow + 1, block.endCol);
      prevCheckbox.insertCheckboxes();
      prevCheckbox.setValue(true);

      const newCheckbox = targetSheet.getRange(targetStartRow + 1, block.endCol);
      newCheckbox.clearDataValidations();
      newCheckbox.insertCheckboxes();
      newCheckbox.setValue(false);

      const newDateCell = targetSheet.getRange(targetStartRow + 1, block.startCol);
      newDateCell.setValue(targetDateObj);
      newDateCell.setNumberFormat('dd/mm/yy');

      const newStatusCell = targetSheet.getRange(targetStartRow + 2, block.startCol);
      newStatusCell.setValue("planned");
      newStatusCell.setFontColor(STATUS_COLORS["planned"] || "#1b73e8");
      newDateCell.setFontColor(STATUS_COLORS["planned"] || "#1b73e8");

      if (customEdits && customEdits.length > 0) {
        for (const edit of customEdits) {
          const writeRow = targetStartRow + edit.rOffset;
          const writeCol = block.startCol + edit.cOffset;
          targetSheet.getRange(writeRow, writeCol).setValue(edit.value);
        }
      }

      snapshot.exercises.push({
        name: exName,
        action: 'duplicated',
        createdStartRow: targetStartRow,
        numRows: block.numRows,
        prevCheckboxRow: block.startRow + 1,
        prevCheckboxCol: block.endCol
      });
    } else {
      const block = getBlockAroundRow(targetSheet, maxDateRow);
      const dateCell = targetSheet.getRange(maxDateRow, 1);
      const statusCell = targetSheet.getRange(maxDateRow + 1, 1);

      const overwrittenCells = [];
      if (customEdits && customEdits.length > 0 && block) {
        for (const edit of customEdits) {
          const writeRow = block.startRow + edit.rOffset;
          const writeCol = block.startCol + edit.cOffset;
          const targetCell = targetSheet.getRange(writeRow, writeCol);
          overwrittenCells.push({
            row: writeRow,
            col: writeCol,
            prevValue: targetCell.getValue()
          });
          targetCell.setValue(edit.value);
        }
      }

      snapshot.exercises.push({
        name: exName,
        action: 'updated',
        dateRow: maxDateRow,
        prevDate: currentDateVal instanceof Date ? Utilities.formatDate(currentDateVal, tz, 'yyyy-MM-dd') : currentDateVal,
        prevStatus: currentStatus,
        overwrittenCells: overwrittenCells
      });

      dateCell.setValue(targetDateObj);
      dateCell.setNumberFormat('dd/mm/yy');
      statusCell.setValue("planned");
      statusCell.setFontColor(STATUS_COLORS["planned"] || "#1b73e8");
      dateCell.setFontColor(STATUS_COLORS["planned"] || "#1b73e8");
    }
  }

  SpreadsheetApp.flush();
  pushSnapshot(snapshot);
  syncDashboardColors();
  sortDashboardStable(dashboardSheet);
  SpreadsheetApp.flush();

  return { success: true, snapshotId: snapshot.id, snapshotTitle: snapshot.title };
}

function reverseSnapshotById(snapshotId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const list = getSnapshotsList();
  if (list.length === 0) {
    return { error: 'No snapshots available to reverse.' };
  }

  const snapshot = snapshotId ? list.find(s => s.id === snapshotId) : list[0];
  if (!snapshot) {
    return { error: 'Snapshot not found.' };
  }

  for (const item of snapshot.exercises) {
    const targetSheet = ss.getSheetByName(item.name);
    if (!targetSheet) continue;

    if (item.action === 'duplicated') {
      targetSheet.deleteRows(item.createdStartRow, item.numRows);
      const prevCheckbox = targetSheet.getRange(item.prevCheckboxRow, item.prevCheckboxCol);
      prevCheckbox.insertCheckboxes();
      prevCheckbox.setValue(false);
    } else if (item.action === 'updated') {
      const dateCell = targetSheet.getRange(item.dateRow, 1);
      const statusCell = targetSheet.getRange(item.dateRow + 1, 1);

      if (item.prevDate) {
        if (/^\d{4}-\d{2}-\d{2}$/.test(item.prevDate)) {
          const parts = item.prevDate.split('-');
          dateCell.setValue(new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])));
        } else {
          dateCell.setValue(item.prevDate);
        }
        dateCell.setNumberFormat('dd/mm/yy');
      }

      const st = item.prevStatus || 'unscheduled';
      statusCell.setValue(st);
      const stColor = STATUS_COLORS[st] || "#000000";
      statusCell.setFontColor(stColor);
      dateCell.setFontColor(stColor);

      if (item.overwrittenCells && item.overwrittenCells.length > 0) {
        for (const cell of item.overwrittenCells) {
          targetSheet.getRange(cell.row, cell.col).setValue(cell.prevValue);
        }
      }
    }
  }

  SpreadsheetApp.flush();
  const dashboardSheet = ss.getSheetByName("dashboard");
  if (dashboardSheet) {
    syncDashboardColors();
    sortDashboardStable(dashboardSheet);
  }

  removeSnapshotById(snapshot.id);
  return { success: true, reversedTitle: snapshot.title };
}

function sortDashboardStable(dashboardSheet) {
  const lastRow = dashboardSheet.getLastRow();
  if (lastRow < 3) return;

  const numRows = lastRow - 1;
  const helperCol = 6;
  const helperRange = dashboardSheet.getRange(2, helperCol, numRows, 1);

  const indexValues = [];
  for (let i = 0; i < numRows; i++) {
    indexValues.push([i + 1]);
  }
  helperRange.setValues(indexValues);
  SpreadsheetApp.flush();

  const sortRange = dashboardSheet.getRange(2, 1, numRows, helperCol);
  sortRange.sort([
    { column: 3, ascending: false },
    { column: helperCol, ascending: true }
  ]);

  helperRange.clearContent();
  SpreadsheetApp.flush();
}