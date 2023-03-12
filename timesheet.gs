const HEADER_ROW = 2;
const LAST_MODIFIED_CELL = [1, 8];
const FROM_DATE_CELL = [1, 9];
const TO_DATE_CELL = [1, 10];
const TASK_STARTED_CELL = [2, 8];

const CLIENT_COLUMN = 1;
const TASK_DATE_COLUMN = 2;
const TASK_TIME_COLUMN = 3;
const DESCRIPTION_COLUMN = 4;
const TASK_BUTTON_COLUMN = 5;
const TASK_STATUS_COLUMN = 6;
const TASK_EVENT_TIME_COLUMN = 7;
const TASK_START_TIME_COLUMN = 8;

const CURRENT_WEEK_SHEET = "Semaine courante";
const PREVIOUS_WEEK_SHEET = "Semaine précédente";


function onOpen(e) {
  const sheet = e.source.getSheetByName(CURRENT_WEEK_SHEET);
  const currentTime = new Date(Date.now());
  sheet.getRange(LAST_MODIFIED_CELL[0], LAST_MODIFIED_CELL[1]).setValue(currentTime.getTime());
  sheet.getRange(LAST_MODIFIED_CELL[0], LAST_MODIFIED_CELL[1] - 1).setValue(currentTime.toLocaleDateString());
}

function onEdit(e) {
  if (e?.range.getRow() <= HEADER_ROW) return;

  switch(e?.range.getColumn()) {
    case CLIENT_COLUMN:
      highlightIfNotFound(e);
      break;
    case TASK_DATE_COLUMN:
      validateDate(e);
      break;
    case TASK_BUTTON_COLUMN:
      setWorkingStatus(e);
      break;
  }
}

function onNewWeek() {
  const ss = SpreadsheetApp.openById("1bO_5pYJd2zhz_wp2AlduGd3BpKxy8eUtRXcYx4Xolag")
  const week = ss.getSheetByName(CURRENT_WEEK_SHEET);
  const previousWeek = ss.getSheetByName(PREVIOUS_WEEK_SHEET);
  
  setWeeks(ss);

  updateHistory(ss.getSheetByName("Historique"), previousWeek.getRange(HEADER_ROW + 1, 1, previousWeek.getLastRow(), DESCRIPTION_COLUMN).getValues()
    .filter(value => value[0] !== ""));
  updatePreviousWeek(previousWeek, week, week.getRange(HEADER_ROW + 1, 1, week.getLastRow(), DESCRIPTION_COLUMN).getValues()
    .filter(value => value[0] !== ""));
}

function validateDate(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range.getCell(1, 1);
  if (!isValidDate(sheet, cell.getValue().getTime())) {
    cell.clearContent();
    cell.setBackground("#ff0000");
  } else {
    cell.setBackground("#ffffff");
  }
}

function isValidDate(sheet, dateNb) {
  return (dateNb >= sheet.getRange(FROM_DATE_CELL[0], FROM_DATE_CELL[1]).getValue()) && 
    (dateNb <= sheet.getRange(TO_DATE_CELL[0], TO_DATE_CELL[1]).getValue());
}

function setWeeks(ss) {
  const week = ss.getSheetByName(CURRENT_WEEK_SHEET);
  const previousWeek = ss.getSheetByName(PREVIOUS_WEEK_SHEET);
  const fromDate = week.getRange(FROM_DATE_CELL[0], FROM_DATE_CELL[1]);
  const toDate = week.getRange(TO_DATE_CELL[0], TO_DATE_CELL[1]);
  previousWeek.getRange(FROM_DATE_CELL[0], FROM_DATE_CELL[1]).setValue(fromDate.getValue());
  previousWeek.getRange(TO_DATE_CELL[0], TO_DATE_CELL[1]).setValue(toDate.getValue());
  fromDate.setValue(Date.now());
  toDate.setValue(Date.now() + (7 * 24 * 60 * 60 * 1000));
}

function updatePreviousWeek(previousWeek, week, data) {
  week.getRange(HEADER_ROW + 1, CLIENT_COLUMN, week.getLastRow(), DESCRIPTION_COLUMN).clearContent();
  week.getRange(HEADER_ROW + 1, TASK_BUTTON_COLUMN, week.getLastRow(), TASK_START_TIME_COLUMN).clearContent();

  previousWeek.getRange(HEADER_ROW - 1, TASK_DATE_COLUMN).setValue(week.getRange(HEADER_ROW - 1, TASK_DATE_COLUMN).getValue());
  const now = new Date(Date.now());
  const month = now.getMonth() < 9 ? "0" + (now.getMonth() + 1) : (now.getMonth() + 1).toString();
  const date = now.getDate() < 9 ? "0" + (now.getDate()) : now.getDate().toString();
  week.getRange(HEADER_ROW - 1, TASK_DATE_COLUMN).setValue(`Feuille de temps - Semaine du ${now.getFullYear()}-${month}-${date}`);

  previousWeek.getRange(HEADER_ROW + 1, 1, previousWeek.getLastRow(), DESCRIPTION_COLUMN).clearContent();
  if (data.length > 0) {
    previousWeek.getRange(HEADER_ROW + 1, 1, data.length, data[0].length).setValues(data);
  }
}

function updateHistory(history, data) {
  if (data.length > 0) {
    history.getRange(history.getLastRow() + 1, CLIENT_COLUMN, data.length, data[0].length).setValues(data);
  }
}

function highlightIfNotFound(e) {
  const cell = e.range.getCell(1, 1);
  if (cell.getValue() === "Introuvable") {
    cell.setBackground("#ffff33");
  } else {
    cell.setBackground("#ffffff");
  }
}

function setWorkingStatus(e) {
  const sheet = e.source.getSheetByName(CURRENT_WEEK_SHEET);

  if (sheet.getRange(e.range.getRow(), CLIENT_COLUMN).getValue() === "") {
    sheet.getRange(e.range.getRow(), CLIENT_COLUMN).setBackground("#ffff33");
    e.range.getCell(1, 1).setValue(false);
    return;
  }

  if (e.range.getCell(1, 1).getValue()) {
    if (getTaskInProgress(sheet)) {
      e.range.getCell(1, 1).setValue(false);
    } else if (isNotToday(sheet.getRange(e.range.getRow(), TASK_DATE_COLUMN).getValue())) {
      e.range.getCell(1, 1).setValue(false);
      copyTask(sheet, e.range.getRow());
      startTask(sheet, e.range.getRow() + 1, e.range.getColumn());
    } else {
      startTask(sheet, e.range.getRow(), e.range.getColumn());
    }
  } else {
    stopTask(sheet, e.range.getRow(), e.range.getColumn());
  }
}

function isNotToday(date) {
  if (!date) return false;
  return new Date(Date.now()).getDate() !== date.getDate();
}

function copyTask(sheet, oldRow) {
  const task = sheet.getRange(oldRow, 1, 1, 4).getValues();
  sheet.insertRowAfter(oldRow);
  sheet.getRange(oldRow + 1, 1, 1, 4).setValues(task);
  sheet.getRange(oldRow + 1, TASK_BUTTON_COLUMN).setValue(true);
}

function getTaskInProgress(sheet) {
  if (sheet) {
    return sheet.getRange(TASK_STARTED_CELL[0], TASK_STARTED_CELL[1]).getValue();
  }
  return false;
}

function setTaskInProgress(sheet, value) {
  if (sheet) {
    const cell = sheet.getRange(HEADER_ROW, TASK_BUTTON_COLUMN);
    if (value) {
      cell.setValue("Travail en cours");
      cell.setBackground("#ff0000");
      sheet.getRange(HEADER_ROW, TASK_BUTTON_COLUMN + 1, 1, 2).setBackground("#ff6666");
    } else {
      cell.setValue("Commencer");
      cell.setBackground("#00cc00");
      sheet.getRange(HEADER_ROW, TASK_BUTTON_COLUMN + 1, 1, 2).setBackground("#33ff99");
    }
    sheet.getRange(TASK_STARTED_CELL[0], TASK_STARTED_CELL[1]).setValue(value);
  }
}

function startTask(sheet, row, column) {
  if (sheet) {
    setTaskInProgress(sheet, true);
    sheet.getRange(row, column).setBackground("#99FFFF");
    sheet.getRange(row, TASK_STATUS_COLUMN, 1, 2).setBackground("#ccffff");
    sheet.getRange(row, TASK_STATUS_COLUMN).setValue("Commencé à");

    const currentTime = new Date(Date.now());
    sheet.getRange(row, column - 3).setValue(currentTime.toLocaleDateString());
    sheet.getRange(row, TASK_EVENT_TIME_COLUMN).setValue(currentTime.toLocaleTimeString());
    sheet.getRange(row, TASK_START_TIME_COLUMN).setValue(currentTime.getTime());
  }
}

function stopTask(sheet, row, column) {
  if (sheet) {
    setTaskInProgress(sheet, false);
    sheet.getRange(row, column, 1, 3).setBackground("#FFFFFF");
    sheet.getRange(row, TASK_STATUS_COLUMN).setValue("Terminé à");

    const currentTime = new Date(Date.now());
    sheet.getRange(row, TASK_EVENT_TIME_COLUMN).setValue(currentTime.toLocaleTimeString());

    const startTime = sheet.getRange(row, TASK_START_TIME_COLUMN).getValue();
    const workingTime = (currentTime.getTime() - startTime) / (24 * 60 * 1000);
    const previousWorkingTime = sheet.getRange(row, TASK_TIME_COLUMN).getValue();
    sheet.getRange(row, TASK_TIME_COLUMN).setValue(Math.round((previousWorkingTime + workingTime) * 10) / 10);
  }
}
