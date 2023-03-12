const HEADER_ROW = 2;

const CURRENT_WEEK_SHEET = "Semaine courante";
const PAST_WEEK_SHEET = "Semaine précédente";

function createSpreadsheetOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('getTimesheetsData')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
}

function getTimesheetsData() {
  getPastWeek();
  getCurrentWeek();
}

function getCurrentWeek() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  const folderFiles = DriveApp.getFolderById("1Mp1N80Y6fKDj40ndFN99SLnmppB7Z2l5").getFiles();

  let data = [];
  while(folderFiles.hasNext()) {
    data = data.concat(getDataFromSpreadsheet(folderFiles.next(), CURRENT_WEEK_SHEET));
  }

  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_WEEK_SHEET);
  ws.getRange(HEADER_ROW + 1, 1, data.length, data[0]?.length).setValues(data);
}

function getPastWeek() {
  ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  const folderFiles = DriveApp.getFolderById("1Mp1N80Y6fKDj40ndFN99SLnmppB7Z2l5").getFiles();

  let data = [];
  while(folderFiles.hasNext()) {
    data = data.concat(getDataFromSpreadsheet(folderFiles.next(), PAST_WEEK_SHEET));
  }

  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAST_WEEK_SHEET);
  ws.getRange(HEADER_ROW + 1, 1, data.length, data[0]?.length).setValues(data);
}

function getDataFromSpreadsheet(file, sheetName) {
  if (file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
    const ss = SpreadsheetApp.openById(file.getId());
    const ws = ss.getSheetByName(sheetName);
    const data = ws.getRange(HEADER_ROW + 1, 1, ws.getLastRow(), 4).getValues().filter(value => value[0] !== "");
    
    return data.map(value => {
      value.push(file.getName());
      return value;
    });
  }
}
