// Modification dès qu'un changement est effectué
function onEdit(e) {
  let rangeStart = 2;

  if (rangeStart < e.range.getRow()) {
    rangeStart = e.range.getRow();
  }

  for (let i = rangeStart; i <= e.range.getLastRow(); ++i) {
    setColor(i);
  }

  return 0;
}


// Vérification à l'ouverture du fichier chaque nouvelle journée.
function onOpen(e) {
  const colorColumn = 2;

  let ss = e.source.getSheetByName("Dossier reçu");
  let lastVerification = ss.getRange(1, colorColumn);

  if (lastVerification.getValue() === ""
    || lastVerification.getValue().split("> ")[1].split(" ")[2] !== Date().split(" ")[2] 
    || lastVerification.getValue().split("> ")[1].split(" ")[1] !== Date().split(" ")[1]) {
    for (let i = 2; i <= ss.getLastRow(); ++i) {
      setColor(i);
    }

    lastVerification.setFontColor("#ffffff")
    lastVerification.setValue("Last checked on -> " + Date());
  }

  return 0;
}



function setColor(rowNb) {
  const dateColumn = 1;
  const colorColumn = 2;

  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dossier reçu");
  let colorCell = ss.getRange(rowNb, colorColumn);

  let dateFromCell = ss.getRange(rowNb, dateColumn).getValue();
  let nbOfDaysPassed = (getNbOfDaysDifference(dateFromCell)); 

  if (nbOfDaysPassed <= 7) {
    colorCell.setBackground("#329c17");
    colorCell.setFontColor("#329c17");
    colorCell.setValue(nbOfDaysPassed + "j");
  }
  else if (nbOfDaysPassed <= 14) {
    colorCell.setBackground("#d46969");
    colorCell.setFontColor("#d46969");
    colorCell.setValue(nbOfDaysPassed + "j");
  }
  else if (nbOfDaysPassed <= 21) {
    colorCell.setBackground("#ff0000");
    colorCell.setFontColor("#ff0000");
    colorCell.setValue(nbOfDaysPassed + "j");
  }
  else if (nbOfDaysPassed <= 28) {
    colorCell.setBackground("#8f0e0e");
    colorCell.setFontColor("#8f0e0e");
    colorCell.setValue(nbOfDaysPassed + "j");
  }
  else if (nbOfDaysPassed < 999) {
    colorCell.setBackground("#000000");
    colorCell.setFontColor("#000000");
    colorCell.setValue(nbOfDaysPassed + "j");
  }
  else {
    colorCell.setBackground("#ffffff");
  }
  
  return true;
}


function getNbOfDaysDifference(date) {
  const month = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  let today = Date().split(" ");
  date = String(date).split(" ");

  if (date[3] === today[3]) {
    if (date[1] === today[1]) {
      return Number(today[2]) - Number(date[2]);
    }

    else if (month[(month.indexOf(date[1]) + 1) % 12] === today[1]) {
      return getNbOfDaysInMonth(month.indexOf(date[1]), date[3]) - Number(date[2]) + Number(today[2]);
    }

    else {
      let nbOfDays = getNbOfDaysInMonth(month.indexOf(date[1]), date[3]) - Number(date[2]) + Number(today[2]);

      for (let i = month.indexOf(date[1]) + 1; i !== month.indexOf(today[1]); ++i) {
        nbOfDays += getNbOfDaysInMonth(i, date[3]);
      }

      return nbOfDays;
    }
  }

  else if (Number(date[3]) + 1 === Number(today[3])) {
      let year = Number(date[3]);
      let nbOfDays = getNbOfDaysInMonth(month.indexOf(date[1]), date[3]) - Number(date[2]) + Number(today[2]);

      for (let i =  month.indexOf(date[1]) + 1; i !== 12; ++i) {
        nbOfDays += getNbOfDaysInMonth(i, year);
      }

      ++year;

      for (let i = 0; i !== month.indexOf(today[1]); ++i) {
        nbOfDays += getNbOfDaysInMonth(i, year);
      }

      return nbOfDays;
  }
  
  return 999;
}


function getNbOfDaysInMonth(month, year) {
  if ([0, 2, 4, 6, 7, 9, 11].indexOf(month) !== -1) {
    return 31;
  }
  else if ([3, 5, 8, 10].indexOf(month) !== -1) {
    return 30;
  }
  else if (month === 1) {
    return (isLeap(year) ? 29 : 28);
  }
}


function isLeap(year) {
  year = Number(year);

  if ((year % 4) === 0 && (year % 100) !== 0) {
    return true;
  }
  else if ((year % 400) === 0) {
    return true;
  }

  return false;
}
