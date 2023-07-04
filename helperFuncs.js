function fillPicFormula() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let sheet;
  let picCol;
  let itemCol;
  let lastRow;
  let val;
  let alph;
  let unwantedSheets = ['SKU', 'CHANGES', 'SUMMARY', 'OLD CHANGES', 'Dashboard', 'TABLE'];

  for (let i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    Logger.log(i);
    if (!unwantedSheets.includes(sheet.getName())) {
      picCol = findSectionCol(sheet, 'PICTURE', 2);
      itemCol = findSectionCol(sheet, 'PART NUMBER', 2);
      lastRow = sheet.getLastRow();
      
      for (let j = 3; j <= lastRow; j++) {
        alph = String.fromCharCode(itemCol + 'A'.charCodeAt(0) - 1);
        alph += j.toString();
        val = "=VLOOKUP(" + alph + ",SKU!$B$3:$G$1058,6,FALSE())";
        if (sheet.getRange(j, picCol - 1, 1, 1).getValue() != "") {
          sheet.getRange(j, picCol, 1, 1).setValue(val);
        }
      }
    }
  }
}

function copyFile(activeSheet, destinationFolder) {
  //Input: Active google sheet, destination folder for the copied file
  //output: file stream of the copied file
  let id = activeSheet.getId();
  const file_name = activeSheet.getName()
  let new_file = DriveApp.getFileById(id).makeCopy(file_name, destinationFolder);//Create the copy of the file
  return new_file;
}

function getFileAsBlob(exportUrl) {
  let response = UrlFetchApp.fetch(exportUrl, {
     muteHttpExceptions: true,
     headers: {
       Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
     },
   });
  return response.getBlob();
}

/*
 * findSectionCol
 * 
 * Purpose: In the changed sheet, find the column associated with the given section
 * Parameters: sheet - the changed sheet to look in, sectionName - the name of the desired section, row - the row
 *             conatining the section labels
 */
function findSectionCol(sheet, sectionName, row) {
  let currSection = "-1";
  let currCol = 0;
  let blank = 0;
  // keep scanning until the section is found or until there are no more sections names to scan
  while (!(currSection.toLowerCase().includes(sectionName.toLowerCase())) && blank < 4) {
    currCol++;
    currSection = (sheet.getRange(row, currCol, 1, 1).getValue()).toString();

    if (currSection == "") {
      blank++;
    }
    else {
      blank = 0;
    }
  }

  if (currSection == "") {
    return -1;
  }
  return currCol;
}

function makeProgressBar(sheet, percentCell, progressCell) {
  sheet.getRange(percentCell).setNumberFormat("##.#%");
  sheet.getRange(progressCell).setValue('=SPARKLINE(' + percentCell + ',{"charttype","bar";"max",1;"min",0;"color1","green"})');
}

function updateProgressBar(sheet, percentage, percentCell) {
  Logger.log(percentage);
  sheet.getRange(percentCell).setValue(percentage);
}