/*
 * findSectionCol
 * 
 * Purpose: In the changed sheet, find the column associated with the given section
 * Parameters: sheet - the changed sheet to look in, sectionName - the name of the desired section, row - the row
 *             conatining the section labels
 */
function findSectionCol(sheet, sectionName, row) {
  currSection = "-1";
  currCol = 0;
  // keep scanning until the section is found or until there are no more sections names to scan
  while (!(String(currSection).toLowerCase().includes(String(sectionName).toLowerCase())) && currCol <= sheet.getLastColumn()) {
    currCol++;
    currSection = sheet.getRange(row, currCol, 1, 1).getValue();
  }

  if (currSection == "") {
    return -1;
  }
  return currCol;
}

function fillPicFormula() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let sheet;
  let picCol;
  let itemCol;
  let lastRow;
  let val;
  let alph;
  let unwantedSheets = ['SKU', 'CHANGES', 'SUMMARY', 'OLD CHANGES', 'Dashboard', 'TABLE'];
  let skuLastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SKU").getLastRow();

  for (let i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    Logger.log(i);
    if (!unwantedSheets.includes(sheet.getName())) {
      picCol = findSectionCol(sheet, 'PICTURE', 2);
      itemCol = findSectionCol(sheet, 'PART NUMBER', 2);
      lastRow = sheet.getLastRow();
      
      if (picCol != -1 && itemCol != -1) {
        for (let j = 3; j <= lastRow; j++) {
          alph = String.fromCharCode(itemCol + 'A'.charCodeAt(0) - 1);
          alph += j.toString();
          val = "=VLOOKUP(" + alph + ",SKU!$B$3:$G$" + skuLastRow + ",6,FALSE())";
          if (sheet.getRange(j, picCol - 1, 1, 1).getValue() != "") {
            sheet.getRange(j, picCol, 1, 1).setValue(val);
          }
        }
      }
    }
  }
}

function getNewFileName(fileType) {
  let cellWithName = getCellWithFileName();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard').getRange(cellWithName).getValue() + fileType;
}

function createFile(fileName, fileID, destinationFolder) {
  let blob = DriveApp.getFileById(fileID).getBlob(); 
  blob.setName(fileName);
  destinationFolder.createFile(blob);
  SpreadsheetApp.getUi().alert(fileName + " created");
}

function copyFile(activeSheet, destinationFolder) {
  //Input: Active google sheet, destination folder for the copied file
  //output: file stream of the copied file
  let id = activeSheet.getId();
  const file_name = activeSheet.getName()
  let new_file = DriveApp.getFileById(id).makeCopy(file_name, destinationFolder);//Create the copy of the file
  return new_file;
}

function deleteUnwantedSheets(fileId, col) {
  let wb = SpreadsheetApp.openById(fileId);
  let dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  let sheetNames = dashboard.getRange(col + "1:" + col + dashboard.getLastRow()).getValues();
  let allSheets = wb.getSheets();
  
  for (let i = 0; i < allSheets.length; i++) {
    for (let j = 0; j < sheetNames.length; j++) {
      if (allSheets[i].getName() == sheetNames[j][0]) {
        wb.deleteSheet(allSheets[i]);
        break;
      }
    }
  }
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

function makeProgressBar(sheet, percentCell, progressCell) {
  sheet.getRange(percentCell).setNumberFormat("##.#%");
  sheet.getRange(progressCell).setValue('=SPARKLINE(' + percentCell + ',{"charttype","bar";"max",1;"min",0;"color1","green"})');
}

function updateProgressBar(sheet, percentage, percentCell) {
  sheet.getRange(percentCell).setValue(percentage);
}

function getCellWithFileName() {
  var ui = SpreadsheetApp.getUi();
  let result = ui.prompt("Cell with file name (ex - E3):");
  return (result.getResponseText()).toUpperCase();
}

// Include the pictures column in the SKU lookup
function changeRange() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    Logger.log(i);
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    let unwantedSheets = ['SKU', 'CHANGES', 'SUMMARY', 'OLD CHANGES', 'Dashboard', 'TABLE'];

    if (!unwantedSheets.includes(sheet.getName())) {
      for (let row = 1; row <= lastRow; row++) {
        for (let col = 1; col <= lastCol; col++) {
          currVal = String(sheet.getRange(row, col, 1, 1).getFormula());
          if (currVal.includes("SKU!B:E")) {
            newVal = currVal.replace("SKU!B:E", "SKU!B:F");
            sheet.getRange(row, col, 1, 1).setValue(newVal);
          }
        }
      }
    }
  }
}


// function help() {
//   let sheetsLocal = SpreadsheetApp.getActiveSpreadsheet().getSheets();
//   let sheetsOrig = SpreadsheetApp.openById("1004nJoM7DAubjdMN8_b_yvV3ReIW_Wl6GBQX9eFlF8o").getSheets();
//   // Logger.log('LOCAL\n' + sheetsLocal + '\n\nORIGINAL\n' + sheetsOrig);
//   for (let i = 4; i < sheetsLocal.length; i++) {
//     // Logger.log(sheetsLocal[i].getName() == sheetsOrig[i].getName());
//     sheetsLocal[i].getRange('B7').setValue(sheetsOrig[i].getRange('B7').getValue());
//   }
// }
