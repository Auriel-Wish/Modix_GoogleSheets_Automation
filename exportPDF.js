function exportPDF() {
  // AUTOMAITON/TEMP
  let tempFolder = DriveApp.getFolderById('18TFKwhzH2m9A6DFN6nM83HWAFD4hdE6u');
  // AUTOMAITON/Packing List
  let destinationFolder = DriveApp.getFolderById('1liz9opXDr3h87oRjzCYW0cxsfSb0obmC');

  // progress bar
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard = ss.getSheetByName('Dashboard');
  let percentCell = 'B11';
  let progressCell = 'C11';
  SKU_Automation.makeProgressBar(dashboard, percentCell, progressCell);
  SKU_Automation.updateProgressBar(dashboard, 0, percentCell);

  ss = SpreadsheetApp.getActive();
  let fstream = SKU_Automation.copyFile(ss, tempFolder);
  let newFile = SpreadsheetApp.openById(fstream.getId());
  let allSheets = newFile.getSheets();
  let numSheets = allSheets.length;
  let unwantedSheets = ['SKU', 'CHANGES', 'SUMMARY', 'OLD CHANGES', 'Dashboard'];
  // MODIX V4-BOM-SKU
  let subAssemblyID = '1cEEd7h09hxiDGv8hC9trVf3w5lsOA_672eQn_n5vFBU';
  let subWorkbook = SpreadsheetApp.openById(subAssemblyID);

  SKU_Automation.updateProgressBar(dashboard, 0.2, percentCell);
  let progressAmt;
  for (let i = 0; i < numSheets; i++) {
    Logger.log(allSheets[i].getName());

    progressAmt = 0.2 + i / numSheets;
    if (progressAmt < 0.2) {
      progressAmt = 0.2;
    }
    else if (progressAmt > 0.8) {
      progressAmt = 0.8;
    }
    SKU_Automation.updateProgressBar(dashboard, progressAmt, percentCell);

    if (unwantedSheets.includes(allSheets[i].getName())) {
      // delete the unwanted sheets
      newFile.deleteSheet(allSheets[i]); 
    }
    else {
      deleteUnwantedColumns(allSheets[i]);
      let partNumberCol = findSectionCol(allSheets[i], "PART NUMBER", 2);
      if (partNumberCol > 0) {
        checkHyperlinks(allSheets[i], partNumberCol, subWorkbook);
      }
    }
  }

  Logger.log('Finished Editing Sheet');

  let newFileName = SpreadsheetApp.getActive().getName();
  let fileType = 'pdf';
  createFile(newFileName, newFile.getId(), destinationFolder, fileType);

  SKU_Automation.updateProgressBar(dashboard, 1, percentCell);

  fstream.setTrashed(true);
}

function createFile(fileName, fileID, destinationFolder, fileType) {
  let blob = DriveApp.getFileById(fileID).getBlob(); 
  blob.setName(fileName + "." + fileType);
  destinationFolder.createFile(blob);
  SpreadsheetApp.getUi().alert(fileName + " created");
}

function checkHyperlinks(sheet, partNumberColOrig, subWorkbook) {
  let lastRow = sheet.getLastRow();
  
  let subSpreadsheet;
  let partNumber;
  let colAdded = false;

  for (let i = 3; i <= lastRow; i++) {
    let hyperlink = sheet.getRange(i, partNumberColOrig, 1, 1).getRichTextValue().getLinkUrl();
    if (hyperlink != null) {
      if (!colAdded) {
        colAdded = true;
        sheet.insertColumnAfter(partNumberColOrig);
        sheet.getRange(2, partNumberColOrig + 1, 1, 1).setValue('SUBASSEMBLY PART NUMBER');
      }

      partNumber = sheet.getRange(i, partNumberColOrig, 1, 1).getValue();
      subSpreadsheet = subWorkbook.getSheetByName(partNumber);

      if (subSpreadsheet != null) {
        copyCols(sheet, subSpreadsheet, partNumberColOrig + 1, findSectionCol(subSpreadsheet, "PART NUMBER", 2), i + 1, 3);
      }
    }
  }

  if (colAdded) {
    sheet.getRange(3, partNumberColOrig + 1, lastRow - 2, 1).setFontColor('black').setFontLine('none');
  }
}

// copy from sheet 2 to sheet 1
function copyCols(sheet1, sheet2, col1, col2, startRow1, startRow2) {
  let numRows = sheet2.getLastRow() - startRow2 + 1;
  sheet1.insertRows(startRow1, numRows);
  let numColsToCopy = 3 // CHANGE TO 4 ONCE PICTURES ARE WORKING
  let vals = sheet2.getRange(startRow2, col2, numRows, numColsToCopy).getValues();
  let pasteRange = sheet1.getRange(startRow1, col1, numRows, numColsToCopy);
  pasteRange.setValues(vals);
}

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
  while (!(currSection.toLowerCase().includes(sectionName.toLowerCase())) && currSection != "") {
    currCol++;
    currSection = sheet.getRange(row, currCol, 1, 1).getValue();
  }

  if (currSection == "") {
    return -1;
  }
  return currCol;
}

// Delete the price and type columns
function deleteUnwantedColumns(sheet) {
  let currVal = '-1';
  for (let i = 1; currVal != ''; i++) {
    currVal = sheet.getRange(2, i, 1, 1).getValue();
    if (currVal.toLowerCase().includes('price') || currVal.toLowerCase().includes('type')) {
      sheet.deleteColumn(i);
      i--;
    }
  }
}