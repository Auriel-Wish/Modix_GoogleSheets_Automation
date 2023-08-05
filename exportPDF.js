function exportPDF() {
  // AUTOMAITON/Packing List/Google Sheets
  let googleSheetsFolder = DriveApp.getFolderById('1dz4HCLRI7BDtPcv5JOeRVckgGHDPYx_Z');
  // AUTOMAITON/Packing List/PDFs
  let destinationFolder = DriveApp.getFolderById('1hfOeIS7QmQu-jK31QSRpuk4kEO-8CSQu');
  // AUTOMAITON/Temp
  let tempFolder = DriveApp.getFolderById('18TFKwhzH2m9A6DFN6nM83HWAFD4hdE6u');

  let newFileName = getNewFileName('.pdf');
  // let newFileName = 'tee.pdf'

  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt("Column with unwanted sheets (ex - G):");
  let unwantedSheetsCol = (result.getResponseText()).toUpperCase();
  // let unwantedSheetsCol = 'G';
  
  let googleSheetsFile = copyFile(SpreadsheetApp.getActive(), googleSheetsFolder);
  let newFile = SpreadsheetApp.openById(googleSheetsFile.getId());
  newFile.rename(newFileName);

  manualImportRange(newFile);

  let skuSheet = newFile.getSheetByName("SKU");
  try {
    let skuCol = findSectionCol(skuSheet, "PRICE", 1);
    skuSheet.getRange(1, skuCol, skuSheet.getLastRow(), 1).clear();

    skuCol = findSectionCol(skuSheet, "Sale Price", 1);
    skuSheet.getRange(1, skuCol, skuSheet.getLastRow(), 1).clear();

    skuSheet.getRange(1, skuSheet.getLastColumn(), 1, 1).setValue("Image");
    skuSheet.getRange(skuSheet.getLastRow() + 1, 1, 1, 1).setValue("SIGNAL-END-OF-SKU");
  }
  catch(e) {
    Logger.log(e)
  }

  deleteUnwantedSheets(newFile.getId(), unwantedSheetsCol);
  
  removeFasteners(newFile);

  let allSheets = newFile.getSheets();
  let numSheets = allSheets.length;

  manualImportRangePDF(allSheets);

  // MODIX V4-BOM-SKU
  let subAssemblyID = '1cEEd7h09hxiDGv8hC9trVf3w5lsOA_672eQn_n5vFBU';
  let subWorkbook = SpreadsheetApp.openById(subAssemblyID);

  for (let i = 0; i < numSheets; i++) {
    Logger.log(allSheets[i].getName());

    let partNumberCol = findSectionCol(allSheets[i], "PART NUMBER", 2);
    if (partNumberCol > 0) {
      checkHyperlinks(allSheets[i], partNumberCol, subWorkbook);
    }
  }

  for (let i = 0; i < numSheets; i++) {
    if (!(String(allSheets[i].getName()).toLowerCase().includes("fastener"))) {
      deleteUnwantedColumns(allSheets[i]);

      allSheets[i].getRange(3, 1, allSheets[i].getLastRow(), allSheets[i].getLastColumn()).setBackground('white').setFontColor('black');
    }
  }

  createFile(newFileName, newFile.getId(), destinationFolder);

  skuSheet.getRange(skuSheet.getLastRow(), 1, 1, 1).setValue("");
}

function checkHyperlinks(sheet, partNumberColOrig, subWorkbook) {
  let lastRow = sheet.getLastRow();
  
  let subSpreadsheet;
  let partNumber;
  let colAdded = false;

  for (let i = 3; i <= sheet.getLastRow(); i++) {
    let hyperlink = sheet.getRange(i, partNumberColOrig, 1, 1).getRichTextValue().getLinkUrl();

    sheet.getRange(i, partNumberColOrig + 1, 1, 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(sheet.getRange(i, partNumberColOrig + 1, 1, 1).getValue()).setLinkUrl(null).build());
      sheet.getRange(i, partNumberColOrig + 1, 1, 1).setFontColor('black').setFontLine('none');

    if (hyperlink != null) {
      sheet.getRange(i, partNumberColOrig, 1, 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(sheet.getRange(i, partNumberColOrig, 1, 1).getValue()).setLinkUrl(null).build());
      sheet.getRange(i, partNumberColOrig, 1, 1).setFontColor('black').setFontLine('none');

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
  let numColsToCopy = 3 // Doesn't include picture column
  
  let picCol1 = findSectionCol(sheet1, "PICTURE", 2);
  let picCol2 = findSectionCol(sheet2, "PICTURE", 2);
  let typeCol = findSectionCol(sheet2, "TYPE", 2);

  if (typeCol == -1) {
    Logger.log('Sheet: ' + sheet2.getName() + ' does not have a type column')
  }
  else {
    Logger.log('Sheet: ' + sheet2.getName() + ' has a type column')
  }

  let currType = '-1';
  let currVal;
  let currPics;
  let alph;
  for (let i = 0; i < numRows; i++) {
    if (typeCol != -1) {
      currType = sheet2.getRange(i + startRow2, typeCol, 1, 1).getValue();
    }

    if (!(String(currType).toLowerCase().includes('fastener'))) {
      sheet1.insertRows(startRow1 + i, 1);
      currVal = sheet2.getRange(i + startRow2, col2, 1, numColsToCopy).getValues();
      sheet1.getRange(i + startRow1, col1, 1, numColsToCopy).setValues(currVal);

      alph = String.fromCharCode(col1 + 'A'.charCodeAt(0) - 1);
      currPics = "=VLOOKUP(" + alph + (i + startRow1) + ",SKU!$B$3:$G$1007,6,FALSE())";
      sheet1.getRange(i + startRow1, picCol1, 1, 1).setFormula(currPics);
    }
  }
}

// Delete the price and type columns
function deleteUnwantedColumns(sheet) {
  let currVal = '-1';
  let lastCol = sheet.getLastColumn();
  for (let i = 1; i < lastCol; i++) {
    try {
      currVal = sheet.getRange(2, i, 1, 1).getValue();
      if (String(currVal).toLowerCase().includes('price') || String(currVal).toLowerCase().includes('type')) {
        Logger.log('Deleting: ' + currVal)
        sheet.deleteColumn(i);
        i--;
      }
    }
    catch (error) {
      Logger.log(error);
    }
  }
}

function manualImportRangePDF(allSheets) {
  for (let i = 0; i < allSheets.length; i++) {
    if (String(allSheets[i].getName()).toLowerCase().includes('fastener')) {
      let importedFasteners = allSheets[i];
      let realFasteners = SpreadsheetApp.getActiveSpreadsheet();
      realFasteners = realFasteners.getSheetByName(allSheets[i].getName());

      let copyVals = realFasteners.getRange(1, 1, realFasteners.getLastRow(), realFasteners.getLastColumn()).getValues();
      importedFasteners.getRange(1, 1, realFasteners.getLastRow(), realFasteners.getLastColumn()).setValues(copyVals);

      importedFasteners.getRange('A1').setValue('');
      Logger.log(importedFasteners.getRange('A1').getValue());
      
      break;
    }
  }
}

function getSheetById(file, id) {
  return file.getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}
