// Call export SKUNF either way but pass in false if you really only want to call exportBOMHelper.
// This allows for the removal of unwanted sheets
function exportSKUNF() {
  exportSKUNFHelper(true);
}

function exportBOM() {
  exportSKUNFHelper(false);
}

/*
 * exportSKUNFHelper
 * 
 * Purpose: Prepare the SKU file to be exported - remove all the items of type fastener
 */
function exportSKUNFHelper(deleteFasteners) {  
  let fileName = getNewFileName('.xlsx');

  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt("Column with unwanted sheets (ex - G):");
  let unwantedSheetsCol = (result.getResponseText()).toUpperCase();
  
  // AUTOMAITON/TEMP
  let tempFolder = DriveApp.getFolderById('18TFKwhzH2m9A6DFN6nM83HWAFD4hdE6u');

  let tempFile = copyFile(SpreadsheetApp.getActive(), tempFolder);
  let newFile = SpreadsheetApp.openById(tempFile.getId());

  manualImportRange(newFile);
  deleteUnwantedSheets(newFile.getId(), unwantedSheetsCol);

  if (deleteFasteners) {
    removeFasteners(newFile);
  }

  exportBOMHelper(newFile.getId(), fileName)

  // move the temp file to trash
  tempFile.setTrashed(true);
}

function removeFasteners(newFile) {
  var sheets = newFile.getSheets();

  // Loop through all the sheets in the workbook
  for(var i = 0; i < sheets.length; i++) {
    Logger.log('Removing fasteners: ' + sheets[i].getName());

    // Find the "Type" column
    let column = findSectionCol(sheets[i], "Type", 2);
    
    // If the "Type" column exists, remove any rows where the type is Fastener
    if(column != -1) {
      // Get all the values in the "Type" column
      var range = sheets[i].getRange(3, column, sheets[i].getLastRow(), 1);
      var values = range.getValues();

      var row = 3
      var count = 0;
      // Remove all rows that are of type Fastener
      for(var j = 0; j < values.length; j++) {
        if(values[j][0] == "Fasteners") {
          sheets[i].deleteRow(row + j - count);
          
          if (count == 0) {
            // Rename the sheet with -NF at the end (No Fasteners)
            var newName = sheets[i].getName() + "-NF";
            sheets[i].setName(newName);
          }

          count++;
        }
      }
    }
  }
}

function manualImportRange(newFile) {
  let importedSKU = newFile.getSheetByName('SKU');

  let realSKU = SpreadsheetApp.openById('1cEEd7h09hxiDGv8hC9trVf3w5lsOA_672eQn_n5vFBU');
  realSKU = realSKU.getSheetByName('SKU');

  let copyVals = realSKU.getRange(1, 1, realSKU.getLastRow(), realSKU.getLastColumn() - 1).getValues();
  importedSKU.getRange(1, 1, realSKU.getLastRow(), realSKU.getLastColumn() - 1).setValues(copyVals);

  copyVals = realSKU.getRange(1, realSKU.getLastColumn(), realSKU.getLastRow(), 1).getFormulas();
  importedSKU.getRange(1, realSKU.getLastColumn(), realSKU.getLastRow(), 1).setFormulas(copyVals);

  importedSKU.getRange('A1').setValue('');
  Logger.log(importedSKU.getRange('A1').getValue());
}

/*
 * exportBOMHelper
 * 
 * Purpose: Export the BOM file to the /drive/AUTOMATION/BOM folder
 * Parameters: id - the file ID (for google drive) of the no-fasteners SKU file
 */
function exportBOMHelper(id, fileName) {
  // Set where the BOM file will be saved - /drive/AUTOMATION/BOM
  let destinationFolder = DriveApp.getFolderById('11l_03Ug2OA3LIT1x2mwOQmwH7ZrlIBL7');
  let exportURL = "https://docs.google.com/spreadsheets/d/" + id.toString() + "/export?format=xlsx";
  
  // Create the file
  let blob = getFileAsBlob(exportURL);
  blob.setName(fileName);

  // Place file in folder
  destinationFolder.createFile(blob);

  SpreadsheetApp.getUi().alert("File Created: " + fileName);
}
