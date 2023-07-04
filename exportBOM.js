/*
 * exportSKUNF
 * 
 * Purpose: Prepare the SKU file to be exported - remove all the items of type fastener
 */
function exportSKUNF()
{  
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let dashboard = ss.getSheetByName('Dashboard');
  let percentCell = 'B3';
  let progressCell = 'C3';
  makeProgressBar(dashboard, percentCell, progressCell);
  updateProgressBar(dashboard, 0, percentCell);
  
  // Access google drive folder - /drive/AUTOMATION/TEMP
  const destination_folder = DriveApp.getFolderById('18TFKwhzH2m9A6DFN6nM83HWAFD4hdE6u');

  // Open MODIX V4-BOM-SKU and copy that sheets file into the folder
  var file = copyFile(ss, destination_folder);
  updateProgressBar(dashboard, 0.1, percentCell);

  // Open this copied file and access its sheets
  let new_file = SpreadsheetApp.openById(file.getId());
  var sheets = new_file.getSheets();

  updateProgressBar(dashboard, 0.2, percentCell);
  // Loop through all the sheets in the workbook
  let progressAmt = 0.2
  for(var i = 0; i < sheets.length; i++) {
    // Find the "Type" column
    let column = findSectionCol(sheets[i], "Type", 2);
    
    progressAmt = 0.2 + i / sheets.length;
    if (progressAmt < 0.2) {
      progressAmt = 0.2;
    }
    else if (progressAmt > 0.8) {
      progressAmt = 0.8;
    }
    updateProgressBar(dashboard, progressAmt, percentCell);
    
    // If the "Type" column exists, remove any rows where the type is Fastener
    if(column != -1)
    {
      // Unecessary
      sheets[0].getRange("B2").setValue(sheets[i].getName())

      // Get all the values in the "Type" column
      var range = sheets[i].getRange(3,column+1,sheets[i].getLastRow())
      var values = range.getValues();

      // Rename the sheet with -NF at the end (No Fasteners)
      var newName = sheets[i].getName() + "-NF"
      sheets[i].setName(newName);

      var row = 3
      var count = 0;
      // Remove all rows that are of type Fastener
      for(var j=0; j<values.length; j++)
      {
        if(values[j] == "Fasteners")
        {
          sheets[i].deleteRow(row + j - count)
          count++;

          // Unnecessary
          sheets[0].getRange("B3").setValue("item " + count + " deleted")
        }
      }
    }
  }

  exportBOM(new_file.getId())

  updateProgressBar(dashboard, 1, percentCell);

  // move the temp file to trash
  file.setTrashed(true);
}

/*
 * exportBOM
 * 
 * Purpose: Export the BOM file to the /drive/AUTOMATION/EXCEL folder
 * Parameters: id - the file ID (for google drive) of the no-fasteners SKU file
 */
function exportBOM(id) {
  // Set where the BOM file will be saved - /drive/AUTOMATION/EXCEL
  const destination_folder = DriveApp.getFolderById('11l_03Ug2OA3LIT1x2mwOQmwH7ZrlIBL7');
  var printer_name = "SKU";
  let exportURL = "https://docs.google.com/spreadsheets/d/" + id.toString() + "/export?format=xlsx";
  
  // Ask user to supply version number for BOM file
  var ui = SpreadsheetApp.getUi();
  let result = ui.prompt("Version Number:")

  // If they don't cancel the export, go through with the export
  if (result.getSelectedButton() == ui.Button.OK)
  {
    let version = result.getResponseText(); 
    let fileName = printer_name + "-V4-BOM_V" + version;

    // Create the file
    let blob = getFileAsBlob(exportURL);
    blob.setName(fileName + ".xlsx");

    // Place file in folder
    destination_folder.createFile(blob);
    ui.alert("File Created:" + fileName);
  }
  else
  {
    ui.alert("Operation Canceled");
  }
}