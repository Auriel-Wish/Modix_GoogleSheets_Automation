/*
 * onChange (built in function type)
 * 
 * Purpose: Runs function whenever change is made to google sheet
 * Parameters: e - the event object (conatins information about the editted cell)
 */
function onChange(e) {
  updateEditSheetWrapper(e);
}


/*
 * updateEditSheetWrapper
 * 
 * Purpose: Call function to update the "CHANGES" sheet whenever an edit is made to the workbook
 * Parameters: e - the event object (conatins information about the editted cell)
 */
function updateEditSheetWrapper(e) {
  changedSheet = SpreadsheetApp.getActiveSheet();
  changedSheetName = changedSheet.getName();
  
  // Record change if it is made on a valid sheet and is either a valid cell edit or inserts/removes a row
  if (validSheetName(changedSheetName)) {
    activeRange = e.source.getActiveRange();
    trackerSheet = SpreadsheetApp.getActive().getSheetByName('CHANGES');

    if (e.changeType == "INSERT_ROW" || e.changeType == "REMOVE_ROW" || (e.changeType == "EDIT" && activeRange.getRow() > 2)) {
      updateEditSheet(e, changedSheet, trackerSheet, activeRange, e.changeType);
    }
  }
}

/*
 * updateEditSheet
 * 
 * Purpose: Log the edit that was made
 * Parameters: the event object (conatins information about the editted cell), the sheet in which the edit took
 *             place, the sheet that keeps track of the edits
 */
function updateEditSheet(e, changedSheet, trackerSheet, activeRange, typeOfChange) {
  
  // Set variables for later
  lastRow = trackerSheet.getLastRow() + 1;
  
  changedSheetName = changedSheet.getName();
  changedRow = activeRange.getRow();
  sectionRow = 2;

  idCol = findSectionCol(changedSheet, "NO.", sectionRow);
  typeCol = findSectionCol(changedSheet, "TYPE", sectionRow);
  boxCol = findSectionCol(changedSheet, "BAG", sectionRow);
  if (boxCol == -1) {
    boxCol = findSectionCol(changedSheet, "BOX", sectionRow);
  }
  partNumberCol = findSectionCol(changedSheet, "PART NUMBER", sectionRow);
  descriptionCol = findSectionCol(changedSheet, "DESCRIPTION", sectionRow);
  qtyCol = findSectionCol(changedSheet, "QTY", sectionRow);

  // If a cell was edited, say what happened in the edit
  if (typeOfChange == "EDIT") {
    value = activeRange.getValue();
    changedSection = changedSheet.getRange(2, activeRange.getColumn(), 1, 1).getValue();
    changedID = changedSheet.getRange(changedRow, idCol, 1, 1).getValue();

    // If the part itself was changed, say that
    if (activeRange.getColumn() == partNumberCol) {
      changeMade = "Updated " + changedSection + " of item #" + changedID + " to: " + value;
    }
    // If something else was changed, say which part it was associated with
    else if (activeRange.getColumn() == descriptionCol || activeRange.getColumn() == qtyCol) {
      partNumber = changedSheet.getRange(changedRow, partNumberCol, 1, 1).getValue();
      changeMade = "Updated " + changedSection + " of item #" + changedID + " (" + partNumber + ") to: " + value;
    }
    else {
      return;
    }
  }
  // If a row was inserted or deleted, say that it was inserted or deleted
  else if (typeOfChange == "INSERT_ROW") {
    changeMade = "Inserted row: " + changedRow;
  }
  else if (typeOfChange == "REMOVE_ROW") {
    changeMade = "Deleted row: " + changedRow;
  }

  // If an edit was made to the same part/row, group the edits together into one big edit
  if (changedSheetName == trackerSheet.getRange('V1').getValue() && changedRow == trackerSheet.getRange('W1').getValue() && typeOfChange == "EDIT") {
      lastRow--;
      if (lastRow < 2) {
        lastRow = 2;
      }
      oldChange = trackerSheet.getRange(lastRow, 5, 1, 1).getValue();
      changeMade = oldChange + "\nAND\n" + changeMade;
  }
  
  // Insert the edits into the tracker
  trackerSheet.getRange(lastRow, 1, 1, 1).setValue(new Date());
  trackerSheet.getRange(lastRow, 2, 1, 1).setValue(changedSheetName);
  if (typeCol != -1) {
    typeData = changedSheet.getRange(changedRow, typeCol, 1, 1).getValue();
    trackerSheet.getRange(lastRow, 3, 1, 1).setValue(typeData);
  }
  if (boxCol != -1) {
    boxData = changedSheet.getRange(changedRow, boxCol, 1, 1).getValue();
    trackerSheet.getRange(lastRow, 4, 1, 1).setValue(boxData);
  }
  trackerSheet.getRange(lastRow, 5, 1, 1).setValue(changeMade);
  trackerSheet.getRange(lastRow, 6, 1, 1).setBackground('red');
  trackerSheet.getRange(lastRow, 6, 1, 1).setValue('Pending');

  // Keep track of which part was updated (in case another edit occurs to the same part)
  trackerSheet.getRange('W1').setValue(changedRow);
  trackerSheet.getRange('V1').setValue(changedSheetName);
}

/*
 * validSheetName
 * 
 * Purpose: Determine if edited sheet is one in which changes should be tracked
 * Parameters: sheetName - name of the edited sheet
 */
function validSheetName(sheetName) {
  invalidSheets = ['Dashboard', 'SKU', 'CHANGES', 'SUMMARY', 'TABLE'];

  lower = invalidSheets.map(element => {
    return String(element).toLowerCase();
  });
  // if the name is not in the array of invalid sheets, return true (valid)
  return !(lower.includes(String(sheetName).toLowerCase()));
}

function clearEdits() {
  sheet = SpreadsheetApp.getActive().getSheetByName('CHANGES');
  startCol = 1;
  startRow = 2;
  numCols = 6;
  lastRow = sheet.getLastRow();
  sheet.getRange(startRow, startCol, lastRow, numCols).clear();
};
