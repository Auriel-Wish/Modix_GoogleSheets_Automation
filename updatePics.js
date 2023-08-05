function updatePics() {
  let activeSS = SpreadsheetApp.getActiveSpreadsheet();
  let sku = activeSS.getSheetByName('SKU');
  let lastRow = sku.getLastRow();
  let itemCol = findSectionCol(sku, 'SKU', 1);
  let picCol = findSectionCol(sku, 'Image', 1);
  let picName;
  let filesDict = initDict();

  let dashboard = activeSS.getSheetByName('dashboard');
  let percentCell = 'B7';
  let progressCell = 'C7';
  makeProgressBar(dashboard, percentCell, progressCell);

  sku.getRange(1, picCol, lastRow, 1).setHorizontalAlignment("center").setFontColor('black');

  for (let i = 3; i <= lastRow; i++) {
    updateProgressBar(dashboard, i / lastRow, percentCell);
    picName = sku.getRange(i, itemCol, 1, 1).getValue();

    if (picName != "") {
      picName = (String(picName)).toLowerCase();
      sku.getRange(i, picCol, 1, 1).setValue(setPic(picName, filesDict));
    }
  }
}

function initDict() {
  let folderName = 'PICTURES';
  let folder = DriveApp.getFoldersByName(folderName).next();
  let files = folder.getFiles();

  let nextFile;
  let nextFileName;
  let fileDict = {};
  while (files.hasNext()) {
    nextFile = files.next();
    nextFileName = nextFile.getName();
    nextFileName = String(nextFileName).toLowerCase();
    nextFileName = nextFileName.slice(0, nextFileName.lastIndexOf("."));
    fileDict[nextFileName] = nextFile;
  }

  return fileDict;
}

function setPic(picName, fileDict) {
  let imageLink = 'NO PICTURE AVAILABLE';
  if (picName in fileDict) {
    let file = fileDict[picName];

    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
    let fileID = file.getId();
    imageLink = '=IMAGE("http://drive.google.com/uc?export=view&id=' + fileID + '")';
  }
  
  return imageLink;
}
