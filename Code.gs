// https://gist.githubusercontent.com/pamelafox/1878143/raw/6c23f71231ce1fa09be2d515f317ffe70e4b19aa/exportjson.js?utm_source=thenewstack&utm_medium=website&utm_content=inline-mention&utm_campaign=platform

// executed upon opening the file 
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "create_jsons", functionName: "scrapeSheets"}
  ];
  ss.addMenu("Export JSON", menuEntries);
}

// used to clear a target drive folder(for testing)
function clearFiles() {
  var targetFolder = DriveApp.getFolderById('1KS-9Zxlbv9dZhFIqzGoCmZe7ocQ0_gk9');
  var tbc = targetFolder.getFiles();
  while (tbc.hasNext()) {
    tbc.next().setTrashed(true);
  }
}

// main function to produce output folder 
function scrapeSheets() {
  // folder to store results 
  var resultFolder = DriveApp.getFolderById('1KS-9Zxlbv9dZhFIqzGoCmZe7ocQ0_gk9');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetsData = {};
  for (var i = 0; i < sheets.length; i++) {
    var currsheet = sheets[i];
    if (isSheetEmpty(currsheet) == true) {
      console.log(currsheet + " is empty")
      continue;
    }

    var compressmeta = getMetaData_(currsheet);

    var rowsData = compressmeta.concat(getRowsData_(currsheet));
    var sheetName = currsheet.getName(); 
    sheetsData[sheetName] = rowsData;
    var json = makeJSON_(sheetsData);
    resultFolder.createFile(currsheet.getSheetName(), json);
    sheetsData = {};
  } 
  displayText_(json);
}

function getMetaData_(sheet) {
  var headersRange = sheet.getRange(findRow(sheet, "Item")+1, 1, 7, 1);
  var headers = headersRange.getValues();
  var dataRange = sheet.getRange(findRow(sheet, "Description")+1, 2, 7, 1);
  var objects = getMetaObjects_(dataRange.getValues(), normalizeHeaders_(headers));
  return objects;
}

function getRowsData_(sheet) {
  var headersRange = sheet.getRange(findRow(sheet, "sample ID"), 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getRange(findRow(sheet, "sample ID")+1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var objects = getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
  return objects;
}

function makeJSON_(object) {
  var jsonString = JSON.stringify(object, null, 4);
  return jsonString;
}

function getMetaObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[i]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader_(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

function normalizeHeader_(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// display json preview in sheet view 
function displayText_(text) {
  var output = HtmlService.createHtmlOutput("<textarea style='width:100%;' rows='20'>" + text + "</textarea>");
  output.setWidth(400)
  output.setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(output, 'Exported JSON');
}

function findRow(currsheet, searchVal) {
  var data = currsheet.getDataRange().getValues();
  var columnCount = currsheet.getDataRange().getLastColumn();

  var i = data.flat().indexOf(searchVal); 
  var columnIndex = i % columnCount
  var rowIndex = ((i - columnIndex) / columnCount);

  //Logger.log({columnIndex, rowIndex }); // zero based row and column indexes of searchVal

  return i >= 0 ? rowIndex + 1 : currsheet.getName();
}

// https://stackoverflow.com/questions/24785987/google-apps-script-find-row-number-based-on-a-cell-value
function tester() {
  //console.log(findRow("sample ID"));
}

// check if a sheet is empty 
function isSheetEmpty(sheet) {
  return sheet.getDataRange().getValues().join("") === "";
}
