function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
      .addItem('Update Data', 'menuItem1')
      .addSubMenu(ui.createMenu('Analisis')
        .addItem('FRh','menuItem2')
        .addItem('Pb','menuItem3')
        .addItem('Q310','menuItem4')
        .addItem('Q350','menuItem5')
        .addItem('Q390','menuItem6')
        .addItem('SPHT 0.6','menuItem7')
        .addItem('SPHT 0.7','menuItem8')
        .addItem('SPHT 0.8','menuItem9')
        .addItem('SPHT 0.9','menuItem10')
        .addItem('b/l = 0.6','menuItem11')
        .addItem('b/l = 0.7','menuItem12')
        .addItem('b/l = 0.8','menuItem13')
        .addItem('b/l = 0.9','menuItem14')
      )
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi()
     agregateData();
}
function menuItem2() {
  SpreadsheetApp.getUi()
     setData(5);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem3() {
  SpreadsheetApp.getUi()
     setData(7);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem4() {
  SpreadsheetApp.getUi()
     setData(8);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem5() {
  SpreadsheetApp.getUi()
     setData(9);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem6() {
  SpreadsheetApp.getUi()
     setData(10);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem7() {
  SpreadsheetApp.getUi()
     setData(11);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem8() {
  SpreadsheetApp.getUi()
     setData(12);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem9() {
  SpreadsheetApp.getUi()
     setData(13);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem10() {
  SpreadsheetApp.getUi()
     setData(14);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem11() {
  SpreadsheetApp.getUi()
     setData(15);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem12() {
  SpreadsheetApp.getUi()
     setData(16);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem13() {
  SpreadsheetApp.getUi()
     setData(17);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function menuItem14() {
  SpreadsheetApp.getUi()
     setData(18);
     sortColumnA();
     generateFrequencyDistribution();
     calculateMed();
     calculateMode();
}
function agregateData() {
  var sheetForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form');
  var sheetConfig = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  var sheetAutomated = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListOfPowder');

  var powderDataCell = sheetForm.getRange('A3:D3').getValues();
  var formDataCell = sheetForm.getRangeList(['H6', 'L6', 'M6']).getRanges();
  var particleDataCell = sheetForm.getRange('F9:F19').getValues();
  Logger.log(powderDataCell[0][1]);
  var dataSheet = powderDataCell[0][1];
  var sheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  var formValues = formDataCell.map(function(range) {
    return range.getValue();
  });

  Logger.log(powderDataCell[0]);

  var combinedData = [].concat(powderDataCell[0], formValues, particleDataCell.flat());

  Logger.log(combinedData);
  var targets;
  if (powderDataCell[0][1] === "**"){
    targets = "A1";
  } else if (powderDataCell[0][1] === "++"){
    targets = "A2";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A3";
  } else if (powderDataCell[0][1] === "++"){
    targets = "A4";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A5";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A6";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A7";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A8";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A9";
  } else if (powderDataCell[0][1] === "**"){
    targets = "A11";
  }else if (powderDataCell[0][1] === "**"){
    targets = "A12";
  }
  
  var configRowCell = sheetConfig.getRange(targets);
  var targetRow = configRowCell.getValue();
  var targetRange = sheetData.getRange(targetRow, 1, 1, combinedData.length);
  targetRange.setValues([combinedData]);

  var configRowCellAutomated = sheetConfig.getRange("A10");
  var targetRowAutomated = configRowCellAutomated.getValue();
  var targetRangeAutomated = sheetAutomated.getRange(targetRowAutomated, 1, 1, powderDataCell[0].length);
  targetRangeAutomated.setValues([powderDataCell[0]]);

  configRowCellAutomated.setValue(targetRowAutomated + 1);
  configRowCell.setValue(targetRow + 1);

  sheetForm.getRange('A3:D3').clearContent();
  sheetForm.getRange('F9:F19').clearContent();
  sheetForm.getRangeList(['F3:F5','G3:G5','J3:J5','K3:K5']).clearContent();

  SpreadsheetApp.getUi().alert("Update complete");
}
function clearClass() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");
  sheet.getRange("H6:N100").clearContent();
}

function generateFrequencyDistribution() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");

  var numClasses = sheet.getRange('E7').getValue();
  var classWidth = sheet.getRange('E11').getValue();
  var minValue = sheet.getRange('E8').getValue();
  var maxValue = sheet.getRange('E9').getValue();
  if (numClasses <= 0 || classWidth <= 0 || minValue >= maxValue) {
    Logger.log(' ERROR!!! Błędne dane wejściowe');
    return;
  }

  var dataRange = sheet.getRange('A:A');
  var ids = dataRange.getValues().flat();
  var dataRange2 = sheet.getRange('E:E');
  var data2 = dataRange2.getValues().flat();

  var currentMin = minValue;

  var startRow = 6;
  var startCol = 8;

  sheet.getRange(startRow - 1, startCol, 1, 8).setValues([[
    "Numer klasy", 
    "Początek klasy <", 
    "Koniec Klasy )", 
    "Liczebność klasy", 
    "Liczebność skumulowana", 
    "Środki przedziałów", 
    "Środki przedziałów razy liczebność",
    "ID"
  ]]);

  var cumulativeFrequency = 0;

  for (var i = 0; i < numClasses; i++) {
    var currentMax = currentMin + classWidth;
    var classMidpoint = (currentMin + currentMax) / 2;
    var classFrequency = data2.filter(function(value, index) {
      return value >= currentMin && value < currentMax && ids[index] !== '';
    }).length;
    cumulativeFrequency += classFrequency;

    sheet.getRange(startRow + i, startCol, 1, 8).setValues([[
      i + 1,
      currentMin,
      currentMax,
      classFrequency,
      cumulativeFrequency,
      classMidpoint,
      classMidpoint * classFrequency,
      ids[i] || "Brak ID"
    ]]);

    currentMin = currentMax;
  }

  Logger.log("Generowanie kozak zakonczone");
}

function calculateMed() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");
  var numA = sheet.getRange("E6").getValue();

  var nrMed = (numA % 2 === 0) ? sheet.getRange("E13").getValue() : sheet.getRange("E14").getValue();

  var rangeL = sheet.getRange("L6:L100");
  var valuesL = rangeL.getValues();
  var tableRange = sheet.getRange("H6:N100");
  var tableData = tableRange.getValues();
  var resultRow = null;
  var previousVal = 0;

  for (var i = 0; i < valuesL.length; i++) {
    if (valuesL[i][0] >= nrMed) {
      resultRow = tableData[i];
      previousVal = (i > 0) ? valuesL[i - 1][0] : 0;
      break;
    }
  }

  if (resultRow) {
    Logger.log("FInd data:" + resultRow);
    sheet.getRange("E15").setValue(tableData[i][3]);

    var medRange = resultRow[2] - resultRow[1];
    var me = resultRow[1] + ((nrMed - previousVal) * (medRange / tableData[i][3]));
    me = Math.round(me * 100) / 100;
    Logger.log("Mediana: " + me);
    sheet.getRange("E3").setValue(me);
  } else {
    Logger.log("Error! Nie znaleziono odpowiedniej klasy dla mediany.");
  }
}

function setData(columnNumber) {
  var sheetAnalisis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");
  var sheetData = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheetAnalisis.getRange("A:A").clear();
  sheetAnalisis.getRange("H6:N100").clear();
  Logger.log(columnNumber);
  var lastRow = sheetData.getLastRow();
  var range = sheetData.getRange(3, columnNumber, lastRow - 2);
  var data = range.getValues();
  sheetAnalisis.getRange(1, 1, data.length, 1).setValues(data);
}

function sortColumnA() {
  var sheetAnalisis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");
  var lastRow = sheetAnalisis.getLastRow();
  if (lastRow > 0) {
    var rangeToSort = sheetAnalisis.getRange(1, 1, lastRow, 1);
    rangeToSort.sort({column: 1, ascending: true});
  } else {
    Logger.log("No data to sort in column A.");
  }
}

function calculateMode() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis");
  var rangeClasses = sheet.getRange("H6:H100");
  var rangeFreq = sheet.getRange("K6:K100");
  var rangeStart = sheet.getRange("I6:I100");
  var rangeEnd = sheet.getRange("J6:J100");

  var classes = rangeClasses.getValues();
  var frequencies = rangeFreq.getValues();
  var startBounds = rangeStart.getValues();
  var endBounds = rangeEnd.getValues();

  var maxFreq = 0;
  var dominantClassIndex = -1;

  for (var i = 0; i < frequencies.length; i++) {
    if (frequencies[i][0] > maxFreq) {
      maxFreq = frequencies[i][0];
      dominantClassIndex = i;
    }
  }

  if (dominantClassIndex === -1) {
    Logger.log("Nie znaleziono dominanty.");
    return;
  }

  var Ld = startBounds[dominantClassIndex][0];
  var Nd = frequencies[dominantClassIndex][0];

  var NdMinus1 = dominantClassIndex > 0 ? frequencies[dominantClassIndex - 1][0] : 0;
  var NdPlus1 = dominantClassIndex < frequencies.length - 1 ? frequencies[dominantClassIndex + 1][0] : 0;

  var h = endBounds[dominantClassIndex][0] - startBounds[dominantClassIndex][0];

  var D = Ld + ((Nd - NdMinus1) / ((Nd - NdMinus1) + (Nd - NdPlus1))) * h;

  D = Math.round(D * 100) / 100;
  sheet.getRange("F3").setValue(D);

  Logger.log("Dominanta: " + D);
}
function sendToAutomated() {
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSpreadsheet.getSheetByName("ListOfPowder");
  var sourceRange = sourceSheet.getRange("A1:C100");
  var data = sourceRange.getValues();
  var targetID = "**";
  var targetSpreadsheet = SpreadsheetApp.openById(targetID);
  var targetSheet = targetSpreadsheet.getSheetByName("PowderList");
  var targetRange = targetSheet.getRange("A1:C100");
  targetRange.setValues(data);
  Logger.log("GitSkopiowane");
}
function translationHashToSerial() {
  var sheetTarget = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Powder Printer Temporary');
  var sheetSerial = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SerialPrinter');

  var range = sheetSerial.getRange(1, 1, sheetSerial.getLastRow(), 2);
  var data = range.getValues();
  
  for (var counter = 810; counter <= 886; counter++) {
    var printer = sheetTarget.getRange("B" + counter).getValue();
    var printerParts = printer.split("#");
    var printerId = printerParts[1];

    if (sheetSerial.getRange("B" + printerId).getValue() == printerId) {
      var serial = sheetSerial.getRange("A" + printerId).getValue();
      sheetTarget.getRange("C" + counter).setValue(serial);
    } else {
      Logger.log("Error! Printer number" + printerId);
    }

    var dateStr = sheetTarget.getRange("A" + counter).getValue();
    sheetTarget.getRange("A" + counter).setValue(convertDateFormat(dateStr));
  }
}

function convertDateFormat(dateStr) {
  var parts = dateStr.split(".");
  
  if (parts.length !== 3) {
    return dateStr;
  }
  
  var day = parts[0];
  var month = parts[1];
  var year = "20" + parts[2];
  
  if (isNaN(day) || isNaN(month) || isNaN(year)) {
    Logger.log("Data zawiera niepoprawne wartości: " + dateStr);
    return dateStr;
  }

  var date = new Date(year, month - 1, day, 14, 0, 0, 0); 

  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd") +
    " " +
    ("0" + date.getHours()).slice(-2) + ":" +
    ("0" + date.getMinutes()).slice(-2) + ":" +
    ("0" + date.getSeconds()).slice(-2) + "." +
    ("000" + date.getMilliseconds()).slice(-3);
  
  Logger.log(formattedDate);
  return formattedDate;
}

function processSheetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = ss.getSheetByName("Powder Printer Temporary");
  var serialSheet = ss.getSheetByName("SerialPrinter");
  
  if (!tempSheet || !serialSheet) {
    Logger.log("Nie znaleziono wymaganych arkuszy.");
    return;
  }
  
  var tempData = tempSheet.getRange("B2:B" + tempSheet.getLastRow()).getValues();
  var serialData = serialSheet.getRange("A2:B" + serialSheet.getLastRow()).getValues();
  
  var serialMap = {};
  serialData.forEach(function(row) {
    var serialNumber = row[1];
    var correspondingValue = row[0];
    if (serialNumber) {
      serialMap[serialNumber] = correspondingValue;
    }
  });
  
  var result = [];
  tempData.forEach(function(row, index) {
    var serialNumber = row[0];
    if (serialNumber && serialNumber.startsWith("#")) {
      serialNumber = serialNumber.substring(1);
    }
    result.push([serialMap[serialNumber] || ""]);
  });
  
  tempSheet.getRange("C2:C" + (result.length + 1)).setValues(result);
}

function countUniquePhrasesPerDate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var datePhraseMap = {};

  for (var i = 0; i < data.length; i++) {
    var date = data[i][0];
    var phrase = data[i][3];

    if (date && phrase) {
      if (!datePhraseMap[date]) {
        datePhraseMap[date] = new Set();
      }
      datePhraseMap[date].add(phrase);
    }
  }

  var output = [["Data", "Unikalne frazy"]];
  for (var date in datePhraseMap) {
    output.push([date, datePhraseMap[date].size]);
  }

  var outputStartRow = 1;
  var outputRange = sheet.getRange(outputStartRow, 6, output.length, 2);
  outputRange.setValues(output);
}
