function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
      .addItem('Aktualizuj_Numery_Wewnętrzne', 'menuItem1')
      .addItem('Trace_Powder_History', 'menuItem2')
      .addSeparator()
      .addSubMenu(ui.createMenu('Hide/Unhide')
        
        .addItem('Hide_Null', 'menuItem3')
        .addItem('UnHide_Null', 'menuItem4'))
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi()
     .alert('Nadawanie numerów wew');
     InsideNumber();
}

function menuItem2() {
  SpreadsheetApp.getUi()
    clearTrace();
    Trace();
}
function menuItem3() {
  SpreadsheetApp.getUi();
  HideRowNull();
}
function menuItem4(){
  SpreadsheetApp.getUi();
  UnHideRowNull();
}
function menuItem5(){
  SpredsheetApp.getUi();
  MoveNull();
}

function InsideNumber() {
  var arkusz = SpreadsheetApp.getActiveSpreadsheet();
  var arcActive = arkusz.getActiveSheet();
  var colC = arcActive.getRange("C2:C" + arcActive.getLastRow()).getValues();
  var colB = arcActive.getRange("B2:B" + arcActive.getLastRow()).getValues();
  
  for (var i = 0; i < colC.length; i++) {
    if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 1000, 2000);
      colB[i][0] = numberr;
    } else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 2000, 2500);
      colB[i][0] = numberr;
    }else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 3000, 4000);
      colB[i][0] = numberr;
    }else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 4000, 5000);
      colB[i][0] = numberr;
    }else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 5000, 6000);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 2500, 3000);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 7000, 7500);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 7500, 8000);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 8000, 8500);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 9000, 9500);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 9500, 10000);
      colB[i][0] = numberr;
    }
    else if (colC[i][0] === "**" && colB[i][0] == "") {
      var numberr = findNumber(arcActive, colB, 10000, 10500);
      colB[i][0] = numberr;
    }

  }
  
  var dataUpdate = [];
  for (var j = 0; j < colB.length; j++) {
    dataUpdate.push([colB[j][0] != "" ? colB[j][0] : ""]);
  }
  arcActive.getRange("B2:B" + arcActive.getLastRow()).setValues(dataUpdate);
}

function findNumber(arkusz, colB, start, end) {
  var numberr = start;

  while (true) {
    var znaleziono = false;
    
    for (var i = 0; i < colB.length; i++) {
      if (colB[i][0] == numberr) {
        znaleziono = true;
        break;
      }
    }

    if (!znaleziono && numberr <= end) {
      return numberr;
    }

    numberr++;
  }
}


function Trace(){

  var fin = 0;
  fin = SpreadsheetApp.getUi().prompt('Find Powders By Wew:').getResponseText();
  var ar = SpreadsheetApp.getActiveSpreadsheet();
  var arDB = ar.getSheetByName("Powders");
  var arTrac = ar.getSheetByName("Powder_History_Tracer");
  
  var datDB = arDB.getDataRange().getValues();
  var datTrac = arTrac.getDataRange().getValues();
  
  console.log("fin:" +fin);

  var excelenceData = 0;
  var g = 2;

  for (var i=0; i<datDB.length; i++){
    if (datDB[i][1] == fin){
      excelenceData = datDB[i];
      arTrac.getRange(g, 1, 1, excelenceData.length).setValues([excelenceData]);
      fin = datDB[i][7];
      i = 0;
      g++;
    }
  }
}

function clearTrace() {
 var range = SpreadsheetApp
               .getActive()
               .getSheetByName("Powder_History_Tracer")
               .getRange(2,1,12,12);
 range.setValue("");
}


function HideRowNull(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Powders");
  var data = sheet.getDataRange().getValues();
  for(var i = 1; i < data.length; i++) {
  if(data[i][12] === true) {
    console.log(i);
    sheet.hideRows(i + 1);
  }
}
}
function UnHideRowNull(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Powders");
  var data = sheet.getDataRange().getValues();
  for(var i = 1; i < data.length; i++) {
  if(data[i][12] === true) {
    console.log(i);
    sheet.showRows(i + 1);
      }
   }
  }

  
