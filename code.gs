 function onOpen(e) {
   SpreadsheetApp.getUi()
   .createMenu('Coniston')
   .addItem('Prepare new availability form', 'Clearavailability_')
   .addItem('Produce and email report', 'Reportmaker_')
   .addSeparator()
   .addItem('Update form', 'setupAvailability_')
   .addToUi();
 };

function setupAvailability_ () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('data');
  var lastRowName = sheet.getRange("A1:A").getValues().filter(String).length;
  var lastRowTime = sheet.getRange("C1:C").getValues().filter(String).length;
  var nameRange = sheet.getRange(1, 1, lastRowName).getValues();
  var dayRange = sheet.getRange(1, 2, 7).getValues();
  var timeRange = sheet.getRange(1, 3, lastRowTime).getValues();
//  Logger.log(dayRange);
//  Logger.log(timeRange);
  Wipeform();
  Setupform(nameRange,dayRange,timeRange);
  
  Deletesheet();
  Renamesheet();
};

function Clearavailability_ () {
  var form = FormApp.openById("1MVvkX4T6P77qQLene4fgKQaRsAwzMBrUi7xSem71oFM");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('rawdata');
  var lastRow =sheet.getLastRow()-1;
  
  form.deleteAllResponses();
  if(lastRow <1) {
  sheet.deleteRows(2,lastRow);
  }

  Clearreport(ss);
};

function Reportmaker_ () {
  //Declaring stuff
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName('rawdata');
  var dataSheet = ss.getSheetByName('data');
  var reportSheet = ss.getSheetByName('report');
  var lastRowName = dataSheet.getRange("A1:A").getValues().filter(String).length;
  var lastRowTime = dataSheet.getRange("C1:C").getValues().filter(String).length;
  var nameRange = dataSheet.getRange(1, 1, lastRowName).getValues();
  var timeRange = dataSheet.getRange(1, 3, lastRowTime).getValues();
  
  makeReport (ss,rawSheet,dataSheet,reportSheet,nameRange,lastRowName,timeRange);
};


function makeReport (ss,rawSheet,dataSheet,reportSheet,nameRange,lastRowName,timeRange) {
  //Fill in latest names from list
//  var range = reportSheet.getRange(4,1,lastRowName); 
//  range.setValues(nameRange);
  
  var newValue = rawSheet.getRange('E4').getValue().split(', ');
  var avail = [];
  var data = [];
  var k= 0;
  var row = 2;
  var repRow = 4;
  var repCol =1;
  // goes though each day of the week to get the ticks.
  for (var j=2; j<rawSheet.getLastRow();j++) {
    avail = [];
    avail.push(rawSheet.getRange(j,2).getValue());
    range = reportSheet.getRange(repRow,repCol,1,42);
    for (var col=3;col<10;col++) {
      k= 0;
      var newValue = rawSheet.getRange(row,col).getValue().split(', ');
      Logger.log("row is "+row + "col is "+col);
      
      //serches for Matches in the rawdata cell and puts a tick if matches and a n if not.  might change that to blank later so i don't have to double handle data
      for (var i=0;i < timeRange.length;i++) {  
      if (timeRange[i] == newValue[k]) {
        //        Logger.log("Time "+timeRange[i] +" Matches value "+ newValue[k])
        avail.push("✔");    
        k++;
        Logger.log(avail);
      }
      else {
        avail.push(" ");
        //        Logger.log("Time "+timeRange[i] +" Did not match "+ newValue[k]);
        //        Logger.log(avail);
      }
      
    };       
    };
    Logger.log(avail);
    range.setValues([avail]);
    repRow++
    };
   //  avail.push("N");
  //  avail.push("✔");
  //  range = reportSheet.getRange(4,2,1,newValue.length);
  //  range.setValues(outerArray);

  
  
  ////grab the data from rawdata
  //  var lastRowRawData = rawSheet.getLastRow()+1; 
  //  var data= [];
  //  for (var i = 2; i <lastRowRawData; i++) {
  //    var startCol = 2;
  //    for (j = startCol; j < 7; j++) {
  //      data.push(rawSheet.getRange(i,startCol,1,j).getValues());
  //    }
//    startCol = 2;
  //    //data.push(rawSheet.getRange(i,2,1,7).getValues());
  //    //prints out the data    
  //  }
  //   var rangeOfNames = rawSheet.getRange(2, 2, 3).getValues();
  //   var data = [];
  //   var values = rawSheet.getRange(2, 3, 3, 6).getValues();
  //   var idx = 0;
  //  
  //   // Get values from raw data sheet
  //   for (var row in values) {
  //     var newData = []; // should be 42 long.
  //     var k = 0;
  //     for (var col in values[row]) {
  //       var splitted = values[row][col].split(', ');
  //       for(var j in splitted) {
  //         newData[k++] = splitted[j]; 
  //         Logger.log(k + " " + splitted[j] + "\n");
  //       }       
  //     }     
  //     data[idx++] = newData;
  //   }
  //  
  //  //Logger.log(data);
  //
  //   var newrange = reportSheet.getRange(4,2,3,42); 
  //   //newrange.setValues(data);
};


function Wipeform () {
  var form = FormApp.openById("1MVvkX4T6P77qQLene4fgKQaRsAwzMBrUi7xSem71oFM");
  form.deleteAllResponses();
  // Deletes Questions
  while (form.getItems().length >0) {
  form.deleteItem(0);
  }
};

function Setupform(nameRange,dayRange,timeRange) {  //Function to create the Form to fill in 
  var form = FormApp.openById("1MVvkX4T6P77qQLene4fgKQaRsAwzMBrUi7xSem71oFM")
  .setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
    //setup a dropdown list with names from Spreadsheet
  var item = form.addListItem().setRequired(true);
  
  var thisValue = "";
  var arrayOfItems = [];
  var newItem = "";
  
  for (var i=0;i<nameRange.length;i++) {
    thisValue = nameRange[i][0];
    newItem = item.createChoice(thisValue);
    arrayOfItems.push(newItem);
  }

  item.setTitle('Select your name') //creates the choose from list question
     .setChoices(arrayOfItems)
   
  //Setup the Days of the week
  for (var x in dayRange) {
  //declare variables adds new question
  var item2 = form.addCheckboxItem().showOtherOption(false);
  var thisValue2 = "";
  var arrayOfTimes = [];
  var newItem2 = "";
  //  This Section will do the Times as this won't change we we do the loop one and just repost the results each day. 
  for (var i=0;i<timeRange.length;i++) {
    thisValue2 = timeRange[i][0];
    newItem2 = item2.createChoice(thisValue2);
    arrayOfTimes.push(newItem2);
   }
  
  //creates title and addstime slots
    item2.setTitle(dayRange[x])
     .setChoices(arrayOfTimes);

    //clear varibales for next loop though
  item2 = "";
  thisValue2 = "";
  arrayOfTimes = [];
  newItem2 = "";
  }
};  

function Deletesheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.setActiveSheet(ss.getSheetByName('rawdata'));
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
};

function Renamesheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet('rawdata');
};

function Clearreport(ss) {
  var sheet = ss.getSheetByName('Report');
  var lastRow = sheet.getLastRow() - 3;
  var lastColumn = sheet.getLastColumn();
      
  var range = sheet.getRange(4, 1, lastRow, lastColumn);
  range.clearContent();
  sheet.setActiveRange(sheet.getRange("A4"));
};