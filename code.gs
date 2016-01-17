 function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Coniston')
       .addItem('Update form', 'setupAvailability_')
       .addSeparator()
       .addItem('Prepare new availability form', 'Clearavailability_')
           .addItem('Produce and email report', 'Reportmaker_')
       .addToUi();
 };
// }function onOpen() {
//  var menu = [{name: 'Update Form', functionName: 'setupAvailability_'}];
//    menu.push({name: 'Prepare new Availability Form', functionName: 'Clearavailability_'});
//  SpreadsheetApp.getActive().addMenu('Coniston', menu);
//};

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
  var form = FormApp.openById("enter your form id here");
  form.deleteAllResponses();
};

function Wipeform () {
  var form = FormApp.openById("enter your form id here");
  form.deleteAllResponses();
  // Deletes Questions
  while (form.getItems().length >0) {
  form.deleteItem(0);
  }
};

function Setupform(nameRange,dayRange,timeRange) {  //Function to create the Form to fill in 
  var form = FormApp.openById("enter your form id here")
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
    SpreadsheetApp.setActiveSheet(ss.getSheetByName('Rawdata'));
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
};

function Renamesheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet('Rawdata');
};

