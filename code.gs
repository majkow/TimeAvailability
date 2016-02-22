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
  var lastRowTime = sheet.getRange("D1:D").getValues().filter(String).length;
  var nameRange = sheet.getRange(1, 1, lastRowName).getValues();
  var dayRange = sheet.getRange(1, 3, 7).getValues();
  var timeRange = sheet.getRange(1, 4, lastRowTime).getValues();
  Wipeform();
  Setupform(nameRange,dayRange,timeRange);
  
  Deletesheet();
  Renamesheet();
};

function Clearavailability_ () {
  var form = FormApp.openById("enter your form id here");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('rawdata');
  var lastRow =sheet.getLastRow()-1;
  
  form.deleteAllResponses();
  if(lastRow >1) {
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
  var lastRowTime = dataSheet.getRange("D1:D").getValues().filter(String).length;
  var lastRowEmail = dataSheet.getRange("E1:E").getValues().filter(String).length;
  var nameRange = dataSheet.getRange(1, 1, lastRowName).getValues();
  var timeRange = dataSheet.getRange(1, 4, lastRowTime).getValues();
  
  Makereport (ss,rawSheet,dataSheet,reportSheet,nameRange,lastRowName,timeRange);
  Emailreport(ss,reportSheet,dataSheet,lastRowEmail);
};

function Emailreport(ss,reportSheet,dataSheet,lastRowEmail) {
  var date = new Date();
  date.setDate(date.getDate() + 1);
  var day1 = Utilities.formatDate(date, "GMT+10", "E dd MMM YYYY");
  date.setDate(date.getDate() + 6);
  var nextMonday = Utilities.formatDate(date, "GMT+10", "E dd MMM YYYY");
  var emails = dataSheet.getRange(1,5,lastRowEmail).getValues();
  var emailSubject = "Coniston Availability Report for " + day1 + " to " + nextMonday;
  var emailBody = "Coniston Availability for "+ day1 + " to " + nextMonday;
  var sheets = ss.getSheets();

  //set tomorrow date on ss
  reportSheet.getRange('B1').setValue(day1);
  reportSheet.getRange('H1').setValue(nextMonday);
  //hide sheets bar report
    for(var i in sheets){
    if (sheets[i].getName()!=reportSheet.getName()){
      sheets[i].hideSheet();
    }
  }
   
 //email stuff
  var url = ss.getUrl();
  url = url.replace(/edit$/,'');
  
  /* Specify PDF export parameters
  // From: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579
    exportFormat = pdf / csv / xls / xlsx
    gridlines = true / false
    printtitle = true (1) / false (0)
    size = legal / letter/ A4
    fzr (repeat frozen rows) = true / false
    portrait = true (1) / false (0)
    fitw (fit to page width) = true (1) / false (0)
    add gid if to export a particular sheet - 0, 1, 2,..
  */
 
  var url_ext = 'export?exportFormat=pdf&format=pdf'   // export as pdf
                + '&size=A3'                       // paper size
                + '&portrait=false'                    // orientation, false for landscape
                + '&fitw=true'           // fit to width, false for actual size
                + '&sheetnames=false&printtitle=false' // hide optional headers and footers
                + '&pagenumbers=false&gridlines=true' // hide page numbers and gridlines
                + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
                + '&gid=';                             // the sheet's Id
  
  var token = ScriptApp.getOAuthToken();
  var sheets = ss.getSheets(); 
     
  //make an empty array to hold your fetched blobs  
  var blobs = [];
 
  for (var i=0; i<sheets.length; i++) {
    
    // Convert individual worksheets to PDF
    var response = UrlFetchApp.fetch(url + url_ext + sheets[i].getSheetId(), {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });
 Logger.log(response)
    //convert the response to a blob and store in our array
    blobs = response.getBlob().setName(sheets[i].getName() + '.pdf');
 
  }

  // If allowed to send emails, send the email with the PDF attachment
   GmailApp.sendEmail(emails, emailSubject, emailBody, {attachments:[blobs]});
    
  
  //show other worksheets again
  for(var i in sheets){
    if (sheets[i].getName()!=reportSheet.getName()){
      sheets[i].showSheet();
    }
  } 
};

function Makereport (ss,rawSheet,dataSheet,reportSheet,nameRange,lastRowName,timeRange) {
  //Declaring varibles 
  var avail = [];
  var k= 0;
  var row = 2;
  var repRow = 4;
  var repCol =1;
  var lastRowRaw = rawSheet.getLastRow()+1;
  
  for (var j=2; j<lastRowRaw;j++) {  //this is geoing through each row in rawdata to putthe names down
    avail = [];
    avail.push(rawSheet.getRange(j,2).getValue());
    range = reportSheet.getRange(repRow,repCol,1,43);
    
    for (var col=3;col<10;col++) { // goes though each day of the week to get the ticks.
      k= 0;
      var newValue = rawSheet.getRange(row,col).getValue().split(', ');

      //serches for Matches in the rawdata cell and puts a tick if matches and a n if not.  might change that to blank later so i don't have to double handle data
      for (var i=0;i < timeRange.length;i++) {  
        if (timeRange[i] == newValue[k]) {
          avail.push("✔");    
          k++;
        }
        else {
          avail.push(" ");
        }
      };       
    };

    range.setValues([avail]);
    repRow++;
    row++      
  };
  // Add remaining members
  var range = reportSheet.getRange(reportSheet.getLastRow()+1,1,lastRowName); 
  range.setValues(nameRange);

//  removing Duplicates
  var data = reportSheet.getRange(4,1, reportSheet.getLastRow()-3).getValues();
  var newData = new Array();
  for (h in data) {
    var newRow = data[h];
    var duplicate = false;
    for (g in newData) {
      if(newRow[0] == newData[g][0]) {
        duplicate = true;
      }
    }
    if(!duplicate) {
      newData.push(newRow);
    }
  }
  Logger.log(newData);
  
  reportSheet.getRange(4,1, reportSheet.getLastRow()-3).clear();
  reportSheet.getRange(4, 1, newData.length).setValues(newData);
  
  //sorting
  var lastRowReport = reportSheet.getLastRow();
  var lastColReport = reportSheet.getLastColumn();
  var sortRange = reportSheet.getRange(4,1,lastRowReport,lastColReport);  
  sortRange.sort({column: 1, ascending: true});
   
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
  //  This Section will do the Times as this won't change we we do the loop once and just repost the results each day. 
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
  var sheet = ss.getSheetByName('report');
  var lastRow = sheet.getLastRow() - 3;
  var lastColumn = sheet.getLastColumn();
      
  var range = sheet.getRange(4, 1, lastRow, lastColumn);
  range.clearContent();
  sheet.setActiveRange(sheet.getRange("A4"));
};