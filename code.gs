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
  
  Updateform(nameRange,dayRange,timeRange);
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
  Makepretty (ss,reportSheet);
  Emailreport(ss,reportSheet,dataSheet,lastRowEmail);
};

function Makepretty (ss,reportSheet) {
  var lastRowReport = reportSheet.getLastRow();
  var lastColReport = reportSheet.getLastColumn();
  var range ="";
 //dashing all the internal vertical and horizontal grids.
  reportSheet.getRange(4, 1, (lastRowReport-3), lastColReport).setBorder(null, null, null, null, true, true, 'black', SpreadsheetApp.BorderStyle.DASHED);
 
  //darker lines for each day to make it easier to read
  reportSheet.getRange(4, 1, (lastRowReport-3)).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 2, (lastRowReport-3),6).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 8, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 14, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 20, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
  reportSheet.getRange(4, 26, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 32, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  reportSheet.getRange(4, 38, (lastRowReport-3),6).setBorder(true, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  
  //make every second line grey
  for (var i=2;i<lastRowReport;i=i+2) {
    range= reportSheet.getRange(i, 1, 1, lastColReport);
    range.setBackgroundRGB(222, 222, 222);
  }
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

function Updateform(nameRange,dayRange,timeRange) {  //Function to edit the Form  
  var form = FormApp.openById("id of your form here");
    //edit a dropdown list with names from Spreadsheet
  var itemName = form.getItems(FormApp.ItemType.LIST);
  itemName[0].asListItem().setChoiceValues(nameRange);
   
  //Update the Days of the week
  for (var x in dayRange) {
    var checkboxItem = form.getItems(FormApp.ItemType.CHECKBOX);    
    //updates title and time slots
    checkboxItem[x].asCheckboxItem().setChoiceValues(timeRange).setTitle(dayRange[x]);
  }
}; 

function Clearreport(ss) {
  var sheet = ss.getSheetByName('report');
  var lastRow = sheet.getLastRow() - 3;
  var lastColumn = sheet.getLastColumn();
      
  var range = sheet.getRange(4, 1, lastRow, lastColumn);
  range.clearContent();
  sheet.setActiveRange(sheet.getRange("A4"));
};