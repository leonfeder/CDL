// This is a copy of the google script code pulled from the google sheet app..
// 
// LMF---- So I got this to work!
// LMF---- I saved the sheet with the name WorkingFormToSheet
// LMF---- I created a form file on my own hard drive. C:\Users\leon\Documents\BB\CDL19\Web\Index.html
// LMF---- I created a javascript utilties file too C:\Users\leon\Documents\BB\CDL19\Web\google-sheet.js
// LMF---- Find the text on this website: http://railsrescue.com/blog/2015-05-28-step-by-step-setup-to-send-form-data-to-google-sheets/
// LMF---- Then there's this code.  
// LMF---- 
// LMF---- Comments from the template I borrowed are included as normal comments.
// LMF---- Comments by me, look different.  // LMF----     are general comments
// LMF----                                  // LMF-app:     are application specific
//  1. Enter sheet name where data is to be written below
var SHEET_NAME = "Sheet1";                  // LMF-app:   This will serve as a log file
var DRAFT_BOARD = "DraftBoard";             // LMF-app:   This is where drafted players go
//  2. Run > setup
//
//  3. Publish > Deploy as web app
//    - enter Project Version name and click 'Save New Version'
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously)
//
//  4. Copy the 'Current web app URL' and post this in your form/script action
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service
// LMF--- Post and Get do the same thing so this funnels the input.
// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return handleResponse(e);
}
function doPost(e){
  return handleResponse(e);
}
function handleResponse(e) {
  // LockService  prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  // 
  // This code is a lock to prevent concurrent updates.
  
  var lock = LockService.getPublicLock();
  lock.waitLock(10000);  // wait 10 seconds before conceding defeat.

  // Here is where the real work gets done...
  //
  try {
    // Set where we write the data - you could write to multiple/alternate destinations
    // LMF-app: Since the skelton I used basically created a log file, which I needed anyway, we'll keep that code for that purpose.
    
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(SHEET_NAME);
    var draftboard = doc.getSheetByName(DRAFT_BOARD);
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var dbgText = "";
    var row = [];
    var thePlayer = "";
    var theTeam = "";
    var theSalary = "";
    var theContract = "";
    var theSlot = "";
    var testSlot = "";
  
    // loop through the header columns
    for (i in headers)
    {
      if (headers[i] == "Timestamp")
      { // special case if you include a 'Timestamp' column
        row.push(new Date());
      } else  // else use header name to get data
      {
        row.push(e.parameter[headers[i]]);
      }
    }
    
    // draftboard.getRange(6, 2).setValue("Start");    // LMF -dbg
    
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    
    // LMF -app: This is where I do my thing.
    // LMF -app: First, let's get proper fields...
    thePlayer = e.parameter["Player"];
    theTeam = e.parameter["Team"];
    theSalary = e.parameter["Salary"];
    theContract = e.parameter["Contract"];
    theSlot = e.parameter["Slot"];
    // LMF -app: Now we're going to loop through all the slots and derive a row number
    theRow = 31;
    for (j = 6; j < 30; j++) 
    {
      var testSlot = draftboard.getRange(j, 1).getValue();
      if (testSlot == theSlot) {theRow = j; }
    }   
    // LMF -app: Next we extract the first 2 bytes of the team name to get the team number
    // LMF -app: Then we multiply by 2 to find the column for the name.  
    // LMF -app: The column for the salary is one more.
    
    var integer = theTeam.substr(0, 2);
    var theNameCol = parseInt(integer, 10) * 2;
    var theSalCol = theNameCol + 1;
    
    // LMF -app: If the slot is occupied, we need to save the info about the player that's there
    var saveName = draftboard.getRange(theRow, theNameCol).getValue();
    var saveSal = draftboard.getRange(theRow, theSalCol).getValue();
    var saveColor = draftboard.getRange(theRow, theSalCol).getBackground();
    
    // LMF -app: Now let's store the values.
    draftboard.getRange(theRow, theNameCol).setValue(thePlayer);   
    draftboard.getRange(theRow, theSalCol).setValue(theSalary);    
    if (theContract == "L3") {
      draftboard.getRange(theRow, theSalCol).setBackground("#FFFFFF");    
      draftboard.getRange(theRow, theNameCol).setBackground("#FFFFFF"); 
    }
    else if (theContract == "L2") {
      draftboard.getRange(theRow, theSalCol).setBackground("#00FF00");    
      draftboard.getRange(theRow, theNameCol).setBackground("#00FF00");     
    }
    else if (theContract == "L1") {
      draftboard.getRange(theRow, theSalCol).setBackground("#FFFF00");    
      draftboard.getRange(theRow, theNameCol).setBackground("#FFFF00");         
    }
    else if (theContract == "S1") {
      draftboard.getRange(theRow, theSalCol).setBackground("#00FFFF");
      draftboard.getRange(theRow, theNameCol).setBackground("#00FFFF");          
    }
    else {
      draftboard.getRange(theRow, theSalCol).setBackground("#FF0000");
      draftboard.getRange(theRow, theNameCol).setBackground("#FF0000");                              
    }
      
      
    // return json success results
    return ContentService
    .createTextOutput(JSON.stringify({"result":"success", "row": nextRow, "data:": dbgText}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
    .createTextOutput(JSON.stringify({"result":"error", "error": e, "row:": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
    
  } finally { //release lock
    lock.releaseLock();
  }
}
    
function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}
