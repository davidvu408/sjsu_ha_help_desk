var marketingEmail = "ha-marketing@sjsu.edu";
var directorEmail = "sheryl.spann@sjsu.edu";

var BLANK_TEMPLATE_ID = "1In-z5ruC6zu-pXwW5C8BVcvnkrDVnMbVOyOhojnH8Q4";
var TICKET_SPREADSHEET_ID = "1im0z43M_vs6T0DosEeaimP-hhbCZCjuuMoO1NSbfWNg";
var WEB_TICKETS_FOLDER_ID ="0B6Xc7NDO1JPrQXF3N1ppOXp2QVE";
var DIGITAL_DATABASE_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrbEtfOHZBcF9ib28";
var EDITORIAL_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrNGdVVmE5NlFLZGM";
var GRAPHIC_DESIGN_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrN1N5Q0hHUjY5eGc";
var MEDIA_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPra1VqdDloQS1tZGs";

var sheetID = "1im0z43M_vs6T0DosEeaimP-hhbCZCjuuMoO1NSbfWNg";
var sheetLink = "https://docs.google.com/spreadsheets/d/" + sheetID + "/edit?usp=sharing";
var sheet = SpreadsheetApp.openById(TICKET_SPREADSHEET_ID).getActiveSheet();


function  doGet(e) { 
  return HtmlService.createTemplateFromFile('index').evaluate(); 
}


/**
*  Runs upon form submission. 
*  Duplicates a blank Google Document into the corresponding folder (Web, Graphic Design, Media, Editorial, Marketing Projects) 
*  depending on the "Project Marketing Focus" and populates the ticket information onto the Google Document.
*  Parameter(s) - Event Object e 
*  
*/
function onFormSubmit(e) {
  // Populate form submission information to ticketObj{}
  var ticketObj = formResponseToObject(e);
  // Populates ticket spreadsheet with ticketObj information
  populateSpreadsheet(ticketObj);
  // Notifys client and marketingEmail that the ticket has been recieved
  sendEmailNotification(ticketObj);
}

/*
* Helper method for onFormSubmit(e). 1) Populates a Javascript Object with the information from Event Object e (form response) with 
* the form response questions as the object properties and the form response answers as the object property values. 2) Creates
* a Google Document with the populated Javascript Object information 3) Adds the Google Doc shareablelink to Javascript Object values.
* Parameter(s) - Event Object e 
* Return: Object ticketObj - Javascript Object with form response form information and Google Doc shareable link
*/
function formResponseToObject(e){
  var ticketObj = {};
  var itemResponses = e.response.getItemResponses();
  ticketObj["Email Address"] = e.response.getRespondentEmail();
  for(var i = 0; i < itemResponses.length; i++){
    var question = itemResponses[i].getItem().getTitle();
    var response = itemResponses[i].getResponse();
    if(question.indexOf("Files to attach") > -1){
      response = addOpenByIDGooglePrefix(response);
    }
    ticketObj[question] = response;
  }
  
  var ticketDocId = createTicketDocument(ticketObj);
  ticketObj["Ticket Shortlink"] = ticketDocId;
   Logger.log(ticketObj["Name of Requester"]);
  return ticketObj;
}

/*
* Helper method for formResponseToObject(e). Creates Google Document with Ticket Object recieved from formResponseToObject(e).
* Parameter(s): Object ticketObj - Javascript Object with form response form information and Google Doc shareable link
* Return: String ticketLink - URL to Google Document with ticket information
*/
function createTicketDocument(ticketObj){
  var ticketDoc;
  var ticketNumber = SpreadsheetApp.openById(TICKET_SPREADSHEET_ID).getActiveSheet().getLastRow();
  switch(ticketObj["Project Marketing Focus"]){
    case "Web":
      ticketDoc = createDuplicateDocument(BLANK_TEMPLATE_ID, WEB_TICKETS_FOLDER_ID, "Ticket #" + ticketNumber + " (Web)");
      break;
    case "Graphic Design":
      ticketDoc = createDuplicateDocument(BLANK_TEMPLATE_ID, GRAPHIC_DESIGN_TICKETS_FOLDER_ID, "Ticket #" + ticketNumber + " (Graphic Design)");
      break;
    case "Media (Photo/Video)":
     ticketDoc = createDuplicateDocument(BLANK_TEMPLATE_ID, MEDIA_TICKETS_FOLDER_ID, "Ticket #" + ticketNumber + " (Media)");
      break;
    case "Editorial and Content Development":
     ticketDoc = createDuplicateDocument(BLANK_TEMPLATE_ID, EDITORIAL_TICKETS_FOLDER_ID, "Ticket #" + ticketNumber + " (Editorial)");
      break;
    case "Digital and Database Marketing Projects":
      ticketDoc = createDuplicateDocument(BLANK_TEMPLATE_ID, DIGITAL_DATABASE_TICKETS_FOLDER_ID, "Ticket #" + ticketNumber + "( Marketing Projects)");
      break;
      
  }
  var table = ticketDoc.getBody().appendTable();
  for(var question in ticketObj){
    var tr = table.appendTableRow();
    if(question.indexOf("Files to attach") > -1){
      var td1 = tr.appendTableCell(question);
      var td2 = tr.appendTableCell();
      var idsArr = ticketObj[question].split(",");
      for(var i = 0; i < idsArr.length; i++){
        td2.appendListItem(idsArr[i]).setLinkUrl(idsArr[i]);
      }
    } else {
      var td1 = tr.appendTableCell(question);
      var td2 = tr.appendTableCell(ticketObj[question]);
    }
  }
  return "https://drive.google.com/open?id=" + ticketDoc.getId();
}

/*
* Helper function for onFormSubmit(e). Populates Ticket Spreadsheet with information from Ticket Object recieved from formResponseToObject(e).
* Parameter(s) - Javascript Object with form response form information and Google Doc shareable link
*
*/
function populateSpreadsheet(ticketObj){
  var lastRow = sheet.getLastRow() + 1;
  for(var i = 1; i <= 8; i++){
    switch(i){
      case 1:
        var emailAdressCell = sheet.getRange(lastRow, i);
        emailAdressCell.setValue(ticketObj["Email Address"]);
        break;
      case 2:
        var requestorNameCell = sheet.getRange(lastRow, i);
        requestorNameCell.setValue(ticketObj["Name of Requester"]);
        break;
      case 3:
        var phoneNumberCell = sheet.getRange(lastRow, i);
        phoneNumberCell.setValue(ticketObj["Phone Number"]);
        break;
      case 4:
        var projectDeadlineCell = sheet.getRange(lastRow, i);
        projectDeadlineCell.setValue(ticketObj["Project Deadline"]);
        break;
      case 5:
        var projectFocusCell = sheet.getRange(lastRow, i);
        projectFocusCell.setValue(ticketObj["Project Marketing Focus"]);
        break;
      case 6:
        var projectDescriptionCell = sheet.getRange(lastRow, i);
        projectDescriptionCell.setValue(ticketObj["Project Description"]);
        break;
      case 7:
        var ticketShortlinkCell = sheet.getRange(lastRow, i);
        var ticketShortlink = ticketObj["Ticket Shortlink"];
        var displayName = "View Ticket #" + (lastRow - 1);
        ticketShortlinkCell.setFormula("=HYPERLINK(\"" + ticketShortlink + "\";\"" + displayName + "\")");
        break;
      case 8:
        var statusCell = sheet.getRange(lastRow, i);
        statusCell.setValue("New");
        break;
    }
  }

}
/*
* Helper method for onFormSubmit(e). Sends email notification to client and marketingEmail that the ticket has been recieved
* Parameter(s) - Event Object e
*/
function sendEmailNotification(ticketObj){
  var lastRow = sheet.getLastRow();
  // Sends confirmation email to client
  var clientName = ticketObj["Name of Requester"];
  var clientEmail = ticketObj["Email Address"];
  MailApp.sendEmail(clientEmail,
                    "Help Desk New Ticket #" + (lastRow - 1),
                    clientName + ",\n\n" +
                    "Thank you for your submission.\n" +
                    "Your ticket has been recorded. This is your confirmation email.\n\n" +
                    "Your ticket will be assigned to one of our team members and you will be contacted soon via your SJSU email regarding updates.\n\n" +
                    "-H&A Marketing Team",
                    {name:"H&A Help Desk", replyTo: marketingEmail});
  
  // Finds email target
  var id = "";
  var body = "A new ticket has been submitted through the Help Desk Ticketing System.\n\nYou can view this at " + sheetLink + "\n\n";
  var projectFocus = ticketObj["Project Marketing Focus"];
  if(projectFocus == "Web"){
    id += "Website";
  }else if(projectFocus == "Graphic Design"){
    id += "Graphic";
  }else if(projectFocus == "Media (Photo/Video)"){
    id += "Photo+Video";
  }else if(projectFocus == "Editorial and Content Development"){
    id += "Editing";
  }else if(projectFocus == "Digital and Database Marketing Projects"){
    MailApp.sendEmail(directorEmail, 
                      "Help Desk New Ticket #" + lastRow,
                      body + "\n\n" + "-H&A Marketing Team",
                      {name:"H&A Help Desk"});
  }
  
  // Sends confirmation email to H&A Marketing Team shared inbox
  MailApp.sendEmail("ha-marketing" + "+" + id + "@sjsu.edu",
                    "Help Desk New Ticket #" + (lastRow - 1),
                    body + "\n\n" + "-H&A Marketing Team",
                    {name:"H&A Help Desk"});       
  
}

/*
*  Helper function for onFormSubmit(e). Takes IDs from "Files to attatch" question and adds the open by ID url prefix to each one.
*  Parameter(s): Array sourceIds - IDs to files on Google Drive
*  Return: String of all IDs with Google's open by ID url prefix concatenated
*/
function addOpenByIDGooglePrefix(sourceIds){
  // Adds Google Open By ID Prefix to each ID
  for(var i = 0; i < sourceIds.length; i++){
    sourceIds[i] = "https://drive.google.com/open?id=" + sourceIds[i];
  }
  return sourceIds.join();
}

/**
 * Duplicates a Google Apps doc
 * Parameter(s): sourceId - ID of document to be duplicated ID
 *               targetFolder - ID of folder where the document is going to be placed
 *               name - Name of newly copied document
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(sourceId, targetFolder, name) {
    var source = DriveApp.getFileById(sourceId);
    var targetFolder = DriveApp.getFolderById(targetFolder);
    var newFile = source.makeCopy(name, targetFolder);
    return DocumentApp.openById(newFile.getId());
}

