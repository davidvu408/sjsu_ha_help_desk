// During development: Disable automatic email update to marketing director
// This script requires the following triggers: 
           // 1) onEdit() - From spreadsheet -> On edit 
           // 2) formSubmitReply() -> From spreadsheet -> On form submit
           // 3) onOpen() -> From spreadsheet -> On open
var colorWhite = "#FFFFFF"; 
var colorGrey = "#CCCCCC";
var colorRed = "#F00";
var sheet= SpreadsheetApp.getActiveSheet();
var marketingEmail = "ha-marketing@sjsu.edu";
var directorEmail = "sheryl.spann@sjsu.edu";
var sheetID = "1GzTILJ8iXWr-NobHCnMRgNCSb_0s3n43ScpZ4QkcD-s";
var sheetLink = "https://docs.google.com/a/sjsu.edu/spreadsheets/d/" + sheetID + "/edit?usp=sharing";

var BLANK_TEMPLATE_ID = "1In-z5ruC6zu-pXwW5C8BVcvnkrDVnMbVOyOhojnH8Q4";
var WEB_TICKETS_FOLDER_ID ="0B6Xc7NDO1JPrQXF3N1ppOXp2QVE";
var DIGITAL_DATABASE_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrbEtfOHZBcF9ib28";
var EDITORIAL_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrNGdVVmE5NlFLZGM";
var GRAPHIC_DESIGN_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPrN1N5Q0hHUjY5eGc";
var MEDIA_TICKETS_FOLDER_ID = "0B6Xc7NDO1JPra1VqdDloQS1tZGs";


/*
* @Override
* Runs when spreadsheet opens
*
*/
function onOpen() {
  var subMenu = [{name:"Delete all responses", functionName: "deleteResponses"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("More", subMenu);
  
}

function deleteResponses(){
  SpreadsheetApp.getUi().alert("OK");
}

/*
* @Override
*  Runs when an edit is made to a cell
*  Parameter(s) - Event Object e is the cell that was edited
*/
function onEdit(e) {
   // If edit was made to Status and is now Assigned => Turn the "Assigned" cell to the right of it red until it is filled
  changeAssignedCellBackground(e);
  // If edit was made to Status and Resolved => Darken that row 
  darkenResolvedTickets(e);
  // If edit was made to Status as Assigned or In Progress => Give assigned employee/client notification email
  assignedEmailStatusUpdate(e); 
  // If edit was made to Status and anything but Resolved => Whiten that row 
  resetStatusColor(e);
}

// Set ticket's rows that are anything but "Resolved" to White #FFFFFF
function resetStatusColor(e) {
  var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
   // If the edited cell was in the status column
  if(e.range.getColumn() == statusColumnIndex){
    // If New then turn that column white 
      if(e.range.getValue() == "New"){
        sheet.getRange(e.range.getRow(), 1,1,sheet.getLastColumn()).setBackground(colorWhite); // Changes ticket row to White
    } 
  }
}

// Set ticket's rows that are "Resolved" to Grey #CCCCCC and sends email to client notifying them that their ticket has been resolved
function darkenResolvedTickets(e) {
  var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
  // If the edit was in the "Status" column and value is "Resolved"
  if(e.range.getColumn() == statusColumnIndex && e.range.getValue() == "Resolved"){
       var row = e.range.getRow();
       var assignedEmployee = sheet.getRange(row, getColIndexByName("Assigned", sheet)).getValue();
       var assignedEmail = getEmployeeEmail(assignedEmployee);
       var client = sheet.getRange(row, getColIndexByName("Name of Requester", sheet)).getValue();
       var clientEmail = sheet.getRange(row, getColIndexByName("Email Address", sheet)).getValue(); 
       var projectFocus = sheet.getRange(row, getColIndexByName("Project Marketing Focus", sheet)).getValue();
    
       var id = "";
       if(projectFocus == "Web"){
         id = "Web";
       }else if(projectFocus == "Graphic Design"){
         id = "Graphic Design";
       }else if(projectFocus == "Media (Photo/Video)"){
         id = "Media";
       }else if(projectFocus == "Editorial and Content Development"){
         id = "Editorial";
       }else if(projectFocus == "Digital and Database Marketing Projects"){
         id = "Digital and Database";
       }
       var issue = sheet.getRange(row, getColIndexByName(id + " Project Description", sheet)).getValue();  
      
     
       var response = SpreadsheetApp.getUi().prompt("Confirm", 
                               "This ticket has been resolved " + client + " (" + clientEmail + ")" +
                               " will recieve a notification email. Any personal notes? (Please click okay regardless so that " + 
                                client + " will recieve a confirmation email stating that their ticket has been resolved)", 
                               SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      var buttonClicked = response.getSelectedButton();
      var employeeNotes = response.getResponseText();
      if(buttonClicked == SpreadsheetApp.getUi().Button.OK) {
        sheet.getRange(e.range.getRow(), 1,1,sheet.getLastColumn()).setBackground(colorGrey); // Changes ticket (now resolved) row to Grey
        MailApp.sendEmail(clientEmail,
                         "HA Marketing Help Desk Ticket #" + row,
                          client +",\n\n" +
                         "Your ticket has now been resolved!\n" +
                         "Issue Description: " + issue + "\n" +
                         "Personal notes from " + assignedEmployee + ": " + employeeNotes + "\n" +
                          "Replying to this email will go to: " + assignedEmployee + "\n\n" +
                          "-H&A Marketing Team",
                          {name: "H&A Help Desk", replyTo: assignedEmail});
     } // End if "OK" button
  } // End if "Status" column
} /// End darkenResolvedTickets()

// Set ticket's "Assigned" cell to red until it is filled
function changeAssignedCellBackground(e){
  var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
  // If the edited cell was in the status column
  if(e.range.getColumn() == statusColumnIndex && e.range.getValue() == "Assigned"){
      sheet.getRange(e.range.getRow(), getColIndexByName("Assigned", sheet)).setBackground(colorRed); // Turns "Assigned" cell to the right red
  }
}

function assignedEmailStatusUpdate(e) { 
  // If edited cell was in "Assigned" column
  if(e.range.getColumn() == assignedColumnIndex) {
    var employeeDataSheet = SpreadsheetApp.openById(sheetID).getSheetByName("Employee Data"); // Employee Data Sheet
    var row = sheet.getActiveRange().getRowIndex();
    var client = sheet.getRange(row, getColIndexByName("Name of Requester", sheet)).getValue();
    var clientEmail = sheet.getRange(row, getColIndexByName("Email Address", sheet)).getValue(); 
    var assignedEmployee = sheet.getRange(row, getColIndexByName("Assigned", sheet)).getValue();
    var assignedEmail = getEmployeeEmail(assignedEmployee);
    var projectFocus = sheet.getRange(row, getColIndexByName("Project Marketing Focus", sheet)).getValue();
    var assignedColumnIndex = getColIndexByName("Assigned", sheet); // Column index for "Assigned"
    var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
    // Render a dialog message confirming the email update to be sent to the assigned employee
    var response = SpreadsheetApp.getUi().alert(
          'Confirm',
          assignedEmail + ' will recieve a confirmation email that he has been assigned to this ticket. (Please click okay to notify ' + 
          client + ' that their ticket has been assigned and will be resolved shortly)',
          SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
       
    if(response == SpreadsheetApp.getUi().Button.OK){
      sheet.getRange(row, assignedColumnIndex).setBackground("#FFFFFF");
        // Finds issue description according to "Project Marketing Focus"
      var id = "";
      if(projectFocus == "Web"){
        id = "Web";
      }else if(projectFocus == "Graphic Design"){
        id = "Graphic Design";
      }else if(projectFocus == "Media (Photo/Video)"){
        id = "Media";
      }else if(projectFocus == "Editorial and Content Development"){
        id = "Editorial";
      }else if(projectFocus == "Digital and Database Marketing Projects"){
        id = "Digital and Database";
      }
         
      var issue = sheet.getRange(row, getColIndexByName(id + " Project Description", sheet)).getValue();  
      
      // Notify employee that he/she has been assigned to a ticket
      MailApp.sendEmail(assignedEmail,
                        "HA Marketing Help Desk Ticket #" + row,
                        "You have been assigned to a ticket via the H&A Help Desk ticketing system\n"+
                        "Client: " + client + " (" + clientEmail + ")\n"+
                        "Issue Description: " + issue + "\n"+
                        "You can view this at: " + sheetLink+"\n"+
                        "Replying to this email will go to: " + client + "\n\n" +
                        "-H&A Marketing Team",
                        {name:"H&A Help Desk", replyTo: clientEmail});  
      
      // Notify client that ticket has been assigned to an employee
      MailApp.sendEmail(clientEmail,
                        "HA Marketing Help Desk Ticket #" + row,
                        client + ",\n\n" +
                        "The status of your ticket is currently: Assigned\n"+ 
                        "You are assigned to: " + assignedEmployee +", "+assignedEmail +"\n"+
                        "Issue Description: " + issue +"\n"+
                        "Replying to this email will go to: " + assignedEmail+"\n\n"+
                        "-H&A Marketing Team",
                        {name:"H&A Help Desk", replyTo: assignedEmail});
      } // end if Button OK 
    

  } 
  // "In-Progress" has been discontinued
  /**
  else if (e.range.getColumn() == statusColumnIndex && e.range.getValue() == "In Progress") {
    var response = SpreadsheetApp.getUi().alert(
          'Confirm',
          client + " (" + clientEmail + ")" + " will recieve a confirmation email that his ticket is in the process of being resolved.",
          SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if(response == SpreadsheetApp.getUi().Button.OK) {
      // Notify client that their ticket is in the process of being resolved
      MailApp.sendEmail(clientEmail,
                        "HA Marketing Help Desk Ticket #" + row,
                        client + ",\n\n" +
                        "The status of your ticket is currently: In Progress\n" + 
                        assignedEmployee + " is assigned to this ticket\n" +
                        "Issue Description: " + issue +"\n"+
                        "Replying to this email will go to: " + assignedEmail + "\n\n"+
                        "-H&A Marketing Team",
                        {name:"H&A Help Desk", replyTo: assignedEmail});
    } // end if Button OK
  } // end if/else-if 
  **/
}


/**
*  Event function - Will run when client submits form
*  Description - Function will send a confirmation email to client and H&A email. Triggered onFormSubmit (see project triggers).
*  Parameter(s) - Event Object e is the form
*/
function formSubmitReply(e) {
  // Initialize new ticket on spreadsheet
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, getColIndexByName("Status", sheet)).setValue("New");
  
  // Sends confirmation email to client
  var userNameIndex = getColIndexByName("Name of Requester", sheet);
  var userEmailIndex = getColIndexByName("Email Address", sheet);
  var userName = e.values[userNameIndex - 1];
  var userEmail = e.values[userEmailIndex - 1];
   MailApp.sendEmail(userEmail,
                    "Help Desk New Ticket #" + lastRow,
                    userName + ",\n\n" +
                    "Thank you for your submission.\n" +
                    "Your ticket has been recorded. This is your confirmation email.\n\n" +
                    "Your ticket will be assigned to one of our team members and you will be contacted soon via your SJSU email regarding updates.\n\n" +
                    "-H&A Marketing Team",
                    {name:"H&A Help Desk", replyTo: marketingEmail});
  
  // Finds email target
  var id = "";
  var body = "A new ticket has been submitted through the Help Desk Ticketing System.\n\nYou can view this at " + sheetLink + "\n\n";
  if(e.namedValues["Project Marketing Focus"] == "Web"){
    id += "Website";
  }else if(e.namedValues["Project Marketing Focus"] == "Graphic Design"){
    id += "Graphic";
  }else if(e.namedValues["Project Marketing Focus"] == "Media (Photo/Video)"){
    id += "Photo+Video";
  }else if(e.namedValues["Project Marketing Focus"] == "Editorial and Content Development"){
    id += "Editing";
  }else if(e.namedValues["Project Marketing Focus Focus"] == "Digital and Database Marketing Projects"){
    MailApp.sendEmail(directorEmail, 
                      "Help Desk New Ticket #" + lastRow,
                      body + "\n\n" + "-H&A Marketing Team",
                      {name:"H&A Help Desk"});
  }
  
  // Sends confirmation email to H&A Marketing Team shared inbox
  MailApp.sendEmail("ha-marketing" + "+" + id + "@sjsu.edu",
                    "Help Desk New Ticket #" + lastRow,
                    body + "\n\n" + "-H&A Marketing Team",
                    {name:"H&A Help Desk"});       
}

/*
*  Returns an employee email from "Employee Data" sheet
*  Parameter(s): String employee, name of employee on Employee Data sheet
*  Return: String email of employee
*/
function getEmployeeEmail(employee){
  var employeeNameIndex = 1;
  var employeeEmailIndex = 2;
  var employeeData = SpreadsheetApp.openById(sheetID).getSheetByName("Employee Data"); // Employee Data Sheet
   for(var i=2; i <= employeeData.getLastRow(); i++){
    // Search for employee name down the name column
    if( employeeData.getRange(i,employeeNameIndex).getValue() == employee){
      // Once the name is found, capture the employee email in the email column
       return employeeData.getRange(i,employeeEmailIndex).getValue();
    }
  }
  // -1 returned if employee name is not found
  return -1;
}

/*
*  Gets the index of specified column index by the name of the column
*  Parameter(s): String colname, name of the column
*  Return: int index, index of the column
*/
function getColIndexByName(colName, sheet) {
  var numColumns = sheet.getLastColumn();
  // Get header row
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  // Index through the header row
  for (i in row[0]) {
    var name = row[0][i];
    // Transverse down the row to each column
    if (name == colName) {
      // If the coloumn we are searching for matches the index
      return parseInt(i) + 1;
      // Return the index of that row offset
    }
  }
  return -1;
}