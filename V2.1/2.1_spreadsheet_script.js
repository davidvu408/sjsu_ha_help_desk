// This script requires the following triggers: 
           // 1) onEdit() - From spreadsheet -> On edit 


var colorWhite = "#FFFFFF"; 
var colorGrey = "#CCCCCC";
var colorRed = "#F00";
var sheet= SpreadsheetApp.getActiveSheet();
var marketingEmail = "ha-marketing@sjsu.edu";
var directorEmail = "sheryl.spann@sjsu.edu";
var sheetID = "1im0z43M_vs6T0DosEeaimP-hhbCZCjuuMoO1NSbfWNg";
var sheetLink = "https://docs.google.com/a/sjsu.edu/spreadsheets/d/" + sheetID + "/edit?usp=sharing";

/*
* @Override
*  Runs when an edit is made to a cell
*  Parameter(s) - Event Object e is the cell that was edited
*/
function onEdit(e) {
   // If edit was made to Status and is now Assigned => Turn the "Assigned" cell to the right of it red until it is filled
  changeAssignedCellBackground(e);
  // If edit was made to Status as Assigned or In Progress => Give assigned employee/client notification email
  assignedEmailStatusUpdate(e); 
  // If edit was made to Status and Resolved => Darken that row 
  darkenResolvedTickets(e);
  // If edit was made to Status and anything but Resolved => Whiten that row 
  resetStatusColor(e);
}

/*
* Helper method for onEdit(e). When ticket status is changed to "Assigned", the assigned cell to the right will be changed to Red.
* Parameter(s) - Event Object e is the cell that was edited
*/
function changeAssignedCellBackground(e){
  var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
  // If the edited cell was in the status column
  if(e.range.getColumn() == statusColumnIndex && e.range.getValue() == "Assigned"){
      sheet.getRange(e.range.getRow(), getColIndexByName("Assigned", sheet)).setBackground(colorRed); // Turns "Assigned" cell to the right red
  }
}

/*
* Helper method for onEdit(e). When ticket becomes assigned to an employee,this function will send email notifications to the client and the employee
* Parameter(s) - Event Object e is the cell that was edited
*/
function assignedEmailStatusUpdate(e) { 
  var assignedColumnIndex = getColIndexByName("Assigned", sheet); // Column index for "Assigned"
  // If edited cell was in "Assigned" column
  if(e.range.getColumn() == assignedColumnIndex) {
    var employeeDataSheet = SpreadsheetApp.openById(sheetID).getSheetByName("Employee Data"); // Employee Data Sheet
    var row = sheet.getActiveRange().getRowIndex();
    var client = sheet.getRange(row, getColIndexByName("Name of Requestor", sheet)).getValue();
    var clientEmail = sheet.getRange(row, getColIndexByName("Email Address", sheet)).getValue(); 
    var assignedEmployee = sheet.getRange(row, getColIndexByName("Assigned", sheet)).getValue();
    var assignedEmail = getEmployeeEmail(assignedEmployee);
    var projectFocus = sheet.getRange(row, getColIndexByName("Project Marketing Focus", sheet)).getValue();
   
    var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
    var issue = sheet.getRange(row, getColIndexByName("Project Description", sheet)).getValue();  
    
    sheet.getRange(row, assignedColumnIndex).setBackground("#FFFFFF");
 
    // Notify employee that he/she has been assigned to a ticket
    MailApp.sendEmail(assignedEmail,
                      "HA Marketing Help Desk Ticket #" + (row - 1),
                      "You have been assigned to a ticket via the H&A Help Desk ticketing system\n"+
                      "Client: " + client + " (" + clientEmail + ")\n"+
                      "Issue Description: " + issue + "\n"+
                      "You can view this at: " + sheetLink+"\n"+
                      "Replying to this email will go to: " + client + "\n\n" +
                      "-H&A Marketing Team",
                      {name:"H&A Help Desk", replyTo: clientEmail});  

    // Notify client that ticket has been assigned to an employee
    MailApp.sendEmail(clientEmail,
                      "HA Marketing Help Desk Ticket #" + (row - 1),
                      client + ",\n\n" +
                      "The status of your ticket is currently: Assigned\n"+ 
                      "You are assigned to: " + assignedEmployee +", "+assignedEmail +"\n"+
                      "Issue Description: " + issue +"\n"+
                      "Replying to this email will go to: " + assignedEmail+"\n\n"+
                      "-H&A Marketing Team",
                      {name:"H&A Help Desk", replyTo: assignedEmail});

     var response = SpreadsheetApp.getUi().alert(
        'Notification',
         assignedEmployee + ' has recieved a confirmation email that he/she has been assigned to this ticket and ' + 
         client + ' has been notified that their ticket has been assigned and will be resolved shortly',
         SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  } 
}

/*
* Helper function for onEdit(e). When ticket status is changed to "Resolved", this functioin will turn the ticket row on the spreadsheet Grey.
* Parameter(s) - Event Object e is the cell that was edited
*/
function darkenResolvedTickets(e) {
  var statusColumnIndex = getColIndexByName("Status", sheet); // Column index for "Status"
  // If the edit was in the "Status" column and value is "Resolved"
  if(e.range.getColumn() == statusColumnIndex && e.range.getValue() == "Resolved"){
       var row = sheet.getActiveRange().getRowIndex();
       var assignedEmployee = sheet.getRange(row, getColIndexByName("Assigned", sheet)).getValue();
       var assignedEmail = getEmployeeEmail(assignedEmployee);
       var client = sheet.getRange(row, getColIndexByName("Name of Requestor", sheet)).getValue();
       var clientEmail = sheet.getRange(row, getColIndexByName("Email Address", sheet)).getValue(); 
       var projectFocus = sheet.getRange(row, getColIndexByName("Project Marketing Focus", sheet)).getValue();
       var issue = sheet.getRange(row, getColIndexByName("Project Description", sheet)).getValue();  

        sheet.getRange(e.range.getRow(), 1,1,sheet.getLastColumn()).setBackground(colorGrey); // Changes ticket (now resolved) row to Grey
        MailApp.sendEmail(clientEmail,
                         "HA Marketing Help Desk Ticket #" + (row - 1),
                          client +",\n\n" +
                         "Your ticket has now been resolved!\n" +
                         "Issue Description: " + issue + "\n" +
                         "Replying to this email will go to: " + assignedEmployee + "\n\n" +
                         "-H&A Marketing Team",
                         {name: "H&A Help Desk", replyTo: assignedEmail});

  } // End if "Status" column
} /// End darkenResolvedTickets()

// Set ticket's rows that are anything but "Resolved" to White #FFFFFF
/*
* Helper function for onEdit(e). When ticket status is changed to "New", this functioin will turn the ticket row on the spreadsheet White.
* Parameter(s) - Event Object e is the cell that was edited
*/
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