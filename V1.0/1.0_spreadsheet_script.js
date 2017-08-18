/*
*  Author: Aditya Shah
*/
var statusColumnIndex = getColIndexByName("Status", SpreadsheetApp.getActiveSheet());

/*
* @Override
* Triggered when document opens, will add "Help Desk Menu" button
*/
function onOpen() {
  // Creates and adds options to "Help Desk Menu" 
  var subMenus = [{name:"Send Status Email", functionName: "emailStatusUpdates"},{name:"Generate Time Data", functionName:"collectTime"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Help Desk Menu", subMenus);
}

/*
* @Override
* Runs when someone changes the spreadsheet
* 
*/
function onEdit (e){
  resetStatusColor(e);
  darkenResolvedTickets(e);
  createLinks();
}

/*
* 
*/
function resetStatusColor(e){
  //if the edited cell was the status column
  if(e.range.getColumn() == statusColumnIndex){

    var sheet= SpreadsheetApp.getActiveSheet();
    //set the background back to white of the status cell since the status has changed
    sheet.getRange(e.range.getRow(), e.range.getColumn()).setBackground("white");
  }
}

function darkenResolvedTickets(e){
  //if the edit was in the status column
  if(e.range.getColumn() == statusColumnIndex){
    //if the status was set to Resolved in that column
    if(e.range.getValue() == "Resolved"){
      var sheet = SpreadsheetApp.getActiveSheet();
      //iterate through each cell in row and set background color to gray
      for(var i=0; i<sheet.getLastColumn();i++){
        sheet.getRange(e.range.getRow(), i+1).setBackground("#CCCCCC");
      }
    }
  }
}

//Description: Gets the index of specified column index by the name of the column
//Parameters: String colname, name of the column
//Returns: int index, index of the column
function getColIndexByName(colName, sheet) {
  var numColumns = sheet.getLastColumn();
  //get header row
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  //index through the header row
  for (i in row[0]) {
    var name = row[0][i];
   
    //transverse down the row to each column
    if (name == colName) {
      //if the coloumn we are searching for matches the index
      return parseInt(i) + 1;
      //return the index of that row offset
    }
  }
  return -1;
}

//Event function
//This function will run automatically when someone submits the form
//Parameters: e, contains the event information for the submission
function formSubmitReply(e) {
  //sheet and index intialization
  var sheet = SpreadsheetApp.getActiveSheet();
  var marketingEmail = "ha-marketing@sjsu.edu";
  var userEmailIndex = getColIndexByName("Username", sheet);
  var userNameIndex = getColIndexByName("Name", sheet);
  var userFocusIndex = getColIndexByName("Focus", sheet);
  var userPriorityIndex = getColIndexByName("Priority", sheet);
  var userIssueIndex = getColIndexByName("Issue Description", sheet);
   
  
  
  var userEmail = e.values[userEmailIndex-1];
  var lastRow = sheet.getLastRow();
  
  //set formatting for new ticket/row submitted
  sheet.getRange(lastRow, getColIndexByName("Status", sheet)).setValue("New");
  sheet.getRange(lastRow, getColIndexByName("Issue Description", sheet)).setWrap(true);
  sheet.getRange(lastRow, getColIndexByName("Notes", sheet)).setWrap(true);
  sheet.getRange(lastRow, getColIndexByName("Resolution", sheet)).setWrap(true);
  sheet.getRange(lastRow, userFocusIndex).setWrap(true);
  sheet.setRowHeight(lastRow, 75);
  
  // Sends confirmation email
  MailApp.sendEmail(userEmail,
                    "HA Marketing Help Desk Ticket #" + lastRow,
                    "Thank you for your submission.\n"+
                    "Your ticket has been recorded. This is your confirmation email.\n\n" +
                    "Your ticket will be assigned to one of our team members and you will be\ncontacted soon via your SJSU email regarding updates.",
                    {name:"Help Desk", replyTo: marketingEmail});
  
  // Sends follow up email to ha-marketing
  MailApp.sendEmail(marketingEmail,
                    "Help Desk New Ticket #" + lastRow,
                    "A new ticket has been submitted through the Help Desk Ticketing System\n\n"+
                    "Name: "+e.values[userNameIndex-1]+", "+e.values[userEmailIndex-1]+"\n"+
                    "Focus: "+e.values[userFocusIndex-1]+"\n"+
                    "Priority: "+e.values[userPriorityIndex-1]+"\n\n"+
                    "Issue Description: "+e.values[userIssueIndex-1]+"\n\n"+
                    "Replying to this email will go to "+e.values[userNameIndex-1],
                    {name:"Help Desk", replyTo: userEmail});
}

function emailStatusUpdates() {
  //index constants
  var sheet = SpreadsheetApp.getActiveSheet();
  var employeeEmailIndex = getColIndexByName("Username", sheet); // This should be deleted?
  var employeeNameIndex = getColIndexByName("Name", sheet);
  
  var row = sheet.getActiveRange().getRowIndex();
  
  //openByID([PUT TICKET SPREADSHEET ID HERE])
  var employee = SpreadsheetApp.openById("1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk").getSheetByName("Employee Data"); // Employee Data Sheet
   
  //grab ticket data
  var userEmail = sheet.getRange(row, getColIndexByName("Email Address", sheet)).getValue();//Username was changed to "Email Address" because column name "Usernadoes not exist 
  var status = sheet.getRange(row, getColIndexByName("Status", sheet)).getValue();
  var assigned = sheet.getRange(row, getColIndexByName("Assigned", sheet)).getValue();
  var notes = sheet.getRange(row, getColIndexByName("Notes", sheet)).getValue();
  var resolution = sheet.getRange(row, getColIndexByName("Resolution", sheet)).getValue();
  var priority = sheet.getRange(row, getColIndexByName("Priority", sheet)).getValue();
  var name = sheet.getRange(row, getColIndexByName("Name", sheet)).getValue();
  var issue = sheet.getRange(row, getColIndexByName("Issue Description", sheet)).getValue();
  var assignedEmail = getEmployeeEmail(assigned);
  
  //Notify assigned employee
  if(status == "Assigned"){
    MailApp.sendEmail(assignedEmail,
                      "HA Marketing Help Desk Ticket #"+row,
                      "You have been assigned to a ticket via the ticketing system\n"+
                      "Name: "+name+", "+userEmail+"\n"+
                      "Priority: "+priority+"\n\n"+
                      "Issue Description: "+issue+"\n\n"+
                      "Replying to this email will go to " + name,
                      {name:"Help Desk", replyTo: userEmail});
  }  
  
  //Start generating email based on ticket status
  
  var body = "Your ticket has been updated recently.\n\n";
  
  if(status == "Assigned"){  
    body += "The status of your ticket is currently: Assigned\n"
    + "You are assigned to: " + assigned +", "+assignedEmail +"\n\n"
    + "Issue Description: " + issue +"\n\n"
    + "Replying to this email will go to "+assigned;
    sendUpdate(userEmail, row, assignedEmail, body);
  }else if(status == "In Progress"){
    body +=  "The status of your ticket is currently: In Progress\n"
    + assigned + " is in progress solving or finding a solution to your ticket\n\n"
    + "Issue Description: " + issue + "\n"
    + "Notes: " + notes +"\n\n"
    + "Replying to this email will go to "+assigned;
    sendUpdate(userEmail, row, assignedEmail, body);
  }else if(status == "Resolved"){
    body+= "Your ticket has now been resolved!\n\n"
    + "Issue Description: " + issue + "\n"
    + "Notes: "+notes+"\n"
    + "Resolution: "+resolution+"\n\n"
    + "Replying to this email will go to "+assigned;
    sendUpdate(userEmail, row, assignedEmail, body);
  }else{
    Browser.msgBox('Your ticket\'s "Status" is not recognized', Browser.Buttons.OK_CANCEL);
  }
  //set to red so that employees know the email has been sent for this status
  sheet.getRange(row, getColIndexByName("Status", sheet)).setBackground("red");
}
//this function looks up employee email in the employee data sheet
function getEmployeeEmail(employee){
  var employeeNameIndex = 1;
  var employeeEmailIndex = 2;
  //load spreadsheet
  var employeeData = SpreadsheetApp.openById("1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk").getSheetByName("Employee Data");
   for(var i=2; i < employeeData.getLastRow(); i++){
    //search for employee name down the name column
    if( employeeData.getRange(i,employeeNameIndex).getValue() == employee){
      //once the name is found capture the employee email in the email column
       return employeeData.getRange(i,employeeEmailIndex).getValue();
    }
  }
  //-1 returned if not found employee name
  return -1;
}
//This function sends the user an email based on their email, ticket number, the body (generated earilier) and the assigned employee's email
function sendUpdate(userEmail, ticket, replyto, message){
  var subject = "HA Marketing Help Desk Ticket #" + ticket;
  MailApp.sendEmail(userEmail, subject, message, {name:"Help Desk", replyTo: replyto});
}

function collectTime(){
//issues accessing id  
  var id = "1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk";
  //open ticket data
  var tickets = SpreadsheetApp.openById("1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk").getSheetByName("Tickets");
  var internal = SpreadsheetApp.openById("1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk").getSheetByName("Internal Tickets");
  //open time sheet where data will be stored after query
  var data = SpreadsheetApp.openById(id).getSheetByName("Time");
  //get the day this script has been ran
  var day = Utilities.formatDate(new Date(), "PST", "MM-dd-yyyy");
  var lastRow = tickets.getLastRow();
  var internalLastRow = internal.getLastRow();
  //get focuses, times, and department columns as arrays for internal tickets
  var internalFocuses = internal.getRange(2, getColIndexByName("Focus", internal) , internalLastRow-1).getValues();
  var internalTimes = internal.getRange(2, getColIndexByName("Time (Hrs)", internal), internalLastRow-1).getValues();
  var internalDepts = internal.getRange(2, getColIndexByName("School/Department", internal), internalLastRow-1).getValues();
  
  //get the entire focus, time, and department columns from ticket sheet as arrays. Then concatenate the internal tickets to this data
  //This will give us three arrays of all ticket data internal and normal tickets alike
  var focuses = tickets.getRange(2, getColIndexByName("Focus",tickets), lastRow-1).getValues().concat(internalFocuses);
  var times = tickets.getRange(2, getColIndexByName("Time (Hrs)",tickets), lastRow-1).getValues().concat(internalTimes);
  var depts = tickets.getRange(2, getColIndexByName("School/Department",tickets), lastRow-1).getValues().concat(internalDepts);
  
  //focusTotals and deptTotals are associative arrays
  //Keys:focuses Values:summed times
  var focusTotals = [];
  var deptTotals = [];
  //initialize focusTotals array
  for(var i=0; i<focuses.length; i++){
  
    focusTotals[focuses[i]] = 0; 
  }
  //intialize deptTotals array
  for(var i=0; i<depts.length; i++){
    
    deptTotals[depts[i]] = 0;
  }
  
  //step through each ticket based on focus it was assigned to
  for(var i=0; i<focuses.length; i++){
    //if the time for that ticket has been filled out 
    if(times[i] != ""){
      //add that time to the total for that focus and dept, times are read as strings so compute them to float for summation
      focusTotals[focuses[i]] += parseFloat(times[i]);
      deptTotals[depts[i]] += parseFloat(times[i]);
    }
  }
  //FOCUS TIMES TOTALS
  //print the timetotals to the data collection sheet for each focus
  var focus = ["Web Design","Graphic Design","Video and Photo","Editor","Administration"];
  for(var j=0; j<focus.length; j++){ 
    
    //print day
    data.getRange(j+2, 1).setValue(day);
    //print focus
    data.getRange(j+2, 2).setValue(focus[j]);
    //match timetotal to focus and print to spreadsheet
    if(focusTotals[focus[j]] != null){
      data.getRange(j+2, 3).setValue(focusTotals[focus[j]]);
    }
    else{
      //print zero if there is no time submitted for that focus
      data.getRange(j+2, 3).setValue(0);
    }
  }
  //DEPARTMENT TIME TOTALS
  //print time totals by department
  var dept = ["Art & Art History","Dean's Office","Design","English & Comparative Literature","Humanities","Linguistics & Language Development",
              "Philosophy","School of Music & Dance","Television Radio Film & Theatre","World Languages & Literature"];
  var last = 7;
 
  for(var j=0; j<dept.length; j++){
   
    //print day
    data.getRange(j+last, 1).setValue(day);
    //print department
    data.getRange(j+last, 2).setValue(dept[j]);
    //match totaltime to the department and print to spreadsheet
    if(deptTotals[dept[j]] != null){
      data.getRange(j+last, 3).setValue(deptTotals[dept[j]]);
    }
    else{
      //print zero if there is no time submitted for that focus
      data.getRange(j+last, 3).setValue(0);
    }
  }
}

//Turns the confirmation code given by the users into a link to find the files
function createLinks(){
  //get the range for the "Uploading Files" collumn
  var tickets = SpreadsheetApp.openById("1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk").getSheetByName("Tickets");
  var range = tickets.getRange(2, getColIndexByName("Uploading Files", tickets), tickets.getLastRow()-1);
  
  //Gets the dimensions for the column
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  //Iterates through the collumn and changes the file IDs into full links
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var currentValue = "" + range.getCell(i,j).getValue();
      Logger.log(typeof currentValue.index);
      
      if(currentValue != "" && currentValue.indexOf("sjsu") == -1){
        var withString = "https://drive.google.com/a/sjsu.edu/file/d/" + currentValue + "/";
        range.getCell(i,j).setValue(withString);
      }
    }
  }
}

// Triggered when user submits form
function sendFormByEmail(e){
  var email = "ha-marketing@sjsu.edu";
  var emailLucy = "lucy.yamakawacox@sjsu.edu";
  
  var txt = "a form has been submited, you can view this at https://docs.google.com/spreadsheets/d/1TpqoaHdMfifehgQ9BXxpLYdZNOdXohJ_dWCDgEhoplk/edit#gid=574526564\n\n";
  var id = "";
  var field = "Focus";
  
  if(e.namedValues[field] == "Web Design"){
    id = "#web";
  }else if(e.namedValues[field] == "Graphic Design"){
    id = "#graphic";
  }else if(e.namedValues[field] == "Video and Photo"){
    id = "#media";
  }else if(e.namedValues[field] == "Editor"){
    id = "#edit";
  }else if(e.namedValues[field] == "Administration"){
    MailApp.sendEmail(emailLucy, e.namedValues[field], txt, {name:"Help Desk", replyTo: email});
  }
  
  txt += id;   
  MailApp.sendEmail(email, e.namedValues[field], txt, {name:"Help Desk", replyTo: email});
}