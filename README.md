 var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner")
  var save = activeSheet.getRange("A1:F10").getValues(); 



//.getRange selects the cells you want
//.setValue allows you to write inside the cell you selected using .range'
//.setbackgrond fills the color of the cell you select
//.setFontColor changes the color of the text
//.setFontSize changes the size of the text
//.setFontFamily changes the family of the tex from arial to the font family written
//.setFontWeight makes the text either bolded or unbolded
function default1(){
  var app = SpreadsheetApp; 
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner");
  activeSheet.getRange("A1:F10").clear()
  activeSheet.getRange("A1").setValue("Subjects:")
  activeSheet.getRange("B1").setValue("Monday (mm/dd/yyyy)")
  activeSheet.getRange("C1").setValue("Tuesday (mm/dd/yyyy)")
  activeSheet.getRange("D1").setValue("Wednesday (mm/dd/yyyy)")
  activeSheet.getRange("E1").setValue("Thursday (mm/dd/yyyy)")
  activeSheet.getRange("F1").setValue("Friday (mm/dd/yyyy)")
  activeSheet.getRange("A2:A10").setValue("SubjectName")
  activeSheet.getRange("B2:F10").setValue("Assignments")
  
}
function template1(){
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template1 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#227C9D")
  activeSheet.getRange("B1").setBackground("#17C3B2")
  activeSheet.getRange("C1").setBackground("#227C9D")
  activeSheet.getRange("D1").setBackground("#17C3B2")
  activeSheet.getRange("E1").setBackground("#227C9D")
  activeSheet.getRange("F1").setBackground("#17C3B2")
  activeSheet.getRange("A2:F2").setBackground("#FFCB77")
  activeSheet.getRange("A3:F3").setBackground("#FE6D73")
  activeSheet.getRange("A4:F4").setBackground("#FFCB77")
  activeSheet.getRange("A5:F5").setBackground("#FE6D73")
  activeSheet.getRange("A6:F6").setBackground("#FFCB77")
  activeSheet.getRange("A7:F7").setBackground("#FE6D73")
  activeSheet.getRange("A8:F8").setBackground("#FFCB77")
  activeSheet.getRange("A9:F9").setBackground("#FE6D73")
  activeSheet.getRange("A10:F10").setBackground("#FFCB77")
  activeSheet.getRange("A1:F1").setFontSize(14).setFontColor("#FEF9EF").setFontWeight("bold").setFontFamily("Lemon")
  activeSheet.getRange("A2:F10").setFontFamily("Lemon").setFontSize(13)
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template2(){
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template2 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setValue("Subjects:").setBackground("#227C9D")
  activeSheet.getRange("B1").setValue("Monday (mm/dd/yyyy)").setBackground("#FFC09F")
  activeSheet.getRange("C1").setValue("Tuesday (mm/dd/yyyy)").setBackground("#FFEE93")
  activeSheet.getRange("D1").setValue("Wednesday (mm/dd/yyyy)").setBackground("#FCF5C7")
  activeSheet.getRange("E1").setValue("Thursday (mm/dd/yyyy)").setBackground("#A0CED9")
  activeSheet.getRange("F1").setValue("Friday (mm/dd/yyyy)").setBackground("#ADF7B6")
  activeSheet.getRange("A2:A10").setValue("SubjectName")
  activeSheet.getRange("B2:F10").setValue("Assignments")
  activeSheet.getRange("A2:A10").setBackground("#ADF7B6")
  activeSheet.getRange("B2:B10").setBackground("#A0CED9")
  activeSheet.getRange("C2:C10").setBackground("#FCF5C7")
  activeSheet.getRange("D2:D10").setBackground("#FFEE93")
  activeSheet.getRange("E2:E10").setBackground("#FFC09F")
  activeSheet.getRange("F2:F10").setBackground("#227C9D")
  activeSheet.getRange("A1:F1").setFontSize(15).setFontColor("#000000").setFontWeight("bold").setFontFamily("Pacifico")
  activeSheet.getRange("A2:F10").setFontFamily("Pacifico").setFontSize(14).setFontColor("#000000")
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template3(){
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template3 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#FEFCFB")
  activeSheet.getRange("B1").setBackground("#FEFCFB")
  activeSheet.getRange("C1").setBackground("#FEFCFB")
  activeSheet.getRange("D1").setBackground("#FEFCFB")
  activeSheet.getRange("E1").setBackground("#FEFCFB")
  activeSheet.getRange("F1").setBackground("#FEFCFB")
  activeSheet.getRange("A2:A10").setBackground("#1282A2")
  activeSheet.getRange("B2:B10").setBackground("#034078")
  activeSheet.getRange("C2:C10").setBackground("#001F54")
  activeSheet.getRange("D2:D10").setBackground("#034078")
  activeSheet.getRange("E2:E10").setBackground("#001F54")
  activeSheet.getRange("F2:F10").setBackground("#034078")
  activeSheet.getRange("A1:F1").setFontSize(18).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("A2:F10").setFontSize(16).setFontColor("#FEFCFB").setFontWeight("bold")
  activeSheet.getRange("A1:F10").setFontFamily("Akshar")
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template4(){
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template4 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#FBE7C6")
  activeSheet.getRange("B1").setBackground("#B4F8C8")
  activeSheet.getRange("C1").setBackground("#FBE7C6")
  activeSheet.getRange("D1").setBackground("#B4F8C8")
  activeSheet.getRange("E1").setBackground("#FBE7C6")
  activeSheet.getRange("F1").setBackground("#B4F8C8")
  activeSheet.getRange("A2:A10").setBackground("#FFAEBC")
  activeSheet.getRange("B2:B10").setBackground("#A0E7E5")
  activeSheet.getRange("C2:C10").setBackground("#FFAEBC")
  activeSheet.getRange("D2:D10").setBackground("#A0E7E5")
  activeSheet.getRange("E2:E10").setBackground("#FFAEBC")
  activeSheet.getRange("F2:F10").setBackground("#A0E7E5")
  activeSheet.getRange("A1:F1").setFontSize(22).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("A2:F10").setFontSize(20).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("A1:F10").setFontFamily("Square Peg")
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template5(){
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template1 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1:F1").setBackground("#EAC5D8")
  activeSheet.getRange("A2:A10").setBackground("#56CBF9")
  activeSheet.getRange("B2:B10").setBackground("#7FBEEB ")
  activeSheet.getRange("C2:C10").setBackground("#DBD8F0")
  activeSheet.getRange("D2:D10").setBackground("#56CBF9")
  activeSheet.getRange("E2:E10").setBackground("#7FBEEB ")
  activeSheet.getRange("F2:F10").setBackground("#DBD8F0")
  activeSheet.getRange("A1:F1").setFontSize(15).setFontColor("#6e5e66").setFontWeight("bold").setFontFamily("Lemon")
  activeSheet.getRange("A2:F10").setFontFamily("Lemon").setFontColor("#FFFFFF").setFontSize(13)
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template6(){ 
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template1 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#AE4308")
  activeSheet.getRange("B1").setBackground("#BC5F04")
  activeSheet.getRange("C1").setBackground("#BC5F04")
  activeSheet.getRange("D1").setBackground("#BC5F04")
  activeSheet.getRange("E1").setBackground("#BC5F04")
  activeSheet.getRange("F1").setBackground("#BC5F04")
  activeSheet.getRange("A2:A10").setBackground("#873406")
  activeSheet.getRange("B2:F10").setBackground("#874000")
  activeSheet.getRange("A1:F1").setFontSize(16).setFontColor("#010001").setFontWeight("bold")
  activeSheet.getRange("A2:F10").setFontSize(16).setFontColor("#2B0504").setFontWeight("bold")
  activeSheet.getRange("A1:F10").setFontFamily("Varela Round").setFontSize(18)
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template7(){ 
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template1 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#6c8ca4")
  activeSheet.getRange("B1").setBackground("#263739")
  activeSheet.getRange("C1").setBackground("#97a9ac")
  activeSheet.getRange("D1").setBackground("#8b493e")
  activeSheet.getRange("E1").setBackground("#2c4c64")
  activeSheet.getRange("F1").setBackground("#6c8ca4")
  activeSheet.getRange("A2:A10").setBackground("#6c8ca4")
  activeSheet.getRange("B2:B10").setBackground("#263739")
  activeSheet.getRange("C2:C10").setBackground("#97a9ac")
  activeSheet.getRange("D2:D10").setBackground("#8b493e")
  activeSheet.getRange("E2:E10").setBackground("#2c4c64")
  activeSheet.getRange("F2:F10").setBackground("#6c8ca4")
  activeSheet.getRange("A1:F1").setFontSize(20).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("A1:F1").setFontFamily("Orelega One").setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK) 
  activeSheet.getRange("A2:A10").setFontSize(18).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("B1:B10").setFontSize(18).setFontColor("#FFF5EE").setFontWeight("bold")
  activeSheet.getRange("C2:C10").setFontSize(18).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("D1:D10").setFontSize(18).setFontColor("#FFF5EE").setFontWeight("bold")
  activeSheet.getRange("E2:E10").setFontSize(18).setFontColor("#000000").setFontWeight("bold")
  activeSheet.getRange("F1:F10").setFontSize(18).setFontColor("#FFF5EE").setFontWeight("bold")
  activeSheet.getRange("A2:F10").setFontFamily("Orelega One")
  activeSheet.getRange("A1:F10").setValues(save)
}
//changes how template looks by changeing the color, fonts, and size 
function template8(){ 
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project SEP11"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath template1 would only work in the google sheet titled "Homework Planner"
  activeSheet.getRange("A1").setBackground("#E5C5B6")
  activeSheet.getRange("B1").setBackground("#E5C5B6")
  activeSheet.getRange("C1").setBackground("#E5C5B6")
  activeSheet.getRange("D1").setBackground("#E5C5B6")
  activeSheet.getRange("E1").setBackground("#E5C5B6")
  activeSheet.getRange("F1").setBackground("#E5C5B6")
  activeSheet.getRange("A2:F10").setBackground("#878284")//sets the background color for the cells A2-F10
  activeSheet.getRange("A1:F1").setFontSize(16).setFontColor("#2A252C").setFontWeight("bold")//sets the font in cells A1-F1 to be size 16, bold and color so dark it looks like black
  activeSheet.getRange("A2:F10").setFontSize(16).setFontColor("#FFFFFF").setFontWeight("bold")//sets the font in cells A2-F10 to be bold, size 16, and white
  activeSheet.getRange("A1:F10").setFontFamily("Bad Script").setFontSize(14) //makes the font in the cells A1-F10 the size 14 and font will be "Bad Script"
  activeSheet.getRange("A1:F10").setValues(save)
}
//the button that deletes the entire Homework planner 
function clearContent() {
  var app = SpreadsheetApp; //locates the google sheet we are in: "Freedom Project"
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Homework Planner"); //all code underneath clearContent would only work in the google sheet titled "Homework Planner"

  var ui = SpreadsheetApp.getUi();
  var user = ui.alert('Are you sure you want to Delete everything in your planner?', ui.ButtonSet.YES_NO);//when the button is clicked the user will recive a alert message where they either click a button that say yes or a button that says no

  // if the the user clicks the yes button then the Homework planner will be deleted 
  if (user == ui.Button.YES) {
    ui.alert('Make sure to click defualt template again after your press delete everything if you want to add another template.')
  activeSheet.getRange("A1:F10").clear() //.clear deletes everything in the cell you slected, the background color, text, everything
  }
  //if the the user chooses to click the no button then a alert will apper telling the user the nothing has been deleted
  else {
     ui.alert('Do not worry nothing has been deleted.');
  }
}

// function sendEmail(){
//   var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SendEmails");
//   var lr = emailSheet.getLastRow(); //finds the last row that has anything written in it's cells
//    for(var i = 2; i <= lr; i++ ){ //i starts at row 2 where the first email is written, the for loop would loop through all the emails from the help of 'var lr' which gets all the rows that has values
//     var email = emailSheet.getRange(i, 1).getValue(); //gets the email in each row after row 1
//     var subject = emailSheet.getRange(i, 2).getValue(); //gets the subject in each row after row 1
//     var message= emailSheet.getRange(i,3).getValue(); //gets the value in (any row after row 1, column 3)

function sendEmail(){
  var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SendEmails");
    var email = emailSheet.getRange("E10").getValue(); //gets the email from cell E10 and stores in a variable called email 
    var subject = emailSheet.getRange("E14").getValue(); //gets the subject from cell E14 and stores it in a variable called subject
    var message= emailSheet.getRange("E18").getValue(); // gets the message from cell E18 and stores it in a variable called message

    MailApp.sendEmail(email, subject, message); 
}

function eraseEmail(){
  var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SendEmails");
  var rangesToClear = ["E10", "E14", "E18"]; //saves the cells that the user writes in so that they can erase it
  for(var i = 0; i < rangesToClear.length; i++) {//loops through the array 
    emailSheet.getRange(rangesToClear[i]).clearContent();// as i increases, the element in the rangesToClear gets deleted one by one until 'i' is equal to the array length, allowing all the cells the user writes in to be deleted when users click the  'clear' button
  }
}

function addEventsToGoogleCalendar(){ //imports events from Google Sheets to Google Calendar
  var ss= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //gets the current active spreadsheet which is sheet2
  var dataRange = ss.getRange("A2:E15").getValues(); //gets all the values from A2 to E16
  var email = ss.getRange("K3").getValue()
  var cal = CalendarApp.getCalendarById(email); //gets the google calendar where your event can be added by using your gmail
  for(var i = 0; i < dataRange.length; i++){
    cal.createEvent(dataRange[i][0], dataRange[i][1], dataRange[i][2], {location: dataRange[i][3], description: dataRange[i][4]}) //grabs the values underneath Event, Start, End, Location and Description, starting from row 2 and ending on the last row
    Logger.log(dataRange[i]); //lets you see what would be added
  
  }
}
function clear() {
  var form = SpreadsheetApp.getActiveSpreadsheet();
  var formClear = form.getSheetByName("Assignment");

  var rangesToClear = ["D6", "D9", "D11"]; //saves the cells that the user writes in so that they can erase it
  for(var i = 0; i < rangesToClear.length; i++) {//loops through the entire array 
    formClear.getRange(rangesToClear[i]).clearContent();//as i increases, the element in the rangesToClear gets deleted one by one until 'i' is equal to the array length, allowing all the cells the user writes in to be deleted when users click the  'clear' button
  }
}
function saveAssignment() {
  var sav = SpreadsheetApp.getActiveSpreadsheet();
  var formSave = sav.getSheetByName("Assignment");
  var inputs = sav.getSheetByName("custom");
  
  var values = [[formSave.getRange("D6").getValue(),formSave.getRange("D9").getValue(),formSave.getRange("D11").getValue()]]; //gets the values for what was written in the cells "D6", "D9", and "D11" and stores it in a variable called values 

  inputs.getRange(inputs.getLastRow()+1, 1,1, 3).setValues(values);//goes to the google sheet titled custom and in the next row that is free it puts the information that was in the values variable 
  
  var color = [[formSave.getRange("D3").getValue()]]; //gets the value of the was written in the cell "D3" and stores in the variable called color 
  inputs.getRange("A1:C100").setBackground(color)//makes the background color of cells A2 - C100 whatever color the user put the hex code for in the form 
  clear();// calls the function clear: this makes so that after the user presses the button save whatever was written in the cells will then be erased
}

function remove() {
  var del = SpreadsheetApp; 
  var activeSheet = del.getActiveSpreadsheet().getSheetByName("custom"); 

  var ui = SpreadsheetApp.getUi();
  var user = ui.alert('Are you sure you want to Delete everything?', ui.ButtonSet.YES_NO); // a prompt is asked asking the user if they are sure they want delete everything and the user can answer either yes or no 

  if (user == ui.Button.YES) {// if the button yes is clicked 
  activeSheet.getRange("A2:C100").clear() //.clear deletes everything in the cell you slected, the background color, text, everything
  }
  else {
     ui.alert('Do not worry nothing has been deleted.'); //if the user clicks the button No this will make sure that the user is alerted with a message saying that they have nothing worry and nothing has been deleted
  }
}
