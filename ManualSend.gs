/* If something went wrong with a Blue Card that has already been submitted, this file can be used by highlighting the
corresponding row in the Google Form, then running "sendEmails" from the App Script UI. */

/* Making button in Google Sheets for OOAs */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Emails')
      .addItem('Instructions and Setup', 'openReadMe')
      .addItem('Send Confirmation Emails', 'sendEmails')
      .addToUi();
}

/* Opens link to GitHub ReadMe */
function openReadMe() {
   var htmlOutput = HtmlService.createHtmlOutputFromFile('openURL').setHeight(40);
   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening instructions...');
}

/* Parse blue card for necessary data */ 
function parseData(row) {
    const student = [];

    student.college = row[7]
    student["Program Date"] = row[9];

    /* Calculating cost of student's orientation */
    student.programType = row[8]; 
    if (student.programType == "FO" || student.programType == "TO" ) {
      student.charge = 95
    } else if (student.programType == "F1" || student.programType == "T1" ) {
      student.charge = 150
    } else if (student.programType == "F2") {
      student.charge = 227
    } else {
      SpreadsheetApp.getUi().alert(`Invalid Program Type "${student.programType}"`, "The code does not know how much this program costs. Please email Jillian!", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
    }

    /* Calculating cost of parent's orientation */
    student["Family Orientation Attendees"] = row[10];
    student.familyDate = student["Family Orientation Attendees"] == "None"? "N/A": row[11];
    student.familyCharge = student["Family Orientation Attendees"] == "None"? 0 : parseInt(student["Family Orientation Attendees"]) * 85; 
  
    return student;

}

 /* Sends an email for each row in the "Emails" sheet */
function sendEmails() {
  /* Getting grid array of selected rows */
  const sheet = SpreadsheetApp.getActive().getSheetByName("Form Responses");
  const rangeSelected = sheet.getActiveRange();

  /* Send an alert if the selection is invalid */
  if (rangeSelected.getColumn() != 1 || rangeSelected.getLastColumn() < 19) {
    SpreadsheetApp.getUi().alert("Invalid Selection", "Please be sure you highlighted an entire row, from column A to S.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  /* Parsing data and sending emails */
  rangeSelected.getValues().forEach(function (row) {

    var htmlTemplate = HtmlService.createTemplateFromFile("template.html");
    htmlTemplate.student = parseData(row);
    var htmlBody = htmlTemplate.evaluate().getContent();

    /* Check alias of OOA logged in */
    const me = Session.getActiveUser().getEmail();
    const aliases = GmailApp.getAliases();
    /* if (!aliases.includes("askorientation@umd.edu")) {
      SpreadsheetApp.getUi().alert("Your account is not set up properly", "Log into your UMD Gmail account > Settings > Accounts > Send mail as > Add another email address. Enter \"UMD Orientation\" as the name and \"askorientation@umd.edu\" as the email. Be sure to verify yourself using the email sent to askorientation@umd.edu.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    } */

    /* Send email */
    GmailApp.sendEmail(
      row[5], // recipient
      `UMD Orientation Confirmation | Office of Student Orientation & Transition`, 
      htmlBody,
      {
        from: 'askorientation@umd.edu', // send from askorientation
        name: 'UMD Orientation', 
        htmlBody: htmlBody
      });
  });

  /* Check "Email Sent" box and changing coloring */
  rangeSelected.setBackground("yellow")
  const emailBoxes = sheet.getRange(rangeSelected.getRow(), 1, rangeSelected.getNumRows());
  emailBoxes.setValue(true);
} 

