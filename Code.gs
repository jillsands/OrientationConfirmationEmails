// TODO: update parking pass on site

/* Making instructions button for OOAs */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Emails')
      .addItem('Instructions and Setup', 'openReadMe')
      .addToUi();
}

/* Opens link to GitHub ReadMe */
function openReadMe() {
   var htmlOutput = HtmlService.createHtmlOutputFromFile('openURL').setHeight(40);
   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening instructions...');
}

/* Parse blue card for necessary data */ 
function parseData(submission) {

    /* Calculating cost of student's orientation */
    if (submission["Program Type"] == "FO" || submission["Program Type"] == "TO" ) {
      submission.charge = 95
    } else if (submission["Program Type"] == "F1" || submission["Program Type"] == "T1" ) {
      submission.charge = 130
    } else if (submission["Program Type"] == "F2") {
      submission.charge = 197
    }

    /* Calculating cost of family's orientation */
    submission.familyDate = submission["Family Orientation Attendees"] == "None"? "N/A": submission["Family Orientation Date"];
    submission.familyCharge = submission["Family Orientation Attendees"] == "None"? 0 : parseInt(submission["Family Orientation Attendees"]) * 74; 
  
    return submission;

}

 /* Sends an email based on form submissions*/
function sendEmails(e) {

  /* Get template for email */
  var htmlTemplate = HtmlService.createTemplateFromFile("template.html");
  htmlTemplate.student = parseData(e.namedValues);
  var htmlBody = htmlTemplate.evaluate().getContent();

  /* Send email */
  GmailApp.sendEmail(
    e.namedValues["Email"], // recipient
    `UMD Orientation Confirmation | Office of Student Orientation & Transition`, 
    htmlBody,
    {
      from: 'askorientation@umd.edu', // send from askorientation
      name: 'UMD Orientation', 
      htmlBody: htmlBody
  });

  /* Check "Email Sent" box */
  const sheet = SpreadsheetApp.getActive().getSheetByName("Form Responses");
  const emailBox = sheet.getRange(e.range.rowStart, 1);
  emailBox.setValue(true);
} 
