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

 /* Sends an email for each row in the "Emails" sheet */
function sendEmails() {
  /* Getting grid array of selected rows */
  const sheet = SpreadsheetApp.getActive().getSheetByName("Form Responses");
  const rangeSelected = sheet.getActiveRange();

  /* Send an alert if the selection is invalid */
  if (rangeSelected.getColumn() != 1 || rangeSelected.getLastColumn() != 19) {
    SpreadsheetApp.getUi().alert("Invalid Selection", "Please be sure you highlighted an entire row, from column A to S.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  /* First part of email template */
  const genericTemplate = `<p><span style="font-family: Arial, Helvetica, sans-serif; font-size: 13px;">Welcome to the University of Maryland! This e-mail confirms your registration for a New Student Orientation program.</span></p>
    <p style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);">It is possible to change your New Student Orientation program using our online system up to seven days before your scheduled date. If you wish to cancel your program altogether, you must call the Office of Student Orientation and Transition at (301) 314-8217.</p>
    <p style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);"><strong>You must complete the REQUIRED To-Do List that is applicable to you prior to attending your New Student Orientation program:</strong></p>
    <ul type="disc" style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);">
        <li style="margin: 0px 0px 0px 15px;"><a href="https://orientation.umd.edu/first-year-orientation/first-year-do-list" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://orientation.umd.edu/first-year-orientation/first-year-do-list&source=gmail&ust=1684430338399000&usg=AOvVaw1AvriIiHbQ1rVA-_Bg4z4T" style="color: rgb(17, 85, 204);"><strong>First Year To-Do List</strong></a></li>
        <li style="margin: 0px 0px 0px 15px;"><a href="https://orientation.umd.edu/freshmen-connection/freshmen-connection-do-list" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://orientation.umd.edu/freshmen-connection/freshmen-connection-do-list&source=gmail&ust=1684430338399000&usg=AOvVaw238EZ605UeVVGpwIlu352S" style="color: rgb(17, 85, 204);"><strong>Freshmen Connection To-Do List</strong></a></li>
        <li style="margin: 0px 0px 0px 15px;"><a href="https://orientation.umd.edu/transfer-orientation/transfer-do-list" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://orientation.umd.edu/transfer-orientation/transfer-do-list&source=gmail&ust=1684430338399000&usg=AOvVaw3xaD0IL1R3nbzCChlKpE9P" style="color: rgb(17, 85, 204);"><strong>Transfer To-Do List</strong></a></li>
    </ul>
    <p style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);">Please note the following important reminders so you can have the best experience:</p>
    <p style="margin: 0px;color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);"><span style="font-family: Symbol;">&middot;</span>&nbsp; <strong>Check-In:</strong> We offer 2-day, 1-day, and online programs. Please see the Check-In information for EACH program below:</p>
    <ul type="disc" style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);">
        <li style="margin: 0px 0px 0px 15px;">2-day programs will begin at 8:00AM with Check-In in the Samuel Riggs IV Alumni Center<ul type="circle">
                <li style="margin: 0px 0px 0px 15px;">The 2-day programs will end on the second day at 3:30PM.</li>
            </ul>
        </li>
        <li style="margin: 0px 0px 0px 15px;">1-day programs will begin at 8:00AM with Check-In in the Cole Student Activities Building.<ul type="circle">
                <li style="margin: 0px 0px 0px 15px;">The 1-day programs should end before 4:30PM.</li>
            </ul>
        </li>
        <li style="margin: 0px 0px 0px 15px;">Online programs will begin at 9:00AM and will end around 4:00PM. You will receive a zoom link closer to the day of your New Student Orientation Session.</li>
    </ul>
    <p style="margin: 0px;color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-align: start;text-indent: 0px;background-color: rgb(255, 255, 255);"><span style="font-family: Symbol;">&middot;</span>&nbsp; <strong>Parking:&nbsp;</strong>Please park in Lot 1 for all orientation programs. You can find the Parking Pass linked in your To-Do List linked above.</p>
    <p style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-indent: 0px;background-color: rgb(255, 255, 255);text-align: center;"><strong><span style="color: red;">The Office of Student Orientation and Transition will be holding Terp Family Orientation programs. There is a \$74.00 fee charged to the student per registered family member.</span></strong></p>
    <p style="color: rgb(34, 34, 34);font-family: Arial, Helvetica, sans-serif;font-size: small;font-style: normal;font-weight: 400;text-indent: 0px;background-color: rgb(255, 255, 255);text-align: center;"><strong><span style="color: red;">Are your family members(s) or supporters registered for Terp Family Orientation? If not, they can register by logging into the online registration portal to register for an in person session.&nbsp;</span></strong><strong><span style="color: red;">If they are interested in attending an online Terp Family Orientation, they can register by calling our office at 301-314-8217</span></strong></p><br>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><strong>Student Orientation&nbsp;</strong></p>`

  /* Parsing data and sending emails */
  rangeSelected.getValues().forEach(function (row) {

    const studentDate = row[9];

    /* Calculating cost of student's orientation */
    const programType = row[8]; 
    var studentCharge; 
    if (programType == "FO" || programType == "TO" ) {
      studentCharge = 95
    } else if (programType == "F1" || programType == "T1" ) {
      studentCharge = 130
    } else if (programType == "F2") {
      studentCharge = 197
    } else {
      SpreadsheetApp.getUi().alert(`Invalid Program Type "${programType}"`, "The code does not know how much this program costs. Please email Jillian!", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
    }

    /* Calculating cost of parent's orientation */
    const parents = row[10];
    const parentDate = parents == "None"? "N/A": row[11];
    const parentCharge = parents == "None"? 0 : parseInt(parents) * 74; 

    /* Personalized part of the email template */
    const template = genericTemplate + `<p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Student Orientation Date: ${studentDate}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><strong>Terp Family Orientation</strong></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Participants: ${parents}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Terp Family Orientation Date: ${parentDate}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><strong>Charges</strong></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Student Orientation: \$${studentCharge}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Terp Family Orientation: \$${parentCharge}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Total Charges: \$${studentCharge + parentCharge}</p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><em>*You will be billed AFTER attending New Student Orientation</em></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br></p>
    <p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">Thank you,<br>
    Office of Student Orientation and Transition<br>
    University of Maryland, College Park<br>
    1102 Cole Field House | College Park, MD 20742<br>
    <a href="mailto:askorientation@umd.edu" target="_blank" style="color: rgb(17, 85, 204);">askorientation@umd.edu</a> | 301-314-8217<br>
    Hours: 8:30 am - 4:30 pm EST, Monday - Friday<br>
    <a href="https://orientation.umd.edu/" target="_blank" data-saferedirecturl="https://www.google.com/url?q=https://orientation.umd.edu&source=gmail&ust=1684430804993000&usg=AOvVaw37-yyXekQOwsH00AUxUv4b" style="color: rgb(17, 85, 204);">https://orientation.umd.edu</a></p>` 

    const htmlTemplate = HtmlService.createTemplate(template);
    const htmlBody = htmlTemplate.evaluate().getContent();

    /* Check alias of OOA logged in */
    const me = Session.getActiveUser().getEmail();
    const aliases = GmailApp.getAliases();
    if (!aliases.includes("askorientation@umd.edu")) {
      SpreadsheetApp.getUi().alert("Your account is not set up properly", "Log into your UMD Gmail account > Settings > Accounts > Send mail as > Add another email address. Enter \"UMD Orientation\" as the name and \"askorientation@umd.edu\" as the email. Be sure to verify yourself using the email sent to askorientation@umd.edu.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

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
