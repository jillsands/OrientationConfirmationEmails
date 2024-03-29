# Orientation Confirmation Emails
*Owner: Jillian Sands (jesands@umd.edu)*

This program is intended to send Orientation Confirmation emails for the University of Maryland Office of Student Orientation & Transition. 

### Motivation & Context
- Student employees at the UMD Orientation Office have been manually sending confirmation emails to students registered over phone/email for several years. This task takes away from the time they can spend helping students. Also, the magnitude of emails needing to be sent prevents us from personalizing these emails, which leads to repetitive questions from students.
- When a student is registered over phone/email for an orientation program, the employee on duty fills out a “Virtual Blue Card”, which is a Google Form that details the registration. I have attached a Google Apps Script to run whenever this form is submitted.

### Usage
The code runs automatically whenever a new Blue Card form is submitted. It will send an email to the student specified using [this template](https://docs.google.com/document/d/1h_KszvItnxJcfPQK9sCOqld164J-nf3TtHLFnq0nNwI/edit?usp=sharing), with dates, charges, and links dynamically generated based on the information that was entered into the Blue Card. If successful, the "Email Sent" checkbox in the Responses sheet will be checked. 

### Setup
 1. Create a new Google Form with the dates of the upcoming Orientation programs
	 - Add a rule to the email field to check for an "@"
	 - All other fields, except the Family Orientation Date, should be required.
 2. Go to Responses > Link to Sheet
 3. Add a column to the left of column A and fill it with empty checkboxes.
 4. At the top of the Sheet, go to Extensions > Apps Script
 5. Copy and paste the 3 files from above into the Apps Script editor. You can copy Code.gs into the default file, but you will need to make new HTML files for the other two.
 6. Run the project using the "Run" button at the top of Code.gs
 7. From the toolbar on the left, click the Triggers tab. Add a trigger with the following details:
	- Function to run: sendEmails 
	 - Event type: On Form Submit
  	- Everything else can be set to the default
 8. If you have not already done so, you must allow Gmail to send emails on your behalf. To do so, log into your UMD Gmail account > Settings > Accounts > Send mail as > Add another email address. Please enter "UMD Orientation" as the name and "askorientation@umd.edu" as the email. Be sure to verify yourself using the email sent to askorientation@umd.edu.

Note: the script will always be run under the account of the person who created the trigger. This means the emails will appear in the creator's "Sent" box, but to students it'll appear as if it was sent by [askorientation@umd.edu](mailto:askorientation@umd.edu). If you get the message "A script attached to this document needs your permission to run" while running the code, please go ahead and authorize it. If you still get a secondary safety message you'll want to click the gray "Advanced" button and allow permission anyway. 

### Maintenance
- Please email me with a screenshot of any errors and I will work on it as soon as possible.
- The code relies heavily on the column structure of the Google Sheet/Form, so please let me know before changing anything.
