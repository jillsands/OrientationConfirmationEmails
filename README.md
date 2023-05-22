# Orientation Confirmation Emails
*Owner: Jillian Sands (jesands@umd.edu)*

This program is intended to send Orientation Confirmation emails for the University of Maryland Office of Student Orientation & Transition. 

### Motivation & Context
- Student employees at the UMD Orientation Office have been manually sending confirmation emails to students registered over phone/email for several years. This task takes away from the time we can spend helping students. Also, the magnitude of emails needing to be sent prevents us from personalizing these emails, which leads to repetitive questions from students.
- When a student is registered over phone/email for an orientation program, the employee on duty fills out a “Virtual Blue Card”, which is a Google Form that details the registration. I have attached a Google Apps Script to the Google Sheet of form responses.

### Setup
1. Log into your UMD Gmail account (the one ending in @umd.edu, not your TerpMail)
2. Settings > Accounts > Send mail as > Add another email address. See [this video](https://youtu.be/-m6SBTzx5n0?t=27) for a step by step.
3. Enter "UMD Orientation" as the name and "[askorientation@umd.edu](mailto:askorientation@umd.edu)" as the email. 
4. Be sure to verify yourself using the email sent to [askorientation@umd.edu](mailto:askorientation@umd.edu). If you don’t have access to askorientation@umd.edu, you probably shouldn’t be running the script.

### Usage
1. After following the set up instructions, open the “Summer 2023 Virtual Blue Card” form responses Google Sheet.  
2. Select all the rows you would like to send emails to, then click the "Emails" > "Send Confirmation Emails" from the menu. Wait until the "Running script" message disappears before clicking out of the rows.
3. If you get the message, "A script attached to this document needs your permission to run," please go ahead and authorize it. If you still get a secondary safety message you'll want to click the gray "Advanced" button and allow permission anyway.
4. The code will check the "Email Sent" checkboxes, turn the row’s background yellow, and send an individual email to each student based on the Blue Card info. The emails will appear in your "Sent" box, but to students it'll appear as if it was sent by [askorientation@umd.edu](mailto:askorientation@umd.edu).

![ezgif com-video-to-gif (2)](https://github.com/jillsands/OrientationConfirmationEmails/assets/67645854/473b5f8b-85a9-4ff8-bf14-03347374224d)

### Template
- The template I’m using is: [Confirmation Email Template (PDF)](https://github.com/jillsands/OrientationConfirmationEmails/files/11535964/Gmail.-.UMD.Orientation.Confirmation._.Office.of.Student.Orientation.Transition.pdf).
- Please feel free to email me with any changes.

### Future Semesters
I'll be around to help with this, but just in case:
- The easiest way is to keep using the same Google Form (& Sheet), and just switch out the dates each semester.
- If you want a new form, the code can be found in the Summer 2023 Virtual Blue Card Google Sheet under Extensions > App Script. You can use this code in other Blue Card sheets by going into the new Blue Card Google Sheet and clicking Extensions > App Script, then pasting the code in the editor. Make sure you hit save! *The sheet itself must be named "Form Responses" for the code to work.*

### Maintenance
- Please email me with a screenshot of any errors and I will work on it as soon as possible.
- The code relies heavily on the column structure of the Google Sheet/Form, so please let me know before changing anything.
- I promise not to do anything sketchy with your Google Account :).
