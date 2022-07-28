<!-- DESCRIPTION -->
## Simple App to Schedule & Send Bulk Emails from Google Sheets

This repo contains a script that let's you schedule and send bulk emails from Google Sheets. Each email is generated based on a template, which can be tailored to individual recipients through the use of variables.

## Getting Started

To use this script, create a new Google Sheet with two tabs. The first tab should contain your contact list, and the first `3` columns need to adhere to the following format: 

| email          | send_date     | subject     |
|:---------------|:--------------|:------------|
|name@email.com  | 7/28/23       | Your Receipt|

You may add any number of variables desired to the right of the "subject" column:

| email          | send_date     | subject     | var_1       | var_2           |
|:---------------|:--------------|:------------|:------------|:----------------|
|name@email.com  | 7/28/23       | Your Receipt| John        | order confirmed |

On the second tab, insert your email template into cell A1. Include curly braces `{}` wherever you would like to insert a variable into the email. Variables will be inserted in order, from left to right (starting at `var_1` column above). It is important that the number of variable columns matches the number of curly braces in your email template.

Once your spreadsheet is set up, open the Apps Script editor and paste in the script from `autoMailer.js`. Make sure the tab names at the top of the script match the names in your spreadsheet. 

Next, insert at least one recipient (with variables) into your spreadsheet for testing -- preferably using your own email address. Try testing the script with your dummy data inserted by clicking 'Run' in the editor. You will be prompted to enable some permissions so the script can access your spreadsheet, and send emails. Once the permissions are allowed, and it runs successfully, you can start populating your spreadsheet with real data. 

To run the script, either set up set up a trigger to run daily, or select the "Send Emails" option from the "Auto Mailer" dropdown menu in your spreadsheet.

To learn more about how to use this script, watch the Youtube tutorial here: [Check out the Tutorial](https://youtu.be/R7T20_nlvbk)

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.md` for more information.