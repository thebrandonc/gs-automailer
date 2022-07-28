/** ------------------------------------------------------------------------------
/** MAKE SURE THE TAB NAMES BELOW MATCH YOUR SPREADSHEET
/** ------------------------------------------------------------------------------*/

const contactTab = 'Sheet1';  // tab name containing contact list and variables
const emailBodyTab = 'Sheet2';  // tab name containing email body

/** ------------------------------------------------------------------------------
/** BEWARE: EDITING BELOW THIS LINE MAY BREAK THE SCRIPT
/** ------------------------------------------------------------------------------*/

function getDate(date='') {
  const today = new Date();
  const newDate = !date ? today : date;
  return Utilities.formatDate(newDate, 'GMT', 'MM/dd/yyyy');
};

function getSheetData(sheetName) {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = doc.getSheetByName(sheetName);
  return sheet.getDataRange().getValues();
};

function getRecipients() {
  const today = getDate();
  const contacts = getSheetData(contactTab);
  contacts.shift();

  const recipients = [];
  contacts.forEach((recipient, i) => {
    const sendDate = getDate(recipient[1]);
    if (sendDate === today) {
      recipient.push(i+2);
      recipients.push(recipient);
    }
  });
  return recipients;
};

function composeEmails(recipients) {
  const email = getSheetData(emailBodyTab)[0][0];
  const emails = [];
  recipients.forEach(recipient => {
    let tempEmail = email.split(" ");
    let placeholders = 0;
    tempEmail.forEach((word, w) => {
      if (word.includes("{}")) placeholders += 1;
      const variable = word.replace('{}', recipient[placeholders+2]);
      tempEmail[w] = variable;
    });
    const formattedEmail = tempEmail.join(' ');
    emails.push(formattedEmail);
  });
  return emails;
};

function insertConfirmation(row) {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = doc.getSheetByName(contactTab);
  const cellLocation = `B${row}`;
  const cell = sheet.getRange(cellLocation);
  const confirmation = `Email sent on: ${getDate()}`;
  cell.setValue(confirmation);
};

function sendEmails() {
  const recipients = getRecipients();
  const emails = composeEmails(recipients);
  
  recipients.forEach((recipient, i) => {
    MailApp.sendEmail(recipient[0], recipient[2], emails[i]);
    insertConfirmation(recipient[recipient.length-1]);
  });
};
