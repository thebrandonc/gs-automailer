/** ------------------------------------------------------------------------------
/** MAKE SURE THE TAB NAMES BELOW MATCH YOUR SPREADSHEET
/** ------------------------------------------------------------------------------*/

const contactTab = 'Sheet1';  // tab name containing contact list and email variables
const emailBodyTab = 'Sheet2';  // tab name containing email body template

/** ------------------------------------------------------------------------------
/** BEWARE: EDITING BELOW THIS LINE MAY BREAK THE SCRIPT
/** ------------------------------------------------------------------------------*/

function getDate(date='') {
  try {
    const today = new Date();
    const curDate = !date ? today : date;
    return Utilities.formatDate(curDate, 'GMT', 'MM/dd/yyyy');
  } catch(err) {
    const msg = {
      type: 'Error',
      msg: `Invalid date format. Recieved ${date}. Expected format: "7/14/25"`
      };
    notify(msg);
  };
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
    try {
      if (typeof recipient[1] === 'object') {  
        const sendDate = getDate(recipient[1]);
        if (sendDate === today) {
          recipient.push(i+2);
          recipients.push(recipient);
        };
      };
    } catch (err) {
      const msg = {
        type: 'Error',
        msg: `Error occured on: ${recipient}. ${err}`
      };
      notify(msg);
    };
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

function sendEmails() {
  const recipients = getRecipients();
  const emails = composeEmails(recipients);
  
  if (recipients) {
    try {
      recipients.forEach((recipient, i) => {
        MailApp.sendEmail(recipient[0], recipient[2], emails[i]);
        insertConfirmation(recipient[recipient.length-1]);
      });
      const msg = {
        type: 'Success!',
        msg: `${recipients.length} email(s) sent!`
        };
      notify(msg);
    } catch (err) {
      const msg = {
        type: 'Error',
        msg: err
      };
      notify(msg);
    };
  };
  if (recipients.length === 0) {
    const msg = {
      type: 'Complete',
      msg: 'No emails were sent.'
      };
    notify(msg);
  };
};

function insertConfirmation(row) {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = doc.getSheetByName(contactTab);
  const cellLocation = `B${row}`;
  const cell = sheet.getRange(cellLocation);
  const confirmation = `Sent on: ${new Date()}`;
  cell.setValue(confirmation);
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Auto Mailer')
      .addItem('📬 Send Emails', 'sendEmails')
      .addToUi();
};

function notify(event) {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
     event.type,
     event.msg,
      ui.ButtonSet.OK);
};
