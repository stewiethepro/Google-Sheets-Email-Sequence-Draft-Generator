//////////////////////////////////////

// Global variables

//////////////////////////////////////

// Get UI
var ui = SpreadsheetApp.getUi();

// Define merge field labels
var emailMergeFieldsLabels = ['{contactFirstName}', '{contactLastName}', '{contactTitle}', '{companyName}', '{domain}', '{START}', '{END}', '{aeName}', '{aeFirstName}','{aeLastName}', '{aeEmail}', '{aeTitle}', '{aeExt}', '{aeBP}'];

// Define email signature template
var emailSignatureTemplate = '<div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature"> <div dir="ltr"> <div><br></div> <div><font color="#666666">--</font></div> <div><font color="#666666"><br></font></div> <div><font color="#000000" style="font-weight:bold">{aeName} </font><font color="#666666">|&nbsp;</font>{aeTitle} at OnceHub<font color="#666666">&nbsp;</font><span style="color:rgb(102,102,102)">|</span><font color="#666666"><b>&nbsp;</b></font><font color="#000000">+1.650.206.5585 Ext. {aeExt}&nbsp;</font><span style="color:rgb(102,102,102)">|&nbsp;</span><a href="mailto:{aeEmail}" target="_blank">{aeEmail}</a></div><div><a href="{aeBP}" target="_blank">Schedule time with me</a><br></div><div><br></div><div><img src="http://cdn.oncehub.com/static-files/images/email/OH-signature.png"><br></div><div><a href="https://www.oncehub.com/" target="_blank">Website</a>&nbsp;<font color="#666666">|&nbsp;<a href="https://blog.oncehub.com/" target="_blank">Blog</a>&nbsp;|&nbsp;<a href="https://www.linkedin.com/company/oncehub/" target="_blank">Linkedin</a>&nbsp;|&nbsp;<a href="https://www.youtube.com/oncehub" target="_blank">Youtube</a>&nbsp;|&nbsp;<a href="https://twitter.com/OnceHub" target="_blank">Twitter</a>&nbsp;|&nbsp;<a href="https://www.facebook.com/OnceHub" target="_blank">Facebook</a></font><br></div><div><br></div><div><font color="#000000">Watch our latest video:</font>&nbsp;<a href="https://youtu.be/QuUtDqZ7Np8" target="_blank">Intro to OnceHub</a></div> </div> </div>'

//////////////////////////////////////

// Create menu

//////////////////////////////////////


function onOpen() {
  ui.createMenu('Email menu')
      .addItem('Create drafts', 'createDrafts')
      .addToUi();
}


//////////////////////////////////////

// App

//////////////////////////////////////



function createDrafts() {

  // Define date
  var dateNow = new Date().valueOf()

  // Get sheets
  var sheets = getSheets();

  // Define sheetsData object
  var sheetsData = getData(sheets);

  // Define contacts sheet object
  var contacts = sheetsData[0].contacts;

  // Define templates sheet object
  var templates = sheetsData[1].templates;

  // Define ae sheet object
  var ae = sheetsData[2].ae;

  // Define aeData object
  var aeData = assignAeVariables(ae);

  // Define templates object
  var templatesData = assignTemplateVariables(templates);

  // Loop through contact data
  for (var i = 0; i < contacts.data.length; ++i) {
    var row = contacts.data[i];

    // Define contactData object
    var contactData = assignContactVariables(contacts, row);

    // Define stage
    var stage = defineStage(contactData);

    // Get templateContent
    var templateContent = lookupTemplate(contactData, templatesData, stage);

    // Check if stage = email 3, then continue if so
    if (stage == 3) {

      continue

    }

      // Check if stage = email 0, then draft an email
      if (stage == 0) {

        // createDraft
        var finalContent = draftEmail(dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage);

      };

        // Check if stage = email 1 or email 2, then check if it's time to draft,
        // if not, continue.
        if (stage == 1 || stage == 2) {

          var timeToDraft = checkDaysLeft(dateNow, contactData, templateContent);


          if (timeToDraft == true) {

            // find sent email
            var message = findSentEmail(contactData);

            if (message != "Message has not yet been sent") {

              // Draft a reply to the email
              var finalContent = draftReplyAll(message, dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage);

            } else continue

          } else continue

        };

  updateFields(sheets, contacts, finalContent, i);

  }

}

//////////////////////////////////////

// Functions

//////////////////////////////////////


function getSheets() {
// Select sheets
var sheet = SpreadsheetApp.getActiveSpreadsheet(); // Use data from the active sheet

  return [{contacts: sheet.getSheetByName('Contacts')}, {templates: sheet.getSheetByName('Templates')}, {ae: sheet.getSheetByName('AE')}]

}

// Get data in sheet
function getData (sheets) {
  var sheetsData = [];
  for (var i = 0; i < sheets.length; ++i) {
    var sheet = sheets[i];
    var sheetName = Object.keys(sheet);
    var startRow = 2;                                                        // First row of data to process
    var numRows = sheet[sheetName].getLastRow() - 1;                            // Number of rows to process
    var lastColumn = sheet[sheetName].getLastColumn();                          // Last column
    var dataRange = sheet[sheetName].getRange(startRow, 1, numRows, lastColumn) // Fetch the data range of the active sheet
    var data = dataRange.getValues();

    var sheetData = {};
    sheetData[sheetName] =
    {
      'startRow': startRow,
      'numRows': numRows,
      'lastColumn': lastColumn,
      'dataRange': dataRange,
      'data': data
    }

    sheetsData.push(sheetData);

  }

  return sheetsData;

};

function assignAeVariables(ae) {

    // Define aeName
    var aeName = splitName(ae.data[0][0]);

    // Define aeData object
    var aeData = {
      "firstName": aeName.firstName,
      "lastName": aeName.lastName,
      "fullName": aeName.fullName,
      "email": aeName.firstName.toLowerCase() + '.' + aeName.lastName.toLowerCase() + '@' + 'oncehub.com',
      "title": ae.data[0][1],
      "ext": ae.data[0][2],
      "bookingPage": ae.data[0][3],
    }

    return aeData
  }


  function assignTemplateVariables(templates) {

    var sequencesData = [];

    // Loop through each row of templates data
    for (var i = 0; i < templates.data.length; ++i) {

      var row = templates.data[i];
      var sequenceData =
        {
          "sequenceID": row[0],
          "templateID": row[1],
          "daysBeforeSend": row[2],
          "email":
            {
              "subject": row[3],
              "body": row[4]
            }
        }

      sequencesData.push(sequenceData);

    }

    // Group by sequence ID (reduce)
    let templatesData = sequencesData.reduce((r, a) => {
     r[a.sequenceID] = [...r[a.sequenceID] || [], a];
     return r;
    }, {});

    return templatesData

  }

function assignContactVariables(contacts, row) {

      // Define contact name
      var contactName = splitName(row[0]);               // Col A: Contact name

      // Define contact data object
      var contactData = {
        "firstName": contactName.firstName,
        "lastName": contactName.lastName,
        "fullName": contactName.fullName,
        "email":
        {
          "to": row[1],
          "bcc": "",
        },
        "title": row[2],
        "company": row[3],
        "domain": row[4],
        "sequenceID": row[5],
        "lastEmailDate": row[6],
        "emailID": row[7],
        "emailStatus": row[8]
      }

      // Check if we have their email
      // if not then generate emails and define them in the contact data object
      if (contactData.email.to == "") {

        var contactEmails = guessEmails(contactData);
        contactData.email.to = contactEmails.to;
        contactData.email.bcc = contactEmails.bcc;

      }

      return contactData

}

function createEmailMergeFields(contactData, aeData, emailMergeFieldsLabels) {

// Create array of email merge fields
var emailMergeFieldsData = [contactData.firstName, contactData.lastName, contactData.title, contactData.company, contactData.domain, '<div>', '</div> <div><br></div>', aeData.fullName, aeData.firstName, aeData.lastName, aeData.email, aeData.title, aeData.ext, aeData.bookingPage];

return emailMergeFieldsData;

}

function defineStage (contactData) {

var emailStatus = contactData.emailStatus;

  // if not yet contacted
  if (emailStatus == "") {
    var stage = 0;
  }
  // if initial email drafted
  else if (emailStatus == "email 1 drafted") {
    var stage = 1;
  }
  // if follup up 1 drafted
  else if (emailStatus == "email 2 drafted") {
    var stage = 2;
  }
  // if follup up 2 drafted
  else if (emailStatus == "email 3 drafted") {
    var stage = 3;
  }

  return stage

}

// Check last date contacted, only run if stage 1 or 2
function lookupTemplate(contactData, templatesData, stage) {

  var targetSequence = contactData.sequenceID;
  var sequence = templatesData[targetSequence];
  var targetStage = stage + 1;

  for (var j = 0; j < sequence.length; ++j) {

    template = sequence[j];
    templateID = template.templateID;

      if (templateID == "Email " + targetStage) {

       var templateContent =
      {
        "sequenceID": template.sequenceID,
        "templateID": template.templateID,
        "daysBeforeSend": template.daysBeforeSend,
        "email":
        {
          "subject": template.email.subject,
          "body": template.email.body
        }
      }

      return templateContent

      }
  }

};

//
function checkDaysLeft(dateNow, contactData, templateContent) {

  var days = daysSinceLastContact(contactData.emailID, dateNow)
  if (days >= templateContent.daysBeforeSend) {
    return true
  }
  else return false

}

// Only run if stage 1 or 2, takes emailID as an argument
function daysSinceLastContact (startDate, endDate) {

  var startDate = new Date().valueOf();
  var endDate = new Date().valueOf();
  var sec = 1000;
  var min = 60*sec;
  var hour = 60*min;
  var day = 24*hour;
  var diff = endDate-startDate;
  var days = Math.floor(diff/day);

  return days;

}

function splitName(name){

  var fullName = name.split(' ');
  var firstName = fullName[0];
  var lastName = fullName[1];

  return {
    "fullName": name,
    "firstName": firstName,
    "lastName": lastName
  }
};

function guessEmails(contactData) {

  var email1 = contactData.firstName + contactData.lastName + '@' + contactData.domain;
  var email2 = contactData.firstName + '@' + contactData.domain;
  var email3 = contactData.firstName[0] + contactData.lastName + '@' + contactData.domain;
  var email4 = contactData.firstName + '.' + contactData.lastName + '@' + contactData.domain;

  return {
    "to": email1,
    "bcc": email2 + ',' + email3 + ',' + email4
  }

};


function findSentEmail(contactData) {

  var emailID = contactData.emailID;
  var threads = GmailApp.search(emailID);
  var foundMessage = threads[0].getMessages();
  var message = foundMessage[0];
  var isDraft = message.isDraft();
  if (isDraft) {
    message = "Message has not yet been sent";
  }
  return message;
};

function replaceMergeFields(templateToMerge, emailMergeFieldsData, emailMergeFieldsLabels) {

  // Loop through the template and fill in the variables
  for (var k = 0; k < emailMergeFieldsLabels.length; ++k) {

    var regexstring = emailMergeFieldsLabels[k];
    var regexp = new RegExp(regexstring, "g");
    var mergedTemplate = templateToMerge.replace(regexp, emailMergeFieldsData[k] || '');
    templateToMerge = mergedTemplate;

  }
  return mergedTemplate;

};

function createEmailID() {

  // Create hidden uniqueID for email based on milisecond timestamp
  var rawEmailID = Date.now();
  var formattedEmailID = '<p style="display:none">emailID: ' + rawEmailID + '</p>'
  var emailID =
  {
    "rawEmailID": rawEmailID,
    "formattedEmailID": formattedEmailID
  }

  return emailID

};

function generateEmailContent(dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage) {

  // Combine body and signature templates
  var emailSubjectTemplate = templateContent.email.subject;
  var emailBodyTemplate = templateContent.email.body + emailSignatureTemplate;

  // Create email merge fields
  var emailMergeFieldsData = createEmailMergeFields(contactData, aeData, emailMergeFieldsLabels);


  // Merge fields on templates
  var emailSubject = replaceMergeFields(emailSubjectTemplate, emailMergeFieldsData, emailMergeFieldsLabels);
  var emailBody = replaceMergeFields(emailBodyTemplate, emailMergeFieldsData, emailMergeFieldsLabels);

  // Generate email ID
  var emailID = createEmailID();

  // Create formatted date
  var date = new Date(dateNow).toDateString();

  // Define email content object
  var emailContent =
  {
    "emailSubject": emailSubject,
    "emailBody": emailBody,
    "emailID": emailID
  }

  // Define output data to update sheet
  var outputData =
  {
   "date": date,
   "emailID": emailID.rawEmailID,
   "stage": stage
  }

  var finalContent = {"emailContent": emailContent, "outputData": outputData}

  return finalContent;

};

function updateFields(sheets, contacts, finalContent, i) {

  var sheet = sheets[0].contacts;
  var newStage = finalContent.outputData.stage + 1;
  var outputData =
  [[
    finalContent.outputData.date,
    finalContent.outputData.emailID,
    "email " + newStage + " drafted"
  ]];

  sheet.getRange(contacts.startRow + i, 7, 1, 3).setValues(outputData); // Update the last columns with date, emailID and emailStatus
  SpreadsheetApp.flush(); // Make sure the last cell is updated right away

};

function draftEmail(dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage) {

  var finalContent = generateEmailContent(dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage);

  // Draft email
  GmailApp.createDraft(
    contactData.email.to,                     // Recipient
    finalContent.emailContent.emailSubject,                      // Subject
    '',                                // Body (plain text)
    {
      htmlBody: finalContent.emailContent.emailBody + finalContent.emailContent.emailID.formattedEmailID,                // Options: Body (HTML)
      bcc: contactData.email.bcc
    }
  );

  return finalContent

};

function draftReplyAll(message, dateNow, contactData, aeData, templateContent, emailSignatureTemplate, emailMergeFieldsLabels, stage) {

  var finalContent = generateEmailContent(dateNow, contactData, aeData, templateContent, emailSignatureTemplate,emailMergeFieldsLabels, stage);

  // Draft reply
  message.createDraftReply(
  "",
  {
  htmlBody: finalContent.emailContent.emailBody + finalContent.emailContent.emailID.formattedEmailID,
  bcc: contactData.email.bcc
  });

  return finalContent

};
