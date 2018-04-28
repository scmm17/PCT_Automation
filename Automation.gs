// Trinity Cathedral, Portland. Pastoral Care Team prayer email automation.

// To Do:
// Better sorting by family name.
// Add menu support for generating emails, and resetting the "last sent" name.
// Set up a time trigger


var emailers = [
  {
    name: "Scott",
    email: "scmm17@gmail.com"
  },
  {
    name: "Carter",
    email: "scmm17@gmail.com"
  }
];

var familiesPerEmail = 6;
var numberOfColumns = 7;

var familyNameIndex = 0;
var fullNameIndex = 1;
var linkIndex = 4;
var phoneNumberIndex = 5;
var statusIndex = 6;

var preamble = '<br><br>Here is the list of families to contact for this week\'s Pastoral Care Team email notifications. Click on "Email Whole Family" to create an email with creating to the entire family, or single individual, and paste in the body of the email. Use the phone number to call if no email link is present<br><br>';

var sheet;

function getSheet()
{
  if (sheet) {
    return sheet;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getSheetByName('Families');  
  return sheet;
}

function getLastFamilyName() {
  var service = PropertiesService;
  var props = service.getDocumentProperties();
  var name = props.getProperty('LastFamilyName');
  if (name != undefined) {
    return name;
  }

  return "";
}

function setLastFamilyName(name) {
  var service = PropertiesService;
  var props = service.getDocumentProperties();
  var name = props.setProperty('LastFamilyName', name);  
}

function resetLastFamily()
{
  //setLastFamilyName('Richard & Gale Morrison');
  setLastFamilyName('');
}

var data;

function getSpreadsheetData() {
  if (data) {
    return data;
  }

  var sheet = getSheet();
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  data = sheet.getSheetValues(2, 1, lastRow-1, lastColumn);
  return data;  
}

var lastFamilyIndex;

function setLastFamilyIndex(index) {
  lastFamilyIndex = index;
}

function getLastFamilyIndex()
{
  if (lastFamilyIndex) {
    return lastFamilyIndex;
  }
  var lastFamily = getLastFamilyName();
  if (lastFamily === "") {
    lastFamilyIndex = 0;
    return 0;
  }
  var data = getSpreadsheetData();
  for(var i = 0; i < data.length; i++) {
    if (data[i][familyNameIndex] === lastFamily) {
      for(var j = i+1; j < data.length; j++) {
          if (data[j][familyNameIndex] !== '') {
            lastFamilyIndex = j;
            return j;
          }
      }
      lastFamilyIndex = 0;
      return 0;
    }
  }
  Logger.log('Could not find last family');
  lastFamilyIndex = 0;
  return 0; // Yikes, this shouldn't happen
}

function getNextFamilyIndexAndSize()
{
  var index = getLastFamilyIndex();
  var data = getSpreadsheetData();
  var size = 1;
  for(var i = index+1; i < data.length; i++) {
    if (data[i][familyNameIndex] !== '') {
      return {size: size, nextIndex: i}
    }
    size++;
  }
  return {size: size, nextIndex: 0};
}

function getNextFamilyRange() {
  var index = getLastFamilyIndex();
  var sizeAndNextIndex = getNextFamilyIndexAndSize();
  var sheet = getSheet();
  var range = sheet.getRange(index+2, 1, sizeAndNextIndex.size, numberOfColumns);
  setLastFamilyIndex(sizeAndNextIndex.nextIndex);
  return range;
}

function getPhoneNumber(index, size) {
  var data = getSpreadsheetData();
  for(var i = index; i < index + size; i++) {
    if (data[i][phoneNumberIndex] !== "") {
      return data[i][phoneNumberIndex];
    }
  }
  return '';
}

function addFamilyToEmail(body, familyRange) {
  var index = familyRange.getRow() - 2;
  var data = getSpreadsheetData();
  var name = data[index][familyNameIndex];
  // Browser.msgBox('Family name: ' + name + ' size: ' + familyRange.getNumRows());
  setLastFamilyName(name);
  // Add family to email body
  body += 'Family: ' + '<span style="font-weight:bold">' + name + ' </span>';
  var sheet = getSheet();
  var range = sheet.getRange(index + 2, linkIndex+1, 1, 1);
  
  var link = range.getFormula();
  var url = link.match(/=hyperlink\("([^"]+)"/i);
  if (url) {
    body += '<br><a href="' + url[1] + '"> Email Whole Family ' + '</a>' ;
  }
  
  var size = familyRange.getNumRows() - 1
  var processingMembers = true;
  body += '<br>Members: ';
  var firstChildIndex = 0;
  for(var i = index+1; i <= index + size; i++) {
    var name = data[i][fullNameIndex];
    var isMember = data[i][statusIndex] === 'Member' || data[i][statusIndex] === "Designated";
    if (!isMember) {
      firstChildIndex = i;
      break;
    }
  }
  
  for(var i = index+1; i <= index + size; i++) {
    var name = data[i][fullNameIndex];
    var isMember = data[i][statusIndex] === 'Member' || data[i][statusIndex] === "Designated";
    if (processingMembers && !isMember) {
      processingMembers = false;
      body += '<br>Children: ' + name;
      if (i !== index + size)
        body += ', ';
    } else if (i !== index + size) {
        if (i === firstChildIndex - 1)
          body += name;
        else
          body += name + ', ';
    } else { 
      body += name;
    }
  }
  body += '<br>';
  
  var phone = getPhoneNumber(index, familyRange.getNumRows());
    
  if (phone !== '') {
    body += 'Phone: ' + phone + '<br><br>';
  }
  return body;
}

function sendEmail(emailer, body) {
  body += 'Sincerely yours,<br><br>--Scott';
  GmailApp.createDraft(emailer.email, 'Weekly Pastoral Care Team emails', "", {htmlBody: body});
}

function createEmail(emailer, familyRanges) {

  var body = 'Dear ' + emailer.name + ',' + preamble;
  for(var i = 0; i < familyRanges.length; i++) {
    body = addFamilyToEmail(body, familyRanges[i]);
  }
  sendEmail(emailer, body);
}


function sendWeeklyEmails() {
  for(var i = 0; i < emailers.length; i++) {
    var emailer = emailers[i];
    var familyRanges = [];
    for(var j = 0; j < familiesPerEmail; j++) {
      familyRanges.push(getNextFamilyRange());
    }      
    createEmail(emailer, familyRanges);

  }
}
