function makeFamiliesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Members');

  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var data = sheet.getSheetValues(2, 1, lastRow-1, lastColumn);
  
  var fullNameIndex = 0;
  var firstNameIndex = 1;
  var preferredNameIndex = 2;
  var lastNameIndex = 3;
  var emailIndex = 4;
  var statusIndex = 5
  var familyIndex = 6
  var phoneIndex = 7;
  var phone1Index = 8;
  var phone2Index = 9;
  
  for(var row = 0; row < data.length; row++) {
    
    var rowData = data[row];
    var fullName = rowData[fullNameIndex];
    var firstName = rowData[firstNameIndex];
    var preferredName = rowData[preferredNameIndex];
    var lastName = rowData[lastNameIndex];
    var email = rowData[emailIndex];
    var status = rowData[statusIndex];
    var family = rowData[familyIndex];
    var phone = rowData[phoneIndex];
    var phone1 = rowData[phone1Index];
    var phone2 = rowData[phone2Index];
    if (!phone)
      phone = phone1;
    if (!phone)
      phone = phone2;
    var member = {fullName: fullName,
                  firstName: firstName,
                  preferredName: preferredName,
                  lastName: lastName,
                  email: email,
                  status: status,
                  family: family,
                  phone: phone};
    
    addMember(member);
  } 
  
  processFamilies();
}

var familiesTable = {};

function addMember(member) {
  var family = familiesTable[member.family];
  if (!family) {
    family = {members: [member]}
    familiesTable[member.family] = family;
  } else { 
    family.members.push(member);
  }
}

var currentRow = 1;
var maxFamilies = undefined;
var rows = [];
var colors = [];
var weights = [];

function processFamilies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var oldSheet = ss.getSheetByName('Families');
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }
  var newSheet = ss.insertSheet('Families');
  // newSheet.appendRow(["Family Name", "Full Name", "First Name", "Last Name", "Email", "Phone", "Status"]);
  rows.push(["Family Name", "Full Name", "First Name", "Last Name", "Email", "Phone", "Status"]);
  var color = "PeachPuff";
  colors.push([color, color, color, color, color, color, color]);
  weights.push(['bold', 'bold', 'bold', 'bold', 'bold', 'bold', 'bold']);
  var range = newSheet.getRange(currentRow, 1, 1, 20);
  range.setFontSize(14);
  range.setFontWeight('bold');
  newSheet.setFrozenRows(1);
  currentRow++;
  newSheet.setColumnWidth(1, 250);
  newSheet.setColumnWidth(2, 150);
  newSheet.setColumnWidth(3, 120);
  newSheet.setColumnWidth(4, 150);
  newSheet.setColumnWidth(5, 200);
  newSheet.setColumnWidth(6, 100);
  newSheet.setColumnWidth(7, 100);

  var familyCount = 0;
  for(var familyName in familiesTable) { 
    var family = familiesTable[familyName];
    family.familyName = familyName;
    processFamily(family, newSheet);
    familyCount++;
    if (familyCount === maxFamilies)
      break;
  }
  // Logger.log(rows);
  // Add the data to the spreadsheet
  range = newSheet.getRange(1, 1, rows.length, 7);
  range.setValues(rows);
  range.setBackgrounds(colors);
  range.setFontWeights(weights);
}

var rowColors = ["PowderBlue", "SkyBlue"];
var currentColor = 0;

function getColor() { 
  var color = rowColors[currentColor];
  currentColor = (currentColor + 1) % rowColors.length;
  return color;
}

function makeMailLink(addresses, greeting, isFamily) {
  var label = isFamily ? "Email Whole Family" : addresses[0];
  if (addresses.length == 1) {
    return '=HYPERLINK("mailto:' + addresses[0] + '?subject=Trinity%20Cycle%20of%20Prayer&body=' + greeting +'", "' + label + '")';
  }
  var link = '=HYPERLINK("mailto:';
  var numAddresses = addresses.length;
  var lastIndex = numAddresses - 1;
  for(var i = 0; i < numAddresses; i++) {
    link += addresses[i];
    if (i !== lastIndex) {
      link += ',';
    }
  }
  return link + '?subject=Trinity%20Cycle%20of%20Prayer&body=' + greeting +' ", "Email Whole Family")';
}

function getFirstName(member) {
  return member.preferredName ? member.preferredName : member.firstName;
}

function hasMemberOrDesignated(family) {
  for(var i = 0; i < family.members.length; i++ ) {
    if (family.members[i].status === 'Member' || family.members[i].status === 'Designated')
      return true;
  }
  return false;
}

function hasContactInfo(family) {
  for(var i = 0; i < family.members.length; i++ ) {
    if (family.members[i].email !== '' || family.members[i].phone !== '')
      return true;
  }
  return false;  
}

function processFamily(family, familySheet) {
  if (!hasMemberOrDesignated(family)) { // Filter out all child families.
    return;
  }
  if (!hasContactInfo(family)) {
    Logger.log('No contact: ' + family.familyName);
    return;
  }
  var color = getColor();
  createFamilyHeader(family, color, familySheet);
  for(var i = 0; i < family.members.length; i++ ) {
    var member = family.members[i];
    var firstName = getFirstName(member);
    var greeting = createGreeting(family);
    var email = makeMailLink([member.email], greeting, false);
    
    rows.push(["", member.fullName, firstName, member.lastName, email, member.phone, member.status]);
    colors.push([color, color, color, color, color, color, color]);
    weights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal']);

    currentRow++;
  }
}
            
if (!String.prototype.startsWith) {
	String.prototype.startsWith = function(search, pos) {
		return this.substr(!pos || pos < 0 ? 0 : +pos, search.length) === search;
	};
}

function includes(arr, val) {
  for(var i = 0; i < arr.length; i++) {
    if (arr[i] === val)
      return true;
  }
  return false;
}

function sortFamily(family) {
  var members = [];
  
  for(var i = 0; i < family.members.length; i++ ) {
    var member = family.members[i];
    var isNonMember = member.status === "Adult Child - Non Member";
    if (isNonMember)
      continue;
    if (member.status === 'Member' || member.status === 'Designated')
      members.push(member);
  }
  for(var i = 0; i < family.members.length; i++ ) {
    var member = family.members[i];
    var isNonMember = member.status === "Adult Child - Non Member";
    if (isNonMember)
      continue;    
    if (member.status !== 'Member' && member.status !== 'Designated')
      members.push(member);
  }
  if (members.length > 1) {
    var primary = members[0];
    // Make sure primary member comes first.
    var firstName = getFirstName(primary);
    if (!primary.family.startsWith(firstName)) {
      var first = members[1];
      members[1] = members[0];
      members[0] = first;
    }
  }
  family.emails = [];
  var hasPrimary = false;
  for(var i = 0; i < members.length; i++ ) {
    var isNonMember = members[i].status === "Adult Child - Non Member";
    if (isNonMember)
      continue;    
    hasPrimary = hasPrimary || members[i].status === 'Member' || members[i].status === 'Designated';
    var email = members[i].email;
    if (email !== undefined && email !== "" && !includes(family.emails, email)) {
      family.emails.push(email);
    }
  }
  if (!hasPrimary) {
    Logger.log('All children: ' + family.familyName);
  }
  family.members = members;
}

function createGreeting(family) {
  var greeting = 'Dear%20';
  var lastIndex = family.members.length-1;
  var secondToLastIndex = lastIndex - 1;
  for(var i = 0; i < family.members.length; i++) {
    var firstName = getFirstName(family.members[i]);

    if (i === lastIndex) {
      if (i === 0) {
        greeting += firstName;
        greeting += '%2C';
      } else {
        greeting += '%20%26%20' + firstName;
        greeting += '%2C';
      }
    } else {
      greeting += firstName;
      if (i !== secondToLastIndex) {
        greeting += '%2C%20';
      }
    }
  }
    return greeting;
}

function createFamilyHeader(family, color, familySheet) {
  sortFamily(family);
  var greeting = createGreeting(family);
  //var emails = family.members.length === 1 ? "" : makeMailLink(family.emails, greeting, true);
  var emails = family.emails.length === 0 ? "" : makeMailLink(family.emails, greeting, true);

  rows.push([family.familyName,"", "", "", emails, "", ""]);
  colors.push([color, color, color, color, color, color, color]);
  weights.push(['bold', 'bold', 'bold', 'bold', 'bold', 'bold', 'bold']);

  currentRow++;
}