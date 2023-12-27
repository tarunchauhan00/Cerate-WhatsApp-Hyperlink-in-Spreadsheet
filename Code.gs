function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Create Hyperlink')
    .addItem('Create WhatsApp Hyperlink', 'createWhatsAppHyperlink')
    .addToUi();
}

function createWhatsAppHyperlink() {
  var spreadsheet = SpreadsheetApp.openById('1w9kk-vP0Zn7lfE7ogTQVv7dbBk3z0ybUBhazPS-1My4'); //enter spreadsheet id here
  var sheetName = 'Sheet1'; //enter your sheet name here
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return;
  }

  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(1, 2, lastRow, 4); // Adjust these indices e.g., here 1 is the starting index, 4 is the last index, 2 is the index column where we have phone numbers.

  var data = dataRange.getValues();
  var whatsappLinks = [];

  for (var i = 0; i < data.length; i++) {
    var phoneNumber = data[i][1];
    var message = data[i][0] + " - " + data[i][2] + " - " + data[i][3]; // Here we merge the data like column no. A, C, and D
    var whatsappLink =
      "https://api.whatsapp.com/send?phone=" + phoneNumber + "&text=" + encodeURIComponent(message); // Added "=" after "phone"
    var displayText = "click to send";
    var hyperLinkFormula = '=HYPERLINK("' + whatsappLink + '", "' + displayText + '")'; // Removed extra spaces
    whatsappLinks.push([hyperLinkFormula]);
  }

  var columnE = sheet.getRange(1, 5, lastRow, 1); // Location of column where we want hyperlinks like here we want the hyperlinks in column E i.e., index 5.
  columnE.setFormulas(whatsappLinks);

  Logger.log("WhatsApp hyperlinks with display text have been created successfully");
}
