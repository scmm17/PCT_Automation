function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();

  var menuItems = [
    {
      name: "Generate Spreadsheet",
      functionName: "makeFamiliesSheet"
    },
    {
      name: "Send Weekly Emails",
      functionName: "sendWeeklyEmails"
    },
    {
      name: "Reset Last Family",
      functionName: "resetLastFamily"
    }
  ];
  spreadsheet.addMenu("Trinity", menuItems);
}
