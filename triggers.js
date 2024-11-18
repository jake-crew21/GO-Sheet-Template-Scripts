
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Custom Menu")
    .addItem("Set Weeks Date", "setWeekDate")
    .addItem("Course By Day", "coursesByDay")
    .addItem("Active Week Con Set","bgConFormating")
    .addItem("Set Weeks","startToEndDate")
    .addItem("Insert Checkboxes and Drop","insertCheckboxes")
    .addItem("Update Staff Needed", "updateStaff")
    .addItem("Resize All Columns", "setColSize")
    .addItem("Update Times", "updateTime")
    .addItem("Remove Cancelled", "cancelled")
    .addSeparator()
    .addItem('Open', 'openDialog')
    .addItem('Get Days Checkin',"writeUpdateCI")
    .addToUi();
}

function onEdit(e) {
  writeUpdateCI();
  // SpreadsheetApp.getUi().alert("Alert message");
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}