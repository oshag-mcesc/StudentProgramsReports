function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Report Helpers')
    .addItem('Run Reports', 'runReportTest')
    .addItem('Create Folders', 'runCreateFolders')
    .addToUi();
}