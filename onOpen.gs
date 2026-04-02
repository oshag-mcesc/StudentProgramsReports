function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Report Helpers')
    .addItem('Run K-12 Reports', 'runK12ReportTest')
    .addItem('Run PSKG Reports', 'runPSKGReportTest')
    .addSeparator()
    .addItem('Create Folders', 'runCreateFolders')
    .addToUi();
}