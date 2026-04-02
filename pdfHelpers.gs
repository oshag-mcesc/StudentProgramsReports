/**
 * This is modified from PDFbyExport script for spreadsheet: https://docs.google.com/spreadsheets/d/10gp1ksm-MSkMqEKd0kU9gNOWHY7qRZUxIUCis3KhxSE/edit?gid=1601234331#gid=1601234331
 */

// By default, PDFs are saved in your Drive Root folder
// To save in the same folder as the spreadsheet, change the value to 'false' without the single quote pair
// You must have EDIT permission to the same folder
var saveToRootFolder = false

function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=false'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam
      
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

function _exportBlob(blob, fileName, spreadsheet, folderID) {
  blob = blob.setName(fileName)
  //var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  let folder = DriveApp.getFolderById(folderID)
  var pdfFile = folder.createFile(blob)
  
  // // Display a modal dialog box with custom HtmlService content.
  // const htmlOutput = HtmlService
  //   .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
  //   .setWidth(300)
  //   .setHeight(80)
  // SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}

//the main PDF function
function exportSheetAsPDF(sheet, fileName, folderID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  var blob = _getAsBlob(spreadsheet.getUrl(), sheet)  // Use passed sheet, not active
  _exportBlob(blob, fileName, spreadsheet, folderID)
}

// Keep old function for backward compatibility with K-12 reports
function exportCurrentSheetAsPDF(fileName, folderID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = SpreadsheetApp.getActiveSheet()
  
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
  _exportBlob(blob, fileName, spreadsheet, folderID)
}
