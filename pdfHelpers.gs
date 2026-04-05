/**
 * @fileoverview PDF export utilities for Google Sheets.
 * 
 * This file contains functions for exporting Google Sheets as PDF files and
 * saving them to specific Google Drive folders.
 * 
 * CREDIT:
 * Modified from PDFbyExport script by Spencer Easton
 * Original: https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
 * 
 * HOW IT WORKS:
 * Google Sheets has an undocumented export API that allows programmatic PDF generation.
 * These functions use reverse-engineered URL parameters to control PDF formatting:
 * - Page size, orientation, margins
 * - Grid lines, page numbers, sheet selection
 * - Specific range export (optional)
 * 
 * IMPORTANT: These parameters are NOT officially documented by Google and may
 * change over time. If PDF export suddenly breaks, check Google's latest API.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * Global setting for PDF save location (LEGACY - not used in current implementation).
 * 
 * Originally used to toggle between saving to Drive root vs spreadsheet's folder.
 * Current implementation always uses explicit folder IDs passed to export functions.
 * 
 * @deprecated This setting is no longer used. Folder IDs are passed explicitly.
 * @type {boolean}
 */
var saveToRootFolder = false;

/**
 * Converts a Google Sheet to a PDF Blob.
 * 
 * This is the core function that communicates with Google's export API.
 * It constructs a special URL with parameters that tell Google Sheets how
 * to format the PDF.
 * 
 * EXPORT URL PARAMETERS (reverse-engineered):
 * - exportFormat=pdf: Generate PDF instead of other formats
 * - size=LETTER: Paper size (LETTER, A4, LEGAL, etc.)
 * - portrait=false: Landscape orientation (true = portrait)
 * - fitw=true: Fit content to page width
 * - margins: top, bottom, left, right in inches
 * - gridlines=true: Show spreadsheet grid lines in PDF
 * - pagenum: Page number position (CENTER, UNDEFINED to hide)
 * - sheetnames=false: Don't show sheet name in header
 * - printtitle=false: Don't repeat row/column headers on each page
 * - gid: Sheet ID to export (specific sheet)
 * - r1, r2, c1, c2: Range to export (optional)
 * 
 * AUTHENTICATION:
 * Uses ScriptApp.getOAuthToken() to authenticate the export request.
 * This token represents the script's authorization to access the spreadsheet.
 * 
 * RATE LIMITING:
 * If Google returns HTTP 429 (Too Many Requests), the function retries up to
 * 5 times with 3-second delays. This prevents failures when generating many PDFs.
 * 
 * @param {string} url - The spreadsheet URL (from SpreadsheetApp.getUrl())
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to export
 * @param {GoogleAppsScript.Spreadsheet.Range} [range] - Optional: specific range to export
 * @returns {GoogleAppsScript.Base.Blob} PDF file as a Blob object
 * 
 * @throws {Error} If export fails after 5 retry attempts
 * 
 * @example
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("Report Template");
 * const blob = _getAsBlob(ss.getUrl(), sheet);
 * // blob can now be saved as a file
 * 
 * @private
 */
function _getAsBlob(url, sheet, range) {
  // ==========================================
  // BUILD RANGE PARAMETERS (if range specified)
  // ==========================================
  var rangeParam = '';
  if (range) {
    // r1, r2 = row start/end (0-based)
    // c1, c2 = column start/end (0-based)
    rangeParam =
      '&r1=' + (range.getRow() - 1)           // Start row (convert to 0-based)
      + '&r2=' + range.getLastRow()           // End row
      + '&c1=' + (range.getColumn() - 1)      // Start column (convert to 0-based)
      + '&c2=' + range.getLastColumn();       // End column
  }
  
  // ==========================================
  // BUILD SHEET PARAMETER
  // ==========================================
  var sheetParam = '';
  if (sheet) {
    // gid = Sheet ID (unique identifier for each sheet in the spreadsheet)
    sheetParam = '&gid=' + sheet.getSheetId();
  }
  
  // ==========================================
  // CONSTRUCT EXPORT URL
  // ==========================================
  // Start with base spreadsheet URL, remove /edit and everything after
  // Then append export parameters
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'              // Paper size: LETTER, A4, LEGAL, TABLOID
      + '&portrait=false'           // false = landscape, true = portrait
      + '&fitw=true'                // Fit to width
      + '&top_margin=0.75'          // Top margin in inches
      + '&bottom_margin=0.75'       // Bottom margin in inches
      + '&left_margin=0.7'          // Left margin in inches
      + '&right_margin=0.7'         // Right margin in inches
      + '&sheetnames=false'         // Don't show sheet names in header
      + '&printtitle=false'         // Don't repeat row/column titles on each page
      + '&pagenum=UNDEFINED'        // Page numbers: CENTER, UNDEFINED (hide)
      + '&gridlines=true'           // Show grid lines
      + '&fzr=FALSE'                // Frozen rows/columns (FALSE = don't repeat)
      + sheetParam                  // Add sheet ID parameter
      + rangeParam;                 // Add range parameters (if specified)
  
  // ==========================================
  // FETCH PDF WITH RETRY LOGIC
  // ==========================================
  var response;
  var i = 0;
  
  // Retry up to 5 times if we hit rate limits
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,    // Don't throw errors, return response object
      headers: { 
        // OAuth token authenticates the request as this script
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
    });
    
    // Check response code
    if (response.getResponseCode() === 429) {
      // HTTP 429 = Too Many Requests (rate limited)
      // Wait 3 seconds and try again
      console.log('Rate limited, retrying in 3 seconds...');
      Utilities.sleep(3000);
    } else {
      // Success or other error - break out of retry loop
      break;
    }
  }
  
  // If we exhausted all retries, throw error
  if (i === 5) {
    throw new Error('PDF export failed after 5 retries. Too many sheets to print or rate limit exceeded.');
  }
  
  // Convert response to Blob (binary file object)
  return response.getBlob();
}

/**
 * Saves a PDF Blob to a Google Drive folder.
 * 
 * This function takes a Blob (binary PDF data) and saves it as a file
 * in the specified Google Drive folder.
 * 
 * BLOB OPERATIONS:
 * - setName(): Renames the Blob (becomes the filename)
 * - createFile(): Saves the Blob as a file in Drive
 * 
 * REMOVED FEATURE:
 * Original implementation showed a modal dialog with link to the PDF.
 * This was removed to avoid interrupting automated batch operations.
 * When generating 60-70 PDFs, showing a dialog for each would be disruptive.
 * 
 * @param {GoogleAppsScript.Base.Blob} blob - PDF data as a Blob
 * @param {string} fileName - Desired filename (without .pdf extension - added automatically)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The source spreadsheet (unused, kept for compatibility)
 * @param {string} folderID - Google Drive folder ID where PDF should be saved
 * @returns {GoogleAppsScript.Drive.File} The created PDF file
 * 
 * @example
 * const blob = _getAsBlob(url, sheet);
 * const pdfFile = _exportBlob(blob, "District A - SpecEd - February", ss, "1abc123def");
 * console.log("PDF saved:", pdfFile.getUrl());
 * 
 * @private
 */
function _exportBlob(blob, fileName, spreadsheet, folderID) {
  // Set the filename for this Blob
  // Note: Google Drive automatically adds .pdf extension
  blob = blob.setName(fileName);
  
  // Get the target folder from Drive
  let folder = DriveApp.getFolderById(folderID);
  
  // Create the PDF file in the folder
  var pdfFile = folder.createFile(blob);
  
  // REMOVED: Modal dialog showing PDF link
  // Reason: When generating 60+ PDFs, showing a dialog for each is disruptive
  // If you need the URL, access it via pdfFile.getUrl() in calling code
  
  return pdfFile;
}

/**
 * Exports a specific sheet as a PDF to a Google Drive folder.
 * 
 * This is the main public function used by the report generation system.
 * It exports a SPECIFIC sheet (not the currently active sheet) as PDF.
 * 
 * WHY THIS EXISTS:
 * The original exportCurrentSheetAsPDF() exported whatever sheet was currently
 * visible in the UI. This caused problems when generating multiple reports
 * because the script might be viewing a different sheet.
 * 
 * This function solves that by accepting the sheet object as a parameter,
 * ensuring we always export the correct sheet regardless of what's active.
 * 
 * USAGE IN REPORT SYSTEM:
 * Called by generateReport_() in reportRunner.gs for each district/program combination.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The specific sheet to export
 * @param {string} fileName - PDF filename (without extension)
 * @param {string} folderID - Google Drive folder ID for saving
 * 
 * @example
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const reportSheet = ss.getSheetByName("Report Template");
 * exportSheetAsPDF(reportSheet, "District A - SpecEd - February", "1abc123def");
 */
function exportSheetAsPDF(sheet, fileName, folderID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get PDF blob for the SPECIFIC sheet (not active sheet)
  var blob = _getAsBlob(spreadsheet.getUrl(), sheet);
  
  // Save blob to Google Drive
  _exportBlob(blob, fileName, spreadsheet, folderID);
}

/**
 * Exports the currently active sheet as a PDF (LEGACY FUNCTION).
 * 
 * This function is kept for backward compatibility with any old code that
 * might still call it. New code should use exportSheetAsPDF() instead.
 * 
 * LIMITATION:
 * This function exports whatever sheet is currently visible/active in the UI.
 * During automated batch operations, this can lead to exporting the wrong sheet.
 * 
 * RECOMMENDATION:
 * Use exportSheetAsPDF() instead, which accepts a sheet parameter and always
 * exports the correct sheet regardless of what's active.
 * 
 * @param {string} fileName - PDF filename (without extension)
 * @param {string} folderID - Google Drive folder ID for saving
 * 
 * @deprecated Use exportSheetAsPDF() instead for explicit sheet control
 * 
 * @example
 * // OLD WAY (don't use):
 * exportCurrentSheetAsPDF("MyReport", "1abc123def");
 * 
 * // NEW WAY (recommended):
 * const sheet = ss.getSheetByName("Report Template");
 * exportSheetAsPDF(sheet, "MyReport", "1abc123def");
 */
function exportCurrentSheetAsPDF(fileName, folderID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the currently active/visible sheet
  var currentSheet = SpreadsheetApp.getActiveSheet();
  
  // Export it
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet);
  _exportBlob(blob, fileName, spreadsheet, folderID);
}