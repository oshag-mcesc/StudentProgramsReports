/**
 * @fileoverview Custom menu setup for the billing report system.
 * 
 * This file creates a custom menu in the Google Sheets UI that provides
 * easy access to the main report generation functions.
 * 
 * MENU STRUCTURE:
 * Report Helpers
 *   ├── Run K-12 Reports
 *   ├── Run PSKG Reports
 *   ├── ───────────────── (separator)
 *   └── Create Folders
 * 
 * The onOpen() function is a special Google Apps Script trigger that runs
 * automatically when the spreadsheet is opened.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * Creates custom menu when spreadsheet is opened.
 * 
 * This is a "simple trigger" - Google Apps Script automatically calls this
 * function whenever the spreadsheet is opened by any user.
 * 
 * SIMPLE TRIGGER LIMITATIONS:
 * - Runs with the permissions of the user opening the spreadsheet
 * - Cannot access services that require authorization if user hasn't authorized
 * - Cannot run longer than 30 seconds
 * 
 * These limitations don't affect menu creation, but they're good to know
 * if you expand this function in the future.
 * 
 * MENU ITEMS:
 * - "Run K-12 Reports": Generates billing PDFs for K-12 programs (SpecEd, Gifted, etc.)
 * - "Run PSKG Reports": Generates billing PDFs for Preschool/Kindergarten programs
 * - "Create Folders": Setup utility for creating district folders in Google Drive
 * 
 * CUSTOMIZATION:
 * To add more menu items, use:
 *   .addItem('Menu Text', 'functionName')
 *   .addSeparator()  // Adds a horizontal line
 *   .addSubMenu()    // Creates a submenu
 * 
 * @example
 * // This function runs automatically on spreadsheet open
 * // To manually refresh the menu, run this function from Script Editor
 * 
 * @see https://developers.google.com/apps-script/guides/menus
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Report Helpers')
    .addItem('Run K-12 Reports', 'runK12ReportTest')        // K-12 billing reports
    .addItem('Run PSKG Reports', 'runPSKGReportTest')       // PSKG billing reports
    .addSeparator()                                         // Visual separator line
    .addItem('Create Folders', 'runCreateFolders')          // Setup utility
    .addToUi();                                             // Adds menu to UI
}