/**
 * @fileoverview Google Drive folder management utilities.
 * 
 * This file contains functions for creating and managing Google Drive folders
 * for the billing report system. Each district needs a dedicated folder where
 * their billing PDFs are saved.
 * 
 * TYPICAL USAGE:
 * 1. Run once during initial setup to create district folders
 * 2. Folder IDs are recorded in the "foldrIds" tab
 * 3. Report generation uses these folder IDs to save PDFs
 * 
 * FOLDER STRUCTURE:
 * Root Folder (nsSettings.saveFolderID)
 *   └── District A
 *       └── Student Programs (billing PDFs saved here)
 *   └── District B
 *       └── Student Programs
 *   └── District C
 *       └── Student Programs
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * @namespace nsFolders
 * @description Namespace for folder creation and management functions.
 * 
 * This namespace encapsulates folder-related functionality to avoid
 * polluting the global scope and to organize related functions together.
 */
const nsFolders = (() => {
  
  /**
   * Creates Google Drive folders for each district/school.
   * 
   * This function is typically run ONCE during initial setup of the billing system.
   * It reads a list of school/district names from a spreadsheet and creates a
   * corresponding folder in Google Drive for each one.
   * 
   * WORKFLOW:
   * 1. Get list of school names from spreadsheet (column H)
   * 2. For each school, create a folder in Google Drive
   * 3. Write folder name and ID back to spreadsheet (columns J and K)
   * 
   * WHY COLUMN H?
   * This appears to be a temporary/test location. In production, you might
   * want to read from the "District & Program Lists" tab instead.
   * 
   * AFTER RUNNING THIS:
   * - Copy the folder names and IDs from columns J-K
   * - Paste into the "foldrIds" tab (columns A-B)
   * - This creates the mapping used by the report generation system
   * 
   * PERMISSIONS:
   * - You must have EDIT permission to the parent folder (nsSettings.saveFolderID)
   * - Anyone running this script must have the same permission
   * 
   * TODO NOTES (from original code):
   * - Consider implementing dynamic tab name retrieval (user input or config)
   * - Consider adding folder hierarchy support (subfolders)
   * - Consider user selection of parent folder
   * 
   * @memberof nsFolders
   * @returns {boolean} True if folders were created successfully, false on error
   * 
   * @example
   * // Run from menu: Report Helpers → Create Folders
   * // Or run directly:
   * const success = nsFolders.createFolders();
   * if (success) {
   *   console.log("Folders created successfully");
   * }
   * 
   * @throws {Error} Logs error to console if folder creation fails
   */
  const createFolders_ = () => {
    try {
      // Get spreadsheet and tab
      // TODO: Make this more robust - currently hardcoded to specific tab
      const { ss, tab } = getSheetTab(nsSettings.tabs.DistrictProgramLists.name);

      // Get school names from column H (starting at H1)
      // getDataRegion() auto-expands to include all contiguous data
      // flat() converts [[name1], [name2]] to [name1, name2]
      const schools = tab.getRange("H1").getDataRegion().getValues().flat();
      
      // Remove header row (first element)
      schools.shift();
      
      // Alternative approach (commented out):
      // const schools = tab.getRange(2, 8, tab.getLastRow()-1, 1).getValues().flat();
      // This would get column H from row 2 to last row, skipping header
      
      // Array to store folder information [name, id] pairs
      const folderInfo = [];

      // Get parent folder where district folders will be created
      // nsSettings.saveFolderID should point to your root "Student Programs" folder
      const parentFolder = DriveApp.getFolderById(nsSettings.saveFolderID);

      // Create a folder for each school/district
      schools.forEach(school => {
        const newFolder = parentFolder.createFolder(school);
        folderInfo.push([newFolder.getName(), newFolder.getId()]);
      });

      // Write folder names and IDs to spreadsheet
      // Columns J and K (columns 10 and 11), starting at row 2
      // This is temporary - you'll copy these to the foldrIds tab
      tab.getRange(2, 10, folderInfo.length, 2).setValues(folderInfo);
      
      return true;
      
    } catch (err) {
      // Log error but don't throw - allows calling code to continue
      console.log('Error creating folders:', err);
      return false;
    }
  };

  // Return public API
  return {
    createFolders: createFolders_,
  };
})();

/**
 * Retrieves a spreadsheet and a specific sheet by name.
 * 
 * This is a convenience function that bundles two common operations:
 * 1. Get the active spreadsheet
 * 2. Get a specific sheet by name
 * 
 * USAGE PATTERN:
 * Instead of writing:
 *   const ss = SpreadsheetApp.getActiveSpreadsheet();
 *   const tab = ss.getSheetByName("SomeTab");
 * 
 * You can write:
 *   const { ss, tab } = getSheetTab("SomeTab");
 * 
 * This uses ES6 destructuring to unpack the returned object.
 * 
 * @param {string} tabName - The name of the sheet to retrieve
 * @returns {{ss: GoogleAppsScript.Spreadsheet.Spreadsheet, tab: GoogleAppsScript.Spreadsheet.Sheet}} 
 *          Object with 'ss' (spreadsheet) and 'tab' (sheet) properties
 * 
 * @example
 * const { ss, tab } = getSheetTab("District & Program Lists");
 * console.log(tab.getName()); // "District & Program Lists"
 * 
 * @example
 * // You can also access properties individually:
 * const info = getSheetTab("Report Template");
 * info.tab.getRange("A1").setValue("Hello");
 */
const getSheetTab = (tabName) => {
  const info = {};
  info.ss = SpreadsheetApp.getActiveSpreadsheet();
  info.tab = info.ss.getSheetByName(tabName);
  
  return info;
};

/**
 * Menu wrapper function for folder creation with user feedback.
 * 
 * This function is called when the user selects: Report Helpers → Create Folders
 * It runs the folder creation process and logs the result.
 * 
 * WHEN TO USE:
 * - Initial setup of the billing system
 * - Adding new districts that need folders
 * - Recreating folder structure after reorganization
 * 
 * AFTER RUNNING:
 * 1. Check the spreadsheet - folder info should appear in columns J-K
 * 2. Copy the folder names and IDs
 * 3. Paste into the "foldrIds" tab
 * 4. (Optional) Create "Student Programs" subfolders using subFolderIDgetter() in nsHelpers.gs
 * 
 * @example
 * // Called automatically from menu (see onOpen.gs)
 * // Or run manually:
 * runCreateFolders();
 */
const runCreateFolders = () => {
  const result = nsFolders.createFolders();
  
  if (result) {
    console.log('✓ Folders created successfully');
    console.log('Next steps:');
    console.log('1. Copy folder names and IDs from columns J-K');
    console.log('2. Paste into the "foldrIds" tab');
  } else {
    console.log('✗ Folder creation failed - check error logs');
  }
};