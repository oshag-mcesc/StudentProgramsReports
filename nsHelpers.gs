/**
 * @fileoverview Helper utility functions for data manipulation and formatting.
 * 
 * This file contains reusable utility functions that support the billing report
 * generation system. These functions handle common tasks like:
 * - Column extraction and filtering from 2D arrays
 * - Date formatting and month name retrieval
 * - Google Drive folder operations
 * 
 * These utilities are designed to be generic and reusable across different
 * parts of the application.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * @namespace nsHelpers
 * @description Namespace containing utility helper functions.
 * 
 * This namespace uses an IIFE pattern to create a clean, encapsulated set of
 * utility functions without polluting the global scope.
 */
const nsHelpers = (function () {
  
  /**
   * Extracts and reorders specified columns from a 2D array.
   * 
   * This is a critical function for the report generation system. It takes a 2D array
   * (like data from a spreadsheet) and returns a new array containing only the
   * specified columns, in the order specified.
   * 
   * HOW IT WORKS:
   * 1. Iterate through each row of the input array
   * 2. For each row, filter to keep only columns whose indices are in the 'columns' array
   * 3. The order of indices in 'columns' determines the order in the output
   * 
   * IMPORTANT: This function both FILTERS and REORDERS columns
   * - If columns = [2, 0, 5], you get column 2, then column 0, then column 5
   * - The output has 3 columns, in that specific order
   * 
   * @memberof nsHelpers
   * @param {Array<Array>} array - The 2D array to extract columns from
   * @param {Array<number>} columns - Array of column indices to keep (0-based)
   * @returns {Array<Array>} New 2D array containing only the specified columns
   * 
   * @example
   * // Extract columns 1, 3, and 5 from source data
   * const sourceData = [
   *   ["A", "B", "C", "D", "E", "F"],
   *   ["1", "2", "3", "4", "5", "6"]
   * ];
   * const result = keepCols_(sourceData, [1, 3, 5]);
   * // Returns: [
   * //   ["B", "D", "F"],
   * //   ["2", "4", "6"]
   * // ]
   * 
   * @example
   * // Reorder columns - get column 3, then 1, then 0
   * const sourceData = [["A", "B", "C", "D"]];
   * const result = keepCols_(sourceData, [3, 1, 0]);
   * // Returns: [["D", "B", "A"]]
   */
  const keepCols_ = (array, columns) => {
    return array.map(row => row.filter((_, index) => columns.includes(index)));
  };

  /**
   * Returns the full name of the current month.
   * 
   * Uses JavaScript's built-in Intl.DateTimeFormat (via toLocaleString) to get
   * the month name in the default locale (typically English for US users).
   * 
   * This function is useful for timestamping reports or displaying current month
   * information in the UI.
   * 
   * @memberof nsHelpers
   * @returns {string} The full name of the current month (e.g., "January", "February")
   * 
   * @example
   * // If today is March 15, 2024
   * const month = nsHelpers.getCurrentMonthName();
   * console.log(month); // "March"
   */
  const getCurrentMonthName_ = () => {
    const currentDate = new Date();
    
    // toLocaleString with { month: 'long' } returns full month name
    // 'default' locale uses the browser/system default (usually en-US)
    const monthName = currentDate.toLocaleString('default', { month: 'long' });
    
    return monthName;
  };

  /**
   * Returns the full name of the previous month.
   * 
   * This function is critical for billing report filenames. Since we typically
   * generate billing reports AFTER the month ends (e.g., generate February reports
   * in early March), we need the PREVIOUS month's name for the filename.
   * 
   * HOW IT WORKS:
   * 1. Get the current date
   * 2. Subtract 1 from the current month (JavaScript handles year rollover automatically)
   * 3. Format the resulting month as a full name
   * 
   * EDGE CASE HANDLING:
   * - If current month is January (month 0), setMonth(-1) automatically rolls back
   *   to December of the previous year. JavaScript Date handles this correctly.
   * 
   * @memberof nsHelpers
   * @returns {string} The full name of the previous month
   * 
   * @example
   * // If today is March 15, 2024
   * const prevMonth = nsHelpers.getPreviousMonthName();
   * console.log(prevMonth); // "February"
   * 
   * @example
   * // If today is January 15, 2024
   * const prevMonth = nsHelpers.getPreviousMonthName();
   * console.log(prevMonth); // "December" (from 2023)
   */
  const getPreviousMonthName_ = () => {
    const currentDate = new Date();
    
    // Subtract one month
    // Note: setMonth() modifies the date object in place
    currentDate.setMonth(currentDate.getMonth() - 1);
    
    // Format as full month name
    const previousMonthName = currentDate.toLocaleString('default', { month: 'long' });
    
    return previousMonthName;
  };

  /**
   * Retrieves a list of all subfolders within a given Google Drive folder.
   * 
   * This function is useful for:
   * - Auditing folder structure
   * - Verifying that district folders exist before running reports
   * - Generating folder ID mappings for the foldrIds tab
   * 
   * The function accepts either a folder ID or a full Google Drive URL,
   * automatically extracting the ID from the URL if needed.
   * 
   * @memberof nsHelpers
   * @param {string} folderIdOrUrl - Google Drive folder ID or full folder URL
   * @returns {Array<{name: string, id: string}>} Array of objects with folder name and ID
   * 
   * @example
   * // Using a folder ID
   * const folders = nsHelpers.getSubfolderList("1abc123def456");
   * console.log(folders);
   * // [
   * //   { name: "District A", id: "1xyz..." },
   * //   { name: "District B", id: "2abc..." }
   * // ]
   * 
   * @example
   * // Using a folder URL
   * const url = "https://drive.google.com/drive/folders/1abc123def456";
   * const folders = nsHelpers.getSubfolderList(url);
   * // Same result as above
   */
  const getSubfolderList_ = (folderIdOrUrl) => {
    // Extract folder ID from URL if a URL was provided
    // URL format: https://drive.google.com/drive/folders/{FOLDER_ID}
    // split('/').pop() gets the last segment after the final slash
    const parentFolder = folderIdOrUrl.includes('/')
      ? DriveApp.getFolderById(folderIdOrUrl.split('/').pop())
      : DriveApp.getFolderById(folderIdOrUrl);

    // Get iterator of all subfolders
    const subfolders = parentFolder.getFolders();

    // Build array of folder info objects
    const folderList = [];
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      folderList.push({
        name: folder.getName(),
        id: folder.getId()
      });
    }

    return folderList;
  };

  /**
   * Converts an array of folder objects to a 2D array suitable for spreadsheet output.
   * 
   * This function takes the output from getSubfolderList() and converts it into
   * a format that can be directly pasted into a Google Sheet.
   * 
   * The resulting 2D array has:
   * - Row 1: Header row with column names ["name", "id"]
   * - Row 2+: Data rows with folder names and IDs
   * 
   * This format matches what range.setValues() expects for pasting into a sheet.
   * 
   * @memberof nsHelpers
   * @param {Array<{name: string, id: string}>} folderList - Array of folder objects from getSubfolderList()
   * @returns {Array<Array<string>>} 2D array with header row and data rows
   * 
   * @example
   * const folders = [
   *   { name: "District A", id: "1abc..." },
   *   { name: "District B", id: "2xyz..." }
   * ];
   * const array2D = nsHelpers.convertToArray2D(folders);
   * console.log(array2D);
   * // [
   * //   ["name", "id"],                    // Header row
   * //   ["District A", "1abc..."],         // Data row 1
   * //   ["District B", "2xyz..."]          // Data row 2
   * // ]
   * 
   * // Can be pasted directly into a sheet:
   * const sheet = SpreadsheetApp.getActiveSheet();
   * sheet.getRange(1, 1, array2D.length, array2D[0].length).setValues(array2D);
   */
  const convertToArray2D_ = (folderList) => {
    // Get header row from the keys of the first object
    // Object.keys({ name: "...", id: "..." }) returns ["name", "id"]
    const header = Object.keys(folderList[0]);

    // Convert each object to an array of its values
    // Object.values({ name: "District A", id: "1abc..." }) returns ["District A", "1abc..."]
    const data = folderList.map(folder => Object.values(folder));

    // Combine header and data using spread operator
    return [header, ...data];
  };

  // Return public API - these functions are accessible as nsHelpers.functionName()
  return {
    keepCols: keepCols_,
    getCurrentMonthName: getCurrentMonthName_,
    getPreviousMonthName: getPreviousMonthName_,
    getSubfolderList: getSubfolderList_,
    convertToArray2D: convertToArray2D_
  };
})();


// ============================================================================
// UTILITY FUNCTIONS (used for setup/maintenance, not in main report flow)
// ============================================================================

/**
 * Test function to verify folder retrieval and 2D array conversion.
 * 
 * This function demonstrates how to:
 * 1. Get a list of subfolders from a Google Drive folder
 * 2. Convert that list to a 2D array
 * 3. Paste the result into a spreadsheet
 * 
 * USAGE: Run this once to populate a "temp" tab with folder info,
 * then use that data to build your foldrIds tab.
 * 
 * @example
 * // To use: Update the URL to your parent folder, then run checkFOlders()
 */
const checkFOlders = () => {
  // Replace this URL with your actual Google Drive folder URL
  let idorurl = "https://drive.google.com/drive/folders/1bAg-5r3RII6RRjfxwQVD6CNDBtU7BUxL";
  
  // Get subfolder list
  let rslt1 = nsHelpers.getSubfolderList(idorurl);
  
  // Convert to 2D array
  let rslt = nsHelpers.convertToArray2D(rslt1);
  
  // Paste into spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tab = ss.getSheetByName("temp");
  tab.getRange(1, 1, rslt.length, rslt[0].length).setValues(rslt);
};

/**
 * Utility function to create "Student Programs" subfolder in each district folder.
 * 
 * This was used during initial setup to create a consistent folder structure.
 * Each district folder needed a "Student Programs" subfolder where the billing
 * PDFs are saved.
 * 
 * SETUP SCRIPT - Run once during initial configuration
 * 
 * PREREQUISITES:
 * - Must have a "temp" tab with folder names in column A and folder IDs in column B
 * 
 * WHAT IT DOES:
 * 1. Reads folder list from "temp" tab
 * 2. For each folder, creates a "Student Programs" subfolder
 * 3. Writes the new subfolder IDs to column C
 * 
 * @example
 * // To use:
 * // 1. Run checkFOlders() first to populate the temp tab
 * // 2. Run subFolderIDgetter() to create Student Programs subfolders
 * // 3. Copy the IDs from column C to your foldrIds tab
 */
const subFolderIDgetter = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let folderList = ss.getSheetByName("temp");
  
  // Get folder data (skip header row)
  let folderData = folderList.getRange("A1").getDataRegion().getValues();
  folderData.shift(); // Remove header row

  let studentPrgmFldr = [];

  // Create "Student Programs" subfolder in each folder
  folderData.forEach(record => {
    let fldr = DriveApp.getFolderById(record[1]); // Column B = folder ID
    let sp = fldr.createFolder("Student Programs");
    studentPrgmFldr.push([sp.getId()]);
  });
  
  // Write subfolder IDs to column C
  folderList.getRange(2, 3, studentPrgmFldr.length, 1).setValues(studentPrgmFldr);
};