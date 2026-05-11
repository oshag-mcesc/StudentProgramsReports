/**
 * @fileoverview Central configuration for the Student Programs Billing Report system.
 * 
 * This namespace contains all tab names, column ranges, and folder IDs used throughout
 * the application. By centralizing these settings, we make the code easier to maintain
 * and update when spreadsheet structure changes.
 * 
 * IMPORTANT: When adding new tabs or changing spreadsheet structure, update this file first.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * @namespace nsSettings
 * @description Central configuration namespace for all spreadsheet tabs, ranges, and settings.
 * 
 * This namespace uses an IIFE (Immediately Invoked Function Expression) pattern to create
 * a clean namespace without polluting the global scope.
 */
const nsSettings = (() => {
  const ns = {};

  /**
   * Tab names and configurations for the spreadsheet.
   * 
   * REMOVED: Monthly K-12 and PS tabs (Aug-Sep, Oct, Nov, etc.) are no longer needed
   * because we now read directly from source data tabs and generate reports on-demand.
   * Historical data is preserved in archived PDFs in Google Drive folders.
   * 
   * @memberof nsSettings
   * @property {Object} tabs - Collection of tab configurations
   */
  ns.tabs = {
    /**
     * District and Program Lists - Configuration tab
     * Contains the master list of districts and programs to process
     */
    DistrictProgramLists: {
      name: "District & Program Lists",
      districtColumn: "A2:A",      // Column A: Districts to loop through
      k12ProgramColumn: "E2:E",    // Column E: K-12 programs (e.g., SpecEd, Gifted, etc.)
      pskgProgramColumn: "F2:F"    // Column F: PSKG programs (PS, KG only)
    },
    
    /**
     * Source Data Tabs - Raw data downloads from student information system
     * These tabs contain unfiltered student enrollment data
     */
    sourceData: {
      k12: "K-12 Source Data",         // K-12 student data (grades 1-12)
      pskg: "Preschool Source Data"     // Preschool/Kindergarten student data
    },
    
    /**
     * Report Template - Universal template used for both K-12 and PSKG reports
     * Contains header formulas (SUBTOTAL, cost calculations) and formatting
     */
    reportTemplate: {
      name: "Report Template",
      districtCell: "C1",        // Cell where district name is written
      programCell: "C2",         // Cell where program name is written
      dataStartRow: 4            // Row where student data begins (row 1-3 are headers)
    },
    
    /**
     * Folder IDs - Google Drive folder mapping
     * Maps district names to their corresponding Google Drive folder IDs
     */
    programFldrIds: {
      name: "foldrIds",
      range: "A1"               // Starting cell for folder data (auto-expands via getDataRegion)
    },
    
    /**
     * LEGACY: Student Info tab (kept for potential future use)
     * Currently not used in the main report generation workflow
     */
    studInfo: {
      name: "Student_Info",
      headers: ["SSID", "Name", "Days"]
    },
    
    /**
     * LEGACY: K-12 District/Program Search tab
     * This tab was used in the old filter-view approach and may be removed
     * in a future cleanup. Keeping for now in case we need to reference old data.
     */
    K12DistrictProgramSearch: {
      name: "K-12 District/Program Search",
      range: "A3:Q",
      checkRange: "G1"
    }
  };

  /**
   * Extract Info Configuration (LEGACY)
   * Used by older data extraction functions - may be deprecated
   * 
   * @memberof nsSettings
   */
  ns.extractInfo = {
    range: "A1",
    colsKeep: [0, 1, 19]      // Columns to keep: Student Number, Name, Days Enrolled
  };

  /**
   * Root folder ID for saving generated PDFs and creating district subfolders
   * 
   * @memberof nsSettings
   * @type {string}
   */
  ns.saveFolderID = "13L0urjlK0b2YfthCIPEtQZ6AFDQRbmG7";

  // Return the namespace object to make it accessible
  return ns;
})();