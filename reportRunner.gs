/**
 * @fileoverview Main report generation engine for student billing PDFs.
 * 
 * This file contains the core logic for generating billing reports for both K-12
 * and Preschool/Kindergarten (PSKG) programs. It replaces the old filter-view
 * approach with a much faster in-memory data processing system.
 * 
 * PERFORMANCE IMPROVEMENT:
 * Old approach: ~5-10 minutes (applied filters 70+ times, checked UI for each)
 * New approach: ~30-60 seconds (loads data once, filters in JavaScript)
 * 
 * HOW IT WORKS:
 * 1. Load ALL source data once into memory (one-time cost)
 * 2. Filter by district in JavaScript (instant)
 * 3. Filter by program in JavaScript (instant)
 * 4. Transform to report format (reorder columns)
 * 5. Paste to template and export as PDF
 * 6. Repeat for each district/program combination
 * 
 * KEY ARCHITECTURAL DECISION:
 * We use a generic function (runReportsGeneric_) that works for both K-12 and PSKG.
 * This eliminates code duplication and makes the system easier to maintain.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * Generic report generation function that works for both K-12 and PSKG data.
 * * This is the workhorse function that handles the entire report generation workflow.
 * It's called by both runK12Reports() and runPSKGReports() with different parameters.
 * * WORKFLOW OVERVIEW:
 * 1. SETUP: Get references to spreadsheet tabs (source, template, config)
 * 2. LOAD: Read ALL source data into memory (happens ONCE, not per district)
 * 3. CONFIGURE: Get districts list and folder ID mappings
 * 4. LOOP: For each district → for each program:
 * a. Filter data in memory (JavaScript - very fast)
 * b. Skip if no students (no PDF generated)
 * c. Transform data to report format
 * d. Generate PDF
 * * WHY IN-MEMORY FILTERING IS FAST:
 * - JavaScript array.filter() operates on data already in RAM
 * - No need to apply spreadsheet filters and wait for UI rendering
 * - Can check if students exist instantly (array.length check)
 * - Old approach: apply filter → wait → check UI → clear filter → repeat
 * - New approach: filter in RAM → check length → done
 * * @param {string} sourceTabName - Name of the source data tab to read from
 * @param {string} districtColumn - Range for district list (e.g., "A2:A" or "C2:C")
 * @param {Array<string>} programs - Array of program names to generate reports for
 * @param {string} reportType - Display name for logging (e.g., 'K-12', 'PSKG')
 * @returns {number} Total number of PDF reports generated
 * * @example
 * // Called internally by runK12Reports:
 * const count = runReportsGeneric_(
 * "K-12 Source Data",
 * "A2:A",
 * ["SpecEd", "Gifted", "Speech", "OT"], 
 * "K-12"
 * );
 * console.log(`Generated ${count} reports`);
 * * @private
 */
const runReportsGeneric_ = (sourceTabName, districtColumn, programs, reportType = 'Report') => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Get tab references
  const sourceTab = ss.getSheetByName(sourceTabName);
  const reportTab = ss.getSheetByName(nsSettings.tabs.reportTemplate.name);
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  const folderTab = ss.getSheetByName(nsSettings.tabs.programFldrIds.name);
  
  // ==========================================
  // STEP 1: LOAD ALL SOURCE DATA (ONE TIME)
  // ==========================================
  // This is the key performance improvement. We load ALL data once,
  // then filter it in memory. Old approach loaded data 70+ times.
  const [headers, ...allSourceData] = sourceTab.getDataRange().getValues();
  
  // ==========================================
  // STEP 2: GET CONFIGURATION DATA
  // ==========================================
  
  // Get list of districts to process using the specified column
  // Different columns for K-12 (A) vs PSKG (C)
  // Filter out empty rows (common at end of range)
  const districts = dplTab
    .getRange(districtColumn)  // Use passed parameter instead of hardcoded value
    .getValues()
    .flat()                    // Convert [[val1], [val2]] to [val1, val2]
    .filter(d => d !== "");    // Remove empty strings
  
  // Get folder ID mappings: district name → Google Drive folder ID
  const [folderHeaders, ...folderData] = folderTab
    .getRange(nsSettings.tabs.programFldrIds.range)
    .getDataRegion()           // Auto-expands to include all data
    .getValues();
  // Build lookup object for O(1) folder ID retrieval
  // Instead of searching array each time, we create a map:
  // { "District A": "folder_id_123", "District B": "folder_id_456" }
  const folderMap = {};
  folderData.forEach(row => {
    folderMap[row[0]] = row[1];  // Assumes: [district name, folder ID]
  });

  // ==========================================
  // STEP 3: PROCESS EACH DISTRICT/PROGRAM
  // ==========================================
  let reportCount = 0;
  const errorLog = []; // Array to collect errors across all loop iterations
  
  // Outer loop: Districts (K-12: typically 30-35, PSKG: may be different)
  districts.forEach(district => {
    
    // FILTER #1: Get all students for this district (IN MEMORY - instant!)
    // Uses "How Received District Name" column for filtering (column 14)
    // This is the district being BILLED, not necessarily where student lives
    const districtData = allSourceData.filter(row => 
      row[nsColumnMaps.filterColumns.district] === district
    );
    
    // OPTIMIZATION: Skip entire district if no students
    // This is much faster than applying a filter and checking the UI
    if (districtData.length === 0) {
      console.log(`No students for district: ${district}`);
      return;  // Skip to next district (forEach continues)
    }
    
    // Inner loop: Programs (K-12: 4 programs, PSKG: 2 programs)
    programs.forEach(program => {
      
      // FILTER #2: Get students in this district for this program (IN MEMORY - instant!)
      // Uses "Program Name" column for filtering (column 4)
      const programData = districtData.filter(row => 
        row[nsColumnMaps.filterColumns.program] === program
      );

      // OPTIMIZATION: Skip this program if no students
      // Old approach: had to apply filter and check UI to know this
      // New approach: instant array length check
      if (programData.length === 0) {
        console.log(`No students for ${district} - ${program}`);
        return;  // Skip to next program (forEach continues)
      }
      
      // ==========================================
      // STEP 4: TRANSFORM DATA TO REPORT FORMAT
      // ==========================================
      
      // Transform source data (21 columns) to report format (11 columns)
      // This extracts only the columns we need and reorders them
      let reportData = nsColumnMaps.transformToReportFormat(programData);
      // Add empty 12th column for Cost (calculated by formula in template)
      reportData = nsColumnMaps.addCostColumn(reportData);
      // OPTIONAL: Sort by student name for easier reading
      // Column 0 = Student Name (after transformation)
      reportData.sort((a, b) => a[0].localeCompare(b[0]));

      // ==========================================
      // STEP 5: GENERATE AND EXPORT PDF (WITH TRY-CATCH)
      // ==========================================
      try {
        generateReport_(reportTab, district, program, reportData, folderMap[district]);
        reportCount++;
        
        // Log success (helpful for monitoring long-running jobs)
        console.log(`✓ Generated ${reportType}: ${district} - ${program} (${programData.length} students)`);
      } catch (error) {
        // Log the failure immediately to the execution logs for real-time monitoring
        console.error(`❌ Failed to generate ${reportType} for ${district} - ${program}: ${error.message}`);
        
        // Collect all essential tracking data required for the visual Error Log sheet
        errorLog.push([
          new Date(),                         // Timestamp
          reportType,                         // K-12 or PSKG
          district,                           // Target District
          program,                            // Target Program
          folderMap[district] || "Not Found", // Target Drive Folder ID
          error.message                       // The exact error message thrown
        ]);
      }
    });
  });
  
  // ==========================================
  // STEP 6: WRITE ERRORS TO LOG TAB IF ANY OCCURRED
  // ==========================================
  if (errorLog.length > 0) {
    writeErrorsToLogSheet_(errorLog);
    
    // Throw a modal dialog telling the user exactly what happened
    SpreadsheetApp.getUi().alert(
      `Batch processing finished with errors.\n\n` +
      `Successfully generated: ${reportCount} reports.\n` +
      `Failed: ${errorLog.length} reports.\n\n` +
      `Please open the "Error Log" tab to check details.`
    );
  }
  
  return reportCount;
};

/**
 * Generates a single billing report PDF.
 * * This function handles the final steps of report generation:
 * 1. Write district/program to header cells (for formulas to reference)
 * 2. Paste student data to template
 * 3. Wait for formulas to calculate (SUBTOTAL, Cost)
 * 4. Export as PDF to Google Drive
 * 5. Clear data for next report
 * * WHY WE CLEAR DATA:
 * The report template is reused for every district/program combination.
 * We must clear the data area after each PDF export to avoid mixing
 * data from different reports.
 * * FORMULA DEPENDENCIES:
 * The template has formulas that reference cells C1 (district) and C2 (program).
 * These must be set BEFORE pasting data so formulas calculate correctly.
 * * @param {GoogleAppsScript.Spreadsheet.Sheet} reportTab - The report template sheet
 * @param {string} district - District name (written to C1)
 * @param {string} program - Program name (written to C2)
 * @param {Array<Array>} data - Transformed student data (ready to paste)
 * @param {string} folderId - Google Drive folder ID where PDF will be saved
 * * @example
 * // Called internally by runReportsGeneric_:
 * generateReport_(
 * reportTemplateSheet,
 * "Springfield City",
 * "SpecEd",
 * transformedData,
 * "1abc123def456"
 * );
 * * @private
 */
const generateReport_ = (reportTab, district, program, data, folderId) => {
  
  // ==========================================
  // STEP 1: SET HEADER VALUES
  // ==========================================
  // These cells are referenced by formulas in the template
  // Must be set BEFORE pasting data
  reportTab.getRange(nsSettings.tabs.reportTemplate.districtCell).setValue(district);
  reportTab.getRange(nsSettings.tabs.reportTemplate.programCell).setValue(program);
  
  // ==========================================
  // STEP 2: PASTE STUDENT DATA (WITH EXCEPTION HANDLING)
  // ==========================================
  const dataStartRow = nsSettings.tabs.reportTemplate.dataStartRow; // Row 4
  
  // Calculate range size based on data dimensions
  // data.length = number of rows (students)
  // data[0].length = number of columns (12: 11 data columns + 1 empty cost column)
  const dataRange = reportTab.getRange(dataStartRow, 1, data.length, data[0].length);
  
  try {
    // Write the compiled student rows into the sheet
    dataRange.setValues(data);
    
    // ==========================================
    // STEP 3: WAIT FOR FORMULAS TO CALCULATE
    // ==========================================
    // SpreadsheetApp.flush() forces Google Sheets to execute all pending changes
    // This ensures formulas (SUBTOTAL, Cost calculation) finish before we export PDF
    // Without this, the PDF might contain #N/A or incomplete calculations
    SpreadsheetApp.flush();
    
    // ==========================================
    // STEP 4: EXPORT AS PDF
    // ==========================================
    // Filename format: "District Name - Program - Month"
    // Example: "Springfield City - SpecEd - February"
    const fileName = `${district} - ${program} - ${nsHelpers.getPreviousMonthName()}`;
    
    // exportSheetAsPDF is defined in pdfHelpers.gs
    // It exports the SPECIFIC sheet (not the active sheet) as PDF
    exportSheetAsPDF(reportTab, fileName, folderId);
    
  } finally {
    // ==========================================
    // STEP 5: CLEAR DATA FOR NEXT REPORT
    // ==========================================
    // CRITICAL SECURITY STEP: This block ALWAYS runs, even if the PDF export fails.
    // Clear only the data area, keep the header formulas intact.
    // This prepares the template and prevents data from bleeding into the next report.
    dataRange.clearContent();
  }
};

/**
 * Runs billing reports for K-12 programs.
 * 
 * This is the main entry point for generating K-12 billing reports.
 * Called from the menu: Report Helpers → Run K-12 Reports
 * 
 * K-12 PROGRAMS (typically):
 * - Special Education (SpecEd)
 * - Gifted
 * - Speech Therapy
 * - Occupational Therapy (OT)
 * - Physical Therapy (PT)
 * - etc.
 * 
 * Programs are defined in "District & Program Lists" tab, Column E.
 * 
 * @returns {number} Total number of PDF reports generated
 * 
 * @example
 * // Called from menu or run manually:
 * const count = runK12Reports();
 * console.log(`Generated ${count} K-12 reports`);
 */
const runK12Reports = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  
  // Get K-12 programs from column E
  // Filter out empty rows
  const k12Programs = dplTab
    .getRange(nsSettings.tabs.DistrictProgramLists.k12ProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  // Call generic report runner with K-12 source data and programs
  return runReportsGeneric_(
    nsSettings.tabs.sourceData.k12,  // "K-12 Source Data" tab
    k12Programs,                     // Programs from column E
    'K-12'                          // Display name for logging
  );
};

/**
 * Runs billing reports for Preschool/Kindergarten (PSKG) programs.
 * 
 * This is the main entry point for generating PSKG billing reports.
 * Called from the menu: Report Helpers → Run PSKG Reports
 * 
 * PSKG PROGRAMS:
 * - PS (Preschool)
 * - KG (Kindergarten)
 * 
 * Programs are defined in "District & Program Lists" tab, Column F.
 * 
 * @returns {number} Total number of PDF reports generated
 * 
 * @example
 * // Called from menu or run manually:
 * const count = runPSKGReports();
 * console.log(`Generated ${count} PSKG reports`);
 */
const runPSKGReports = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  
  // Get PSKG programs from column F (typically just "PS" and "KG")
  // Filter out empty rows
  const pskgPrograms = dplTab
    .getRange(nsSettings.tabs.DistrictProgramLists.pskgProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  // Call generic report runner with PSKG source data and programs
  return runReportsGeneric_(
    nsSettings.tabs.sourceData.pskg,  // "Preschool Source Data" tab
    pskgPrograms,                     // Programs from column F
    'PSKG'                           // Display name for logging
  );
};

/**
 * Menu wrapper for K-12 reports with user feedback.
 * 
 * This function is called when the user clicks: Report Helpers → Run K-12 Reports
 * It provides visual feedback via toast notifications during the report generation.
 * 
 * TOAST NOTIFICATIONS:
 * 1. Start: "K-12 Reports starting..." (stays visible)
 * 2. Complete: "Generated X K-12 reports!" (5 second timeout)
 * 
 * @example
 * // Called automatically from menu (see onOpen.gs)
 * // Or run manually for testing
 */
const runK12ReportTest = () => {
  // Show starting notification (timeout = -1 means no auto-dismiss)
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'K-12 Reports starting...', 
    'Started', 
    -1
  );
  
  // Run the actual report generation
  const reportCount = runK12Reports();
  
  // Show completion notification (only if reports were generated)
  if (reportCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Generated ${reportCount} K-12 reports!`, 
      'DONE', 
      5  // Auto-dismiss after 5 seconds
    );
  }
  
  // Log to console for debugging/audit trail
  console.log(`Total K-12 reports generated: ${reportCount}`);
};

/**
 * Menu wrapper for PSKG reports with user feedback.
 * 
 * This function is called when the user clicks: Report Helpers → Run PSKG Reports
 * It provides visual feedback via toast notifications during the report generation.
 * 
 * TOAST NOTIFICATIONS:
 * 1. Start: "PSKG Reports starting..." (stays visible)
 * 2. Complete: "Generated X PSKG reports!" (5 second timeout)
 * 
 * @example
 * // Called automatically from menu (see onOpen.gs)
 * // Or run manually for testing
 */
const runPSKGReportTest = () => {
  // Show starting notification (timeout = -1 means no auto-dismiss)
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'PSKG Reports starting...', 
    'Started', 
    -1
  );
  
  // Run the actual report generation
  const reportCount = runPSKGReports();
  
  // Show completion notification (only if reports were generated)
  if (reportCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Generated ${reportCount} PSKG reports!`, 
      'DONE', 
      5  // Auto-dismiss after 5 seconds
    );
  }
  
  // Log to console for debugging/audit trail
  console.log(`Total PSKG reports generated: ${reportCount}`);
};

/**
 * Appends collected errors to a dedicated "Error Log" tab.
 * Creates the sheet and headers if they do not exist.
 * * @param {Array<Array>} errorRows - 2D array of collected error data
 * @private
 */
const writeErrorsToLogSheet_ = (errorRows) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logTab = ss.getSheetByName("Error Log");
  
  // Create the sheet if it doesn't exist yet
  if (!logTab) {
    logTab = ss.insertSheet("Error Log");
    logTab.appendRow(["Timestamp", "Report Type", "District", "Program", "Target Folder ID", "Error Message"]);
    logTab.getRange("A1:F1").setFontWeight("bold").setBackground("#f4f4f4");
    logTab.setFrozenRows(1);
  }
  
  // Write all collected errors at once to minimize API calls
  const startRow = logTab.getLastRow() + 1;
  logTab.getRange(startRow, 1, errorRows.length, errorRows[0].length).setValues(errorRows);
  
  // Auto-resize columns for readability
  logTab.autoResizeColumns(1, 6);
};
