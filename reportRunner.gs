/**
 * Generic function to run reports for any data source
 * @param {string} sourceTabName - Name of the source data tab
 * @param {Array<string>} programs - Array of program names to filter by
 * @param {string} reportType - Type of report ('K-12' or 'PSKG') for logging
 * @returns {number} - Number of reports generated
 */
const runReportsGeneric_ = (sourceTabName, programs, reportType = 'Report') => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceTab = ss.getSheetByName(sourceTabName);
  const reportTab = ss.getSheetByName(nsSettings.tabs.reportTemplate.name);
  
  // Get configuration
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  const folderTab = ss.getSheetByName(nsSettings.tabs.programFldrIds.name);
  
  // 1. READ: Load all source data once
  const [headers, ...allSourceData] = sourceTab.getDataRange().getValues();
  
  // 2. GET: Districts and folder mappings
  const districts = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.districtColumn)
    .getValues()
    .flat()
    .filter(d => d !== "");
  
  const [folderHeaders, ...folderData] = folderTab
    .getRange(nsSettings.tabs.programFldrIds.range)
    .getDataRegion()
    .getValues();
  
  // Create folder lookup map {districtName: folderId}
  const folderMap = {};
  folderData.forEach(row => {
    folderMap[row[0]] = row[1];
  });
  
  // 3. LOOP: Process each district and program
  let reportCount = 0;
  
  districts.forEach(district => {
    // Filter source data by district (IN MEMORY - instant!)
    const districtData = allSourceData.filter(row => 
      row[nsColumnMaps.filterColumns.district] === district
    );
    
    // Skip if no students in this district
    if (districtData.length === 0) {
      console.log(`No students for district: ${district}`);
      return;
    }
    
    programs.forEach(program => {
      // Filter district data by program (IN MEMORY - instant!)
      const programData = districtData.filter(row => 
        row[nsColumnMaps.filterColumns.program] === program
      );
      
      // Skip if no students in this program
      if (programData.length === 0) {
        console.log(`No students for ${district} - ${program}`);
        return;
      }
      
      // 4. TRANSFORM: Convert to report format
      let reportData = nsColumnMaps.transformToReportFormat(programData);
      
      // 5. ADD: Empty cost column (formula will calculate)
      reportData = nsColumnMaps.addCostColumn(reportData);
      
      // 6. SORT: By student name (optional)
      reportData.sort((a, b) => a[0].localeCompare(b[0]));
      
      // 7. GENERATE: Create and export the report
      generateReport_(reportTab, district, program, reportData, folderMap[district]);
      reportCount++;
      
      console.log(`✓ Generated ${reportType}: ${district} - ${program} (${programData.length} students)`);
    });
  });
  
  return reportCount;
};

/**
 * Generates a single report (works for both K-12 and PSKG)
 * @param {Sheet} reportTab - The report template sheet
 * @param {string} district - District name
 * @param {string} program - Program name
 * @param {Array<Array>} data - Transformed report data
 * @param {string} folderId - Google Drive folder ID for saving
 */
const generateReport_ = (reportTab, district, program, data, folderId) => {
  // Set district and program in header (for formulas to reference)
  reportTab.getRange(nsSettings.tabs.reportTemplate.districtCell).setValue(district);
  reportTab.getRange(nsSettings.tabs.reportTemplate.programCell).setValue(program);
  
  // Paste data
  const dataStartRow = nsSettings.tabs.reportTemplate.dataStartRow;
  const dataRange = reportTab.getRange(dataStartRow, 1, data.length, data[0].length);
  dataRange.setValues(data);
  
  // Force calculation of formulas
  SpreadsheetApp.flush();
  
  // Export as PDF
  const fileName = `${district} - ${program} - ${nsHelpers.getPreviousMonthName()}`;
  exportSheetAsPDF(reportTab, fileName, folderId);
  
  // Clear data for next iteration
  dataRange.clearContent();
};

/**
 * Runs K-12 reports
 * @returns {number} - Number of reports generated
 */
const runK12Reports = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  
  // Get K-12 programs from column E
  const k12Programs = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.k12ProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  return runReportsGeneric_(nsSettings.tabs.sourceData.k12, k12Programs, 'K-12');
};

/**
 * Runs PSKG reports
 * @returns {number} - Number of reports generated
 */
const runPSKGReports = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  
  // Get PSKG programs from column F
  const pskgPrograms = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.pskgProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  return runReportsGeneric_(nsSettings.tabs.sourceData.pskg, pskgPrograms, 'PSKG');
};

/**
 * Test wrapper for K-12 reports (called from menu)
 */
const runK12ReportTest = () => {
  SpreadsheetApp.getActiveSpreadsheet().toast('K-12 Reports starting...', 'Started', -1);
  
  const reportCount = runK12Reports();
  
  if (reportCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Generated ${reportCount} K-12 reports!`, 
      'DONE', 
      5
    );
  }
  
  console.log(`Total K-12 reports generated: ${reportCount}`);
};

/**
 * Test wrapper for PSKG reports (called from menu)
 */
const runPSKGReportTest = () => {
  SpreadsheetApp.getActiveSpreadsheet().toast('PSKG Reports starting...', 'Started', -1);
  
  const reportCount = runPSKGReports();
  
  if (reportCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Generated ${reportCount} PSKG reports!`, 
      'DONE', 
      5
    );
  }
  
  console.log(`Total PSKG reports generated: ${reportCount}`);
};

