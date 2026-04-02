// Main function to run PSKG reports
const runPSKGReports = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceTab = ss.getSheetByName("Preschool Source Data");
  const reportTab = ss.getSheetByName("PSKG Report Template"); // You'll need to create this
  
  // Get configuration
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  const folderTab = ss.getSheetByName(nsSettings.tabs.programFldrIds.name);
  
  // 1. READ: Load all source data once
  const [headers, ...allSourceData] = sourceTab.getDataRange().getValues();
  
  // 2. GET: Districts, programs, and folder mappings
  const districts = dplTab.getRange("A2:A").getValues()
    .flat()
    .filter(d => d !== "");
  
  const programs = ["PS", "KG"]; // Or get from column F if needed
  
  const [folderHeaders, ...folderData] = folderTab.getRange(nsSettings.tabs.programFldrIds.range)
    .getDataRegion()
    .getValues();
  
  // Create folder lookup map {districtName: folderId}
  const folderMap = {};
  folderData.forEach(row => {
    folderMap[row[0]] = row[1]; // Assuming [name, id] structure
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
      generatePSKGReport(reportTab, district, program, reportData, folderMap[district]);
      reportCount++;
      
      console.log(`✓ Generated: ${district} - ${program} (${programData.length} students)`);
    });
  });
  
  return reportCount;
};

/**
 * Generates a single PSKG report
 */
const generatePSKGReport = (reportTab, district, program, data, folderId) => {
  // Set district and program in header (for formulas to reference)
  reportTab.getRange("C1").setValue(district);
  reportTab.getRange("C2").setValue(program);
  
  // Paste data starting at row 4 (adjust to match your template)
  const dataStartRow = 4;
  const dataRange = reportTab.getRange(dataStartRow, 1, data.length, data[0].length);
  dataRange.setValues(data);
  
  // Force calculation of formulas (SUBTOTAL, Cost calculation, etc.)
  SpreadsheetApp.flush();
  
  // Export as PDF - PASS THE SPECIFIC SHEET!
  const fileName = `${district} - ${program} - ${nsHelpers.getPreviousMonthName()}`;
  exportSheetAsPDF(reportTab, fileName, folderId);  // ← Changed from exportCurrentSheetAsPDF
  
  // Clear data for next iteration (keep formulas in header)
  dataRange.clearContent();
};

/**
 * Test wrapper function (call from menu)
 */
const runPSKGReportTest = () => {
  SpreadsheetApp.getActiveSpreadsheet().toast('PSKG Reports starting...', 'Started', -1);
  
  const reportCount = runPSKGReports();
  
  if (reportCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Generated ${reportCount} reports!`, 
      'DONE', 
      5
    );
  }
  
  console.log(`Total reports generated: ${reportCount}`);
};