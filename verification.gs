/**
 * @fileoverview Configuration verification and troubleshooting utilities.
 * 
 * This file contains functions to verify that the billing report system is
 * properly configured before running reports. It checks:
 * - Required tabs exist
 * - Column mappings are correct
 * - Folder IDs are valid
 * - Data sources have data
 * - Programs and districts are configured
 * 
 * WHEN TO USE:
 * - After initial setup
 * - Before running reports for a new month
 * - When troubleshooting errors
 * - After making configuration changes
 * 
 * HOW TO RUN:
 * Call verifyConfiguration() from the Script Editor and check the console output.
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * Main verification function - runs all configuration checks.
 * 
 * This function performs a comprehensive health check of the billing system.
 * It verifies that all required components are in place and properly configured.
 * 
 * CHECKS PERFORMED:
 * 1. Required tabs exist
 * 2. Source data tabs have data
 * 3. District and program lists are populated
 * 4. Folder mappings exist and are valid
 * 5. Column mappings are correct
 * 6. Report template is properly set up
 * 
 * OUTPUT:
 * Results are logged to the console with ✓ (pass) or ✗ (fail) indicators.
 * A summary is displayed at the end.
 * 
 * @returns {Object} Summary object with pass/fail counts
 * 
 * @example
 * // Run from Script Editor
 * const results = verifyConfiguration();
 * console.log(`Passed: ${results.passed}, Failed: ${results.failed}`);
 */
function verifyConfiguration() {
  console.log('='.repeat(60));
  console.log('BILLING REPORT SYSTEM - CONFIGURATION VERIFICATION');
  console.log('='.repeat(60));
  console.log('');
  
  let passed = 0;
  let failed = 0;
  
  // Track individual check results
  const checks = [
    { name: 'Required Tabs', fn: checkRequiredTabs },
    { name: 'Source Data', fn: checkSourceData },
    { name: 'Districts and Programs', fn: checkDistrictsAndPrograms },
    { name: 'Folder Mappings', fn: checkFolderMappings },
    { name: 'Column Mappings', fn: checkColumnMappings },
    { name: 'Report Template', fn: checkReportTemplate }
  ];
  
  // Run each check
  checks.forEach(check => {
    console.log(`\n--- ${check.name} ---`);
    try {
      const result = check.fn();
      if (result.success) {
        passed++;
        console.log(`✓ ${check.name} - PASSED`);
      } else {
        failed++;
        console.log(`✗ ${check.name} - FAILED`);
        if (result.errors) {
          result.errors.forEach(err => console.log(`  • ${err}`));
        }
      }
    } catch (err) {
      failed++;
      console.log(`✗ ${check.name} - ERROR: ${err.message}`);
    }
  });
  
  // Print summary
  console.log('\n' + '='.repeat(60));
  console.log('SUMMARY');
  console.log('='.repeat(60));
  console.log(`Passed: ${passed}/${checks.length}`);
  console.log(`Failed: ${failed}/${checks.length}`);
  
  if (failed === 0) {
    console.log('\n✓ ALL CHECKS PASSED - System is ready to run reports!');
  } else {
    console.log('\n✗ SOME CHECKS FAILED - Please fix the issues above before running reports.');
  }
  
  return { passed, failed, total: checks.length };
}

/**
 * Verifies that all required spreadsheet tabs exist.
 * 
 * REQUIRED TABS:
 * - District & Program Lists: Configuration for what to process
 * - K-12 Source Data: Raw K-12 enrollment data
 * - Preschool Source Data: Raw PSKG enrollment data
 * - Report Template: Template for PDF generation
 * - foldrIds: Folder ID mappings for PDF saving
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkRequiredTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  
  const requiredTabs = [
    nsSettings.tabs.DistrictProgramLists.name,
    nsSettings.tabs.sourceData.k12,
    nsSettings.tabs.sourceData.pskg,
    nsSettings.tabs.reportTemplate.name,
    nsSettings.tabs.programFldrIds.name
  ];
  
  requiredTabs.forEach(tabName => {
    const tab = ss.getSheetByName(tabName);
    if (!tab) {
      errors.push(`Missing required tab: "${tabName}"`);
    } else {
      console.log(`  ✓ Tab exists: "${tabName}"`);
    }
  });
  
  return { success: errors.length === 0, errors };
}

/**
 * Verifies that source data tabs contain data.
 * 
 * CHECKS:
 * - Tabs are not empty
 * - Have header row plus at least one data row
 * - Column count matches expected structure (21 columns)
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkSourceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  
  // Check K-12 Source Data
  const k12Tab = ss.getSheetByName(nsSettings.tabs.sourceData.k12);
  if (k12Tab) {
    const k12Data = k12Tab.getDataRange().getValues();
    if (k12Data.length < 2) {
      errors.push(`K-12 Source Data has no data rows (only ${k12Data.length} rows)`);
    } else {
      console.log(`  ✓ K-12 Source Data: ${k12Data.length - 1} rows`);
    }
    
    if (k12Data[0].length !== 21) {
      errors.push(`K-12 Source Data has ${k12Data[0].length} columns, expected 21`);
    } else {
      console.log(`  ✓ K-12 Source Data: 21 columns (correct)`);
    }
  }
  
  // Check PSKG Source Data
  const pskgTab = ss.getSheetByName(nsSettings.tabs.sourceData.pskg);
  if (pskgTab) {
    const pskgData = pskgTab.getDataRange().getValues();
    if (pskgData.length < 2) {
      errors.push(`Preschool Source Data has no data rows (only ${pskgData.length} rows)`);
    } else {
      console.log(`  ✓ Preschool Source Data: ${pskgData.length - 1} rows`);
    }
    
    if (pskgData[0].length !== 21) {
      errors.push(`Preschool Source Data has ${pskgData[0].length} columns, expected 21`);
    } else {
      console.log(`  ✓ Preschool Source Data: 21 columns (correct)`);
    }
  }
  
  return { success: errors.length === 0, errors };
}

/**
 * Verifies that district and program lists are populated.
 * 
 * CHECKS:
 * - District list (column A) has entries
 * - K-12 program list (column E) has entries
 * - PSKG program list (column F) has entries
 * - No duplicate districts
 * - No duplicate programs
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkDistrictsAndPrograms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  
  if (!dplTab) {
    errors.push('District & Program Lists tab not found');
    return { success: false, errors };
  }
  
  // Check districts (Column A)
  const districts = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.districtColumn)
    .getValues()
    .flat()
    .filter(d => d !== "");
  
  if (districts.length === 0) {
    errors.push('No districts found in column A');
  } else {
    console.log(`  ✓ Districts: ${districts.length} configured`);
    
    // Check for duplicates
    const uniqueDistricts = [...new Set(districts)];
    if (uniqueDistricts.length !== districts.length) {
      errors.push(`Duplicate districts found (${districts.length} total, ${uniqueDistricts.length} unique)`);
    }
  }
  
  // Check K-12 programs (Column E)
  const k12Programs = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.k12ProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  if (k12Programs.length === 0) {
    errors.push('No K-12 programs found in column E');
  } else {
    console.log(`  ✓ K-12 Programs: ${k12Programs.length} configured (${k12Programs.join(', ')})`);
  }
  
  // Check PSKG programs (Column F)
  const pskgPrograms = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.pskgProgramColumn)
    .getValues()
    .flat()
    .filter(p => p !== "");
  
  if (pskgPrograms.length === 0) {
    errors.push('No PSKG programs found in column F');
  } else {
    console.log(`  ✓ PSKG Programs: ${pskgPrograms.length} configured (${pskgPrograms.join(', ')})`);
  }
  
  // Estimate total reports
  const totalK12Reports = districts.length * k12Programs.length;
  const totalPSKGReports = districts.length * pskgPrograms.length;
  console.log(`  ℹ Expected K-12 reports: ${totalK12Reports} (${districts.length} districts × ${k12Programs.length} programs)`);
  console.log(`  ℹ Expected PSKG reports: ${totalPSKGReports} (${districts.length} districts × ${pskgPrograms.length} programs)`);
  
  return { success: errors.length === 0, errors };
}

/**
 * Verifies that folder ID mappings are valid.
 * 
 * CHECKS:
 * - foldrIds tab has data
 * - Each folder ID is valid and accessible
 * - All districts have folder mappings
 * - No missing folders
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkFolderMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  const folderTab = ss.getSheetByName(nsSettings.tabs.programFldrIds.name);
  
  if (!folderTab) {
    errors.push('foldrIds tab not found');
    return { success: false, errors };
  }
  
  const [headers, ...folderData] = folderTab
    .getRange(nsSettings.tabs.programFldrIds.range)
    .getDataRegion()
    .getValues();
  
  if (folderData.length === 0) {
    errors.push('No folder mappings found in foldrIds tab');
    return { success: false, errors };
  }
  
  console.log(`  ✓ Folder mappings: ${folderData.length} configured`);
  
  // Verify each folder ID is valid
  let validCount = 0;
  let invalidCount = 0;
  
  folderData.forEach(row => {
    const [districtName, folderId] = row;
    try {
      const folder = DriveApp.getFolderById(folderId);
      validCount++;
    } catch (err) {
      errors.push(`Invalid folder ID for "${districtName}": ${folderId}`);
      invalidCount++;
    }
  });
  
  console.log(`  ✓ Valid folder IDs: ${validCount}/${folderData.length}`);
  if (invalidCount > 0) {
    console.log(`  ✗ Invalid folder IDs: ${invalidCount}`);
  }
  
  // Check if all districts have folders
  const dplTab = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name);
  if (dplTab) {
    const districts = dplTab.getRange(nsSettings.tabs.DistrictProgramLists.districtColumn)
      .getValues()
      .flat()
      .filter(d => d !== "");
    
    const folderDistricts = folderData.map(row => row[0]);
    const missingFolders = districts.filter(d => !folderDistricts.includes(d));
    
    if (missingFolders.length > 0) {
      errors.push(`Districts missing folder mappings: ${missingFolders.join(', ')}`);
    } else {
      console.log(`  ✓ All districts have folder mappings`);
    }
  }
  
  return { success: errors.length === 0, errors };
}

/**
 * Verifies that column mappings are correct.
 * 
 * CHECKS:
 * - Source column headers match expected structure
 * - Column indices are within valid range
 * - Filter columns exist in source data
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkColumnMappings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  
  // Expected source headers (first 21 columns)
  const expectedHeaders = [
    "Student Number", "Student Name", "Birthdate", "Grade", "Program Name",
    "Address", "District Admission Date", "District Withdrawal Date", "School Name",
    "FS Effective Date", "FS Effective End Date", "IRN District Of Residence",
    "District of Residence Name", "How Received IRN", "How Received District Name",
    "Membership Code", "Membership Name", "Membership Start Date", "Membership Stop Date",
    "Days Enrolled", "Percent Of Time"
  ];
  
  // Check K-12 headers
  const k12Tab = ss.getSheetByName(nsSettings.tabs.sourceData.k12);
  if (k12Tab) {
    const k12Headers = k12Tab.getRange(1, 1, 1, 21).getValues()[0];
    let headerMatches = 0;
    
    for (let i = 0; i < expectedHeaders.length; i++) {
      if (k12Headers[i] === expectedHeaders[i]) {
        headerMatches++;
      } else {
        errors.push(`K-12 column ${i} mismatch: expected "${expectedHeaders[i]}", got "${k12Headers[i]}"`);
      }
    }
    
    console.log(`  ✓ K-12 headers match: ${headerMatches}/${expectedHeaders.length}`);
  }
  
  // Check PSKG headers
  const pskgTab = ss.getSheetByName(nsSettings.tabs.sourceData.pskg);
  if (pskgTab) {
    const pskgHeaders = pskgTab.getRange(1, 1, 1, 21).getValues()[0];
    let headerMatches = 0;
    
    for (let i = 0; i < expectedHeaders.length; i++) {
      if (pskgHeaders[i] === expectedHeaders[i]) {
        headerMatches++;
      } else {
        errors.push(`PSKG column ${i} mismatch: expected "${expectedHeaders[i]}", got "${pskgHeaders[i]}"`);
      }
    }
    
    console.log(`  ✓ PSKG headers match: ${headerMatches}/${expectedHeaders.length}`);
  }
  
  // Verify filter columns are within range
  if (nsColumnMaps.filterColumns.district > 20 || nsColumnMaps.filterColumns.district < 0) {
    errors.push(`District filter column index (${nsColumnMaps.filterColumns.district}) is out of range`);
  } else {
    console.log(`  ✓ District filter column: ${nsColumnMaps.filterColumns.district} (${expectedHeaders[nsColumnMaps.filterColumns.district]})`);
  }
  
  if (nsColumnMaps.filterColumns.program > 20 || nsColumnMaps.filterColumns.program < 0) {
    errors.push(`Program filter column index (${nsColumnMaps.filterColumns.program}) is out of range`);
  } else {
    console.log(`  ✓ Program filter column: ${nsColumnMaps.filterColumns.program} (${expectedHeaders[nsColumnMaps.filterColumns.program]})`);
  }
  
  return { success: errors.length === 0, errors };
}

/**
 * Verifies that the report template is properly configured.
 * 
 * CHECKS:
 * - Template tab exists
 * - Header cells (C1, C2) are accessible
 * - Data start row is valid
 * - Cost formula exists in column L
 * 
 * @returns {{success: boolean, errors: Array<string>}} Check result
 * 
 * @private
 */
function checkReportTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  
  const templateTab = ss.getSheetByName(nsSettings.tabs.reportTemplate.name);
  
  if (!templateTab) {
    errors.push('Report Template tab not found');
    return { success: false, errors };
  }
  
  console.log(`  ✓ Report Template tab exists`);
  
  // Check that header cells exist
  try {
    const districtCell = templateTab.getRange(nsSettings.tabs.reportTemplate.districtCell);
    console.log(`  ✓ District cell (${nsSettings.tabs.reportTemplate.districtCell}) accessible`);
  } catch (err) {
    errors.push(`Cannot access district cell: ${nsSettings.tabs.reportTemplate.districtCell}`);
  }
  
  try {
    const programCell = templateTab.getRange(nsSettings.tabs.reportTemplate.programCell);
    console.log(`  ✓ Program cell (${nsSettings.tabs.reportTemplate.programCell}) accessible`);
  } catch (err) {
    errors.push(`Cannot access program cell: ${nsSettings.tabs.reportTemplate.programCell}`);
  }
  
  // Check data start row is valid
  const dataStartRow = nsSettings.tabs.reportTemplate.dataStartRow;
  if (dataStartRow < 1 || dataStartRow > 1000) {
    errors.push(`Data start row (${dataStartRow}) seems invalid`);
  } else {
    console.log(`  ✓ Data start row: ${dataStartRow}`);
  }
  
  // Check if Cost formula exists in column L (row 4)
  try {
    const costCell = templateTab.getRange(`L${dataStartRow}`);
    const formula = costCell.getFormula();
    if (formula && formula.includes('ARRAYFORMULA')) {
      console.log(`  ✓ Cost formula exists in column L`);
    } else if (formula) {
      console.log(`  ⚠ Cost column has a formula, but not ARRAYFORMULA`);
    } else {
      errors.push('Cost formula missing in column L - should contain ARRAYFORMULA');
    }
  } catch (err) {
    errors.push(`Error checking cost formula: ${err.message}`);
  }
  
  return { success: errors.length === 0, errors };
}

/**
 * Quick test to verify column mapping with actual data.
 * 
 * This function tests the column transformation with the first row of
 * actual source data to ensure the mapping produces correct results.
 * 
 * Useful for debugging column mapping issues.
 * 
 * @example
 * // Run from Script Editor
 * testColumnMapping();
 */
function testColumnMapping() {
  console.log('='.repeat(60));
  console.log('COLUMN MAPPING TEST');
  console.log('='.repeat(60));
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const k12Tab = ss.getSheetByName(nsSettings.tabs.sourceData.k12);
  
  if (!k12Tab) {
    console.log('✗ K-12 Source Data tab not found');
    return;
  }
  
  // Get headers and first data row
  const [headers, firstRow] = k12Tab.getRange("A1:U2").getValues();
  
  console.log('\n--- Source Data (first row) ---');
  headers.forEach((header, idx) => {
    console.log(`[${idx}] ${header}: ${firstRow[idx]}`);
  });
  
  // Transform the first row
  const transformed = nsColumnMaps.transformToReportFormat([firstRow]);
  const withCost = nsColumnMaps.addCostColumn(transformed);
  
  console.log('\n--- Transformed Data (report format) ---');
  const reportHeaders = [
    "Student Name", "Birthdate", "Grade", "Program Name", "School Name",
    "How Received IRN", "How Received District Name", "Membership Code",
    "Membership Name", "Days Enrolled", "Percent Of Time", "Cost"
  ];
  
  reportHeaders.forEach((header, idx) => {
    console.log(`[${idx}] ${header}: ${withCost[0][idx]}`);
  });
  
  console.log('\n✓ Column mapping test complete');
}