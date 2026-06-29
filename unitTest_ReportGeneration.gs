/**
 * ==========================================================================
 * UNIT TESTS FOR REPORT GENERATION ERROR HANDLING
 * ==========================================================================
 * * PURPOSE:
 * This file contains isolated tests to verify that our new try-catch-finally 
 * architecture behaves correctly under both successful and failing conditions.
 * * WHAT IS A UNIT TEST?
 * A unit test runs a small, isolated "unit" of code (like a single function) 
 * with controlled inputs to verify it produces the exact expected outcome.
 * * HOW TO RUN:
 * 1. Open the Script Editor.
 * 2. Select "test_errorHandlingArchitecture" from the function dropdown.
 * 3. Click "Run" and watch the Execution Log.
 */

function test_errorHandlingArchitecture() {
  console.log('='.repeat(60));
  console.log('STARTING UNIT TEST: ERROR HANDLING ARCHITECTURE');
  console.log('='.repeat(60));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportTab = ss.getSheetByName(nsSettings.tabs.reportTemplate.name);
  
  if (!reportTab) {
    console.error('✗ TEST ABORTED: Report Template tab not found.');
    return;
  }

  // ==========================================
  // SETUP MOCK DATA (CONTROLLED INPUTS)
  // ==========================================
  const testDistrict = "TEST DISTRICT LLC";
  const testProgram = "TEST PROGRAM";
  
  // Create fake student data matching the 12 columns expected by the template
  const mockStudentData = [
    ["Test Student A", "2015-01-01", "5th", testProgram, "Test Elementary", "123456", "Test District", "M1", "Reg", "20", "1.0", ""],
    ["Test Student B", "2015-02-02", "5th", testProgram, "Test Elementary", "123456", "Test District", "M1", "Reg", "20", "1.0", ""]
  ];

  console.log('\n--- UNIT TEST 1: TESTING THE "FINALLY" BLOCK DATA CLEARANCE ---');
  
  // An invalid Drive Folder ID string that WILL force a crash in exportSheetAsPDF
  const guaranteedBadFolderId = "NOT_A_REAL_FOLDER_ID_FOR_TESTING"; 
  
  let errorWasCaught = false;

  try {
    console.log('ℹ Invoking generateReport_ with an intentionally BAD Folder ID...');
    
    // This call WILL fail inside exportSheetAsPDF when it tries to access this Drive ID
    generateReport_(reportTab, testDistrict, testProgram, mockStudentData, guaranteedBadFolderId);
    
  } catch (error) {
    errorWasCaught = true;
    console.log(`✓ Success: generateReport_ threw the expected error: "${error.message}"`);
  }

  // ==========================================
  // VERIFY THE "FINALLY" CLAUSE WORKED
  // ==========================================
  // Even though it crashed, the template data area MUST be completely empty right now.
  const dataStartRow = nsSettings.tabs.reportTemplate.dataStartRow;
  const testRange = reportTab.getRange(dataStartRow, 1, mockStudentData.length, mockStudentData[0].length);
  const remainingValues = testRange.getValues();
  
  // Check if all cells in the test row are cleared (empty strings)
  const isDataAreaClear = remainingValues.flat().every(cell => cell === "");

  if (errorWasCaught && isDataAreaClear) {
    console.log('✓ TEST 1 PASSED: Code crashed as expected, but the template data was safely WIPED clean.');
  } else {
    console.error('✗ TEST 1 FAILED: Either the code did not crash, or student data was left stuck on the template.');
  }


  console.log('\n--- UNIT TEST 2: TESTING THE VISUAL ERROR LOG WRITER ---');
  
  // Define a sample row of error data matching our Error Log structure
  const sampleErrorRows = [[
    new Date(),
    "UNIT-TEST",
    testDistrict,
    testProgram,
    guaranteedBadFolderId,
    "Simulated error message for unit testing verification."
  ]];

  try {
    console.log('ℹ Sending a test entry to the Error Log tab via writeErrorsToLogSheet_...');
    writeErrorsToLogSheet_(sampleErrorRows);
    
    // Check if the tab was created and the row was written
    const logTab = ss.getSheetByName("Error Log");
    if (logTab) {
      const lastRowValues = logTab.getRange(logTab.getLastRow(), 1, 1, 6).getValues()[0];
      if (lastRowValues[2] === testDistrict && lastRowValues[1] === "UNIT-TEST") {
        console.log('✓ TEST 2 PASSED: The "Error Log" tab exists and recorded the error accurately.');
      } else {
        console.error('✗ TEST 2 FAILED: Error Log sheet exists, but the written data does not match.');
      }
    } else {
      console.error('✗ TEST 2 FAILED: "Error Log" sheet was not created.');
    }
  } catch (e) {
    console.error(`✗ TEST 2 FAILED: writeErrorsToLogSheet_ crashed: ${e.message}`);
  }

  console.log('\n' + '='.repeat(60));
  console.log('UNIT TESTING COMPLETE');
  console.log('='.repeat(60));
}

