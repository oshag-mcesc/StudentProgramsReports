const verifyPSKGColumnMapping = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceTab = ss.getSheetByName("Preschool Source Data");
  
  // Get headers and first data row
  const [sourceHeaders, firstRow] = sourceTab.getRange("A1:V2").getValues();
  
  // Transform the first row
  const transformed = nsColumnMaps.transformToReportFormat([firstRow]);
  const withCost = nsColumnMaps.addCostColumn(transformed);
  
  // Expected report headers
  const reportHeaders = [
    "Student Name",
    "Birthdate",
    "Grade",
    "Program Name",
    "School Name",
    "How Received IRN",
    "How Received District Name",
    "Membership Code",
    "Membership Name",
    "Days Enrolled",
    "Percent Of Time",
    "Cost"
  ];
  
  console.log("=== VERIFICATION ===");
  console.log("\nSource Headers:", sourceHeaders);
  console.log("\nFirst Source Row:", firstRow);
  console.log("\nTransformed Row:", withCost[0]);
  console.log("\nExpected Report Headers:", reportHeaders);
  console.log("\n=== MAPPING ===");
  
  reportHeaders.forEach((header, idx) => {
    if (idx < withCost[0].length - 1) { // Skip Cost (it's empty)
      const sourceIdx = nsColumnMaps.reportColumns[idx];
      console.log(`${idx + 1}. ${header} ← Source Column ${sourceIdx} (${sourceHeaders[sourceIdx]}): ${withCost[0][idx]}`);
    } else {
      console.log(`${idx + 1}. ${header} ← (Calculated by formula)`);
    }
  });
};