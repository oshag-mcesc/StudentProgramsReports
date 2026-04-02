const nsColumnMaps = (() => {
  
  // Source data column indices (0-based)
  // NOTE: Both K-12 Source Data and Preschool Source Data have identical column structure
  const SOURCE = {
    studentNumber: 0,              // Student Number
    studentName: 1,                // Student Name
    birthdate: 2,                  // Birthdate
    grade: 3,                      // Grade
    programName: 4,                // Program Name
    address: 5,                    // Address
    districtAdmissionDate: 6,      // District Admission Date
    districtWithdrawalDate: 7,     // District Withdrawal Date
    schoolName: 8,                 // School Name
    fsEffectiveDate: 9,            // FS Effective Date
    fsEffectiveEndDate: 10,        // FS Effective End Date
    irnDistrictOfResidence: 11,    // IRN District Of Residence
    districtOfResidenceName: 12,   // District of Residence Name
    howReceivedIRN: 13,            // How Received IRN
    howReceivedDistrictName: 14,   // How Received District Name (FILTER BY THIS!)
    membershipCode: 15,            // Membership Code
    membershipName: 16,            // Membership Name
    membershipStartDate: 17,       // Membership Start Date
    membershipStopDate: 18,        // Membership Stop Date
    daysEnrolled: 19,              // Days Enrolled
    percentOfTime: 20              // Percent Of Time
  };
  
  // Report columns in exact order for the template
  // Same for both K-12 and PSKG reports
  const REPORT_COLUMNS = [
    SOURCE.studentName,            // 1. Student Name
    SOURCE.birthdate,              // 2. Birthdate
    SOURCE.grade,                  // 3. Grade
    SOURCE.programName,            // 4. Program Name
    SOURCE.schoolName,             // 5. School Name
    SOURCE.howReceivedIRN,         // 6. How Received IRN
    SOURCE.howReceivedDistrictName,// 7. How Received District Name
    SOURCE.membershipCode,         // 8. Membership Code
    SOURCE.membershipName,         // 9. Membership Name
    SOURCE.daysEnrolled,           // 10. Days Enrolled
    SOURCE.percentOfTime           // 11. Percent Of Time
    // Cost (column 12) is calculated by formula in template
  ];
  
  // Columns used for filtering
  const FILTER_COLUMNS = {
    district: SOURCE.howReceivedDistrictName,  // Column 14
    program: SOURCE.programName                // Column 4
  };
  
  return {
    source: SOURCE,
    reportColumns: REPORT_COLUMNS,
    filterColumns: FILTER_COLUMNS,
    
    /**
     * Transforms source data to report format
     * @param {Array<Array>} sourceData - Raw data from source tab
     * @returns {Array<Array>} - Data in report column order
     */
    transformToReportFormat: (sourceData) => {
      return nsHelpers.keepCols(sourceData, REPORT_COLUMNS);
    },
    
    /**
     * Adds empty Cost column to transformed data (for formula to fill)
     * @param {Array<Array>} reportData - Already transformed data
     * @returns {Array<Array>} - Data with empty cost column added
     */
    addCostColumn: (reportData) => {
      return reportData.map(row => [...row, ""]); // Add empty string for Cost
    }
  };
})();