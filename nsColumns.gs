/**
 * @fileoverview Column mapping configuration for student billing reports.
 * 
 * This file defines how source data columns map to report template columns.
 * Both K-12 and Preschool source data have identical column structures, so we use
 * a single mapping for both report types.
 * 
 * WHY THIS EXISTS:
 * The source data tabs contain 21+ columns, but our billing reports only need 11 of them,
 * and in a specific order. This module handles:
 * 1. Extracting only the columns we need
 * 2. Reordering them to match the report template
 * 3. Identifying which columns to use for filtering (district and program)
 * 
 * PERFORMANCE BENEFIT:
 * By transforming data in-memory (JavaScript arrays) instead of using spreadsheet
 * QUERY formulas, we achieve massive speed improvements (seconds vs minutes).
 * 
 * @author Michael O'Shaughnessy
 * @version 2.0.0
 * @since 2024
 */

/**
 * @namespace nsColumnMaps
 * @description Namespace for column mapping and data transformation functions.
 * 
 * This namespace provides a centralized configuration for how source data columns
 * map to report columns. Both K-12 and PSKG data have identical structures.
 */
const nsColumnMaps = (() => {
  
  /**
   * Source data column indices (0-based).
   * 
   * IMPORTANT: Both "K-12 Source Data" and "Preschool Source Data" tabs have
   * identical column structures. These indices represent the column positions
   * in both source tabs.
   * 
   * Headers in source data:
   * Student Number, Student Name, Birthdate, Grade, Program Name, Address,
   * District Admission Date, District Withdrawal Date, School Name,
   * FS Effective Date, FS Effective End Date, IRN District Of Residence,
   * District of Residence Name, How Received IRN, How Received District Name,
   * Membership Code, Membership Name, Membership Start Date, Membership Stop Date,
   * Days Enrolled, Percent Of Time
   * 
   * @memberof nsColumnMaps
   * @constant
   * @type {Object}
   */
  const SOURCE = {
    studentNumber: 0,              // Student Number (internal ID)
    studentName: 1,                // Student Name (Last, First)
    birthdate: 2,                  // Birthdate (Date)
    grade: 3,                      // Grade (PK, K, 1-12)
    programName: 4,                // Program Name (PS, KG, SpecEd, Gifted, etc.) - FILTER BY THIS
    address: 5,                    // Address (student residence)
    districtAdmissionDate: 6,      // District Admission Date
    districtWithdrawalDate: 7,     // District Withdrawal Date
    schoolName: 8,                 // School Name (where student attends)
    fsEffectiveDate: 9,            // FS Effective Date (funding start)
    fsEffectiveEndDate: 10,        // FS Effective End Date (funding end)
    irnDistrictOfResidence: 11,    // IRN District Of Residence
    districtOfResidenceName: 12,   // District of Residence Name
    howReceivedIRN: 13,            // How Received IRN (billing code)
    howReceivedDistrictName: 14,   // How Received District Name - FILTER BY THIS (billing district)
    membershipCode: 15,            // Membership Code (enrollment type code)
    membershipName: 16,            // Membership Name (enrollment type description)
    membershipStartDate: 17,       // Membership Start Date
    membershipStopDate: 18,        // Membership Stop Date
    daysEnrolled: 19,              // Days Enrolled (used to calculate cost)
    percentOfTime: 20              // Percent Of Time (FTE percentage)
  };
  
  /**
   * Report column order.
   * 
   * This array defines the ORDER of columns in the final billing report.
   * Each element is a reference to a SOURCE column index.
   * 
   * The nsHelpers.keepCols() function uses this array to:
   * 1. Extract only these columns from source data
   * 2. Reorder them to match this sequence
   * 
   * Example: [SOURCE.studentName, SOURCE.birthdate, SOURCE.grade]
   * Would extract columns 1, 2, 3 from source and create a new array with just those columns.
   * 
   * Report columns (in order):
   * 1. Student Name
   * 2. Birthdate
   * 3. Grade
   * 4. Program Name
   * 5. School Name
   * 6. How Received IRN
   * 7. How Received District Name
   * 8. Membership Code
   * 9. Membership Name
   * 10. Days Enrolled
   * 11. Percent Of Time
   * 12. Cost (not in source - calculated by formula in template)
   * 
   * @memberof nsColumnMaps
   * @constant
   * @type {Array<number>}
   */
  const REPORT_COLUMNS = [
    SOURCE.studentName,            // Column 1: Student Name
    SOURCE.birthdate,              // Column 2: Birthdate
    SOURCE.grade,                  // Column 3: Grade
    SOURCE.programName,            // Column 4: Program Name
    SOURCE.schoolName,             // Column 5: School Name
    SOURCE.howReceivedIRN,         // Column 6: How Received IRN
    SOURCE.howReceivedDistrictName,// Column 7: How Received District Name
    //SOURCE.membershipCode,         // Column 8: Membership Code
    //SOURCE.membershipName,         // Column 9: Membership Name
    SOURCE.daysEnrolled,           // Column 10: Days Enrolled
    SOURCE.percentOfTime           // Column 11: Percent Of Time
    // Column 12: Cost - calculated by formula in template (Days * Rate)
  ];
  
  /**
   * Filter column indices.
   * 
   * These specify which source columns to use when filtering data:
   * - district: Filter by "How Received District Name" (column 14)
   *   This is the district being BILLED (not necessarily where student lives)
   * - program: Filter by "Program Name" (column 4)
   *   Examples: PS, KG, SpecEd, Gifted, Speech, etc.
   * 
   * WHY "How Received District Name" for filtering?
   * - This represents the district that receives the service and gets billed
   * - A student might live in District A but attend school in District B
   * - We bill District B (How Received District), not District A (Residence)
   * 
   * @memberof nsColumnMaps
   * @constant
   * @type {Object}
   */
  const FILTER_COLUMNS = {
    district: SOURCE.howReceivedDistrictName,  // Column 14: District to bill
    program: SOURCE.programName                // Column 4: Program type
  };
  
  /**
   * Transforms raw source data into report format.
   * 
   * This function takes the full source data (all 21 columns) and:
   * 1. Extracts only the columns needed for the report (11 columns)
   * 2. Reorders them to match the report template layout
   * 
   * Uses the nsHelpers.keepCols() utility function which:
   * - Filters columns by index
   * - Automatically reorders to match the REPORT_COLUMNS sequence
   * 
   * Example transformation:
   * Source row: [0:"12345", 1:"Smith, John", 2:"1/1/2010", ..., 20:"100%"]
   * Result row: [1:"Smith, John", 2:"1/1/2010", 3:"5", 4:"SpecEd", ...]
   * 
   * @memberof nsColumnMaps
   * @param {Array<Array>} sourceData - Raw 2D array from source data tab
   * @returns {Array<Array>} Transformed 2D array with report columns only
   * 
   * @example
   * const rawData = [
   *   ["12345", "Smith, John", "1/1/2010", "5", "SpecEd", ...],
   *   ["12346", "Doe, Jane", "2/2/2011", "3", "Gifted", ...]
   * ];
   * const reportData = nsColumnMaps.transformToReportFormat(rawData);
   * // Returns: [
   * //   ["Smith, John", "1/1/2010", "5", "SpecEd", ...],
   * //   ["Doe, Jane", "2/2/2011", "3", "Gifted", ...]
   * // ]
   */
  const transformToReportFormat = (sourceData) => {
    return nsHelpers.keepCols(sourceData, REPORT_COLUMNS);
  };
  
  /**
   * Adds an empty Cost column to transformed report data.
   * 
   * The Cost column (column 12) is not in the source data - it's calculated
   * by a formula in the report template: =Days Enrolled × Cost Per Day
   * 
   * This function adds an empty string as the 12th column so that when we paste
   * data into the template, it doesn't overwrite the Cost formula column.
   * 
   * WHY ADD AN EMPTY COLUMN?
   * The report template has a formula in column L (12th column):
   *   =ARRAYFORMULA(IF(J4:J="", "", J4:J * IFNA(INDEX(formulas!$B:$B, MATCH(D4:D, formulas!$A:$A, 0)), 0)))
   * 
   * This formula needs to exist BEFORE we paste data. By adding an empty column,
   * we ensure our data paste operation (11 columns) doesn't accidentally overwrite
   * the formula column.
   * 
   * @memberof nsColumnMaps
   * @param {Array<Array>} reportData - Already transformed report data (11 columns)
   * @returns {Array<Array>} Report data with empty 12th column added
   * 
   * @example
   * const reportData = [
   *   ["Smith, John", "1/1/2010", "5", "SpecEd", ..., "10", "100%"],  // 11 columns
   * ];
   * const withCost = nsColumnMaps.addCostColumn(reportData);
   * // Returns: [
   * //   ["Smith, John", "1/1/2010", "5", "SpecEd", ..., "10", "100%", ""]  // 12 columns
   * // ]
   */
  const addCostColumn = (reportData) => {
    return reportData.map(row => [...row, ""]); // Spread existing row, add empty string
  };
  
  // Return public API
  return {
    source: SOURCE,                          // Expose source column indices
    reportColumns: REPORT_COLUMNS,           // Expose report column order
    filterColumns: FILTER_COLUMNS,           // Expose filter column indices
    transformToReportFormat: transformToReportFormat,  // Transformation function
    addCostColumn: addCostColumn             // Cost column adder function
  };
})();