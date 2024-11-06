/**
 * @namespace nsHelpers
 * @description A namespace containing helper functions.
 */
const nsHelpers = (function () {
  /**
   * Extracts specified columns from a 2D array.
   *
   * @param {Array} array - The 2D array to extract columns from.
   * @param {Array} columns - An array containing the indices of the columns to extract.
   * @returns {Array} A new array containing only the specified columns from each row.
   */
  const keepCols_ = (array, columns) => {
    return array.map(row => row.filter((_, index) => columns.includes(index)));
  };

  /**
 * Returns the full name of the current month.
 * 
 * This function uses the `toLocaleString` method of the Date object
 * to get the name of the current month based on the default locale.
 *
 * @returns {string} The full name of the current month (e.g., "January", "February").
 */
const getCurrentMonthName_=()=> {
  // Create a new Date object representing the current date and time
  const currentDate = new Date();

  // Use toLocaleString to format the month as a long (full) name
  const monthName = currentDate.toLocaleString('default', { month: 'long' });

  // Return the full name of the current month
  return monthName;
}


  return {
    keepCols: keepCols_,
    getCurrentMonthName:getCurrentMonthName_
  };
})();