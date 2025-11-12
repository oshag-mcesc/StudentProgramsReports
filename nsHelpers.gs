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
  const getCurrentMonthName_ = () => {
    // Create a new Date object representing the current date and time
    const currentDate = new Date();

    // Use toLocaleString to format the month as a long (full) name
    const monthName = currentDate.toLocaleString('default', { month: 'long' });

    // Return the full name of the current month
    return monthName;
  }

  /**
 * Gets the name of the previous month.
 * 
 * This function creates a Date object representing the current date, 
 * then subtracts one month to get the previous month. It returns the 
 * full name of the previous month as a string (e.g., "January", "February").
 *
 * @returns {string} The full name of the previous month.
 */
const getPreviousMonthName_ = () => {
  // Create a new Date object representing the current date
  const currentDate = new Date();

  // Subtract one month by setting the month to the previous one
  currentDate.setMonth(currentDate.getMonth() - 1);

  // Get the full name of the previous month
  const previousMonthName = currentDate.toLocaleString('default', { month: 'long' });

  // Return the full name of the previous month
  return previousMonthName;
};

  /**
   * Retrieves an array of folder names and IDs for all subfolders in a given folder.
   * 
   * @param {string} folderIdOrUrl - The ID or URL of the parent folder.
   * @returns {Array<{name: string, id: string}>} - An array of objects containing the name and ID of each subfolder.
   */
  const getSubfolderList_ = (folderIdOrUrl) => {
    // Get the parent folder
    const parentFolder = folderIdOrUrl.includes('/')
      ? DriveApp.getFolderById(folderIdOrUrl.split('/').pop())
      : DriveApp.getFolderById(folderIdOrUrl);

    // Get an array of all subfolders
    const subfolders = parentFolder.getFolders();

    // Create an array of folder names and IDs
    const folderList = [];
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      folderList.push({
        name: folder.getName(),
        id: folder.getId()
      });
    }

    return folderList;
  };

  /**
   * Converts an array of subfolder objects to a 2D array with a header row.
   * 
   * @param {Array<{name: string, id: string}>} folderList - An array of folder name and ID objects.
   * @returns {Array<Array<string>>} - A 2D array with a header row and the folder name in the first column and the folder ID in the second column.
   */
  const convertToArray2D_ = (folderList) => {
    // Get the header row from the keys of the first object in the array
    const header = Object.keys(folderList[0]);

    // Convert the folder list to a 2D array
    const data = folderList.map(folder => Object.values(folder));

    // Add the header row to the beginning of the array
    return [header, ...data];
  };


  return {
    keepCols: keepCols_,
    getCurrentMonthName: getCurrentMonthName_,
    getSubfolderList: getSubfolderList_,
    convertToArray2D: convertToArray2D_,
    getPreviousMonthName:getPreviousMonthName_
  };
})();




const checkFOlders = () => {
  let idorurl = "https://drive.google.com/drive/folders/1bAg-5r3RII6RRjfxwQVD6CNDBtU7BUxL"
  let rslt1 = nsHelpers.getSubfolderList(idorurl)
  let rslt =  nsHelpers.convertToArray2D(rslt1)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let tab = ss.getSheetByName("temp")
  tab.getRange(1,1,rslt.length,rslt[0].length).setValues(rslt)
  //console.log(rslt);
}


const subFolderIDgetter = ()=>{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let folderList = ss.getSheetByName("temp")
  let folderData = folderList.getRange("A1").getDataRegion().getValues()
  folderData.shift()

  let studentPrgmFldr =[]

  folderData.forEach(record =>{
    let fldr = DriveApp.getFolderById(record[1])
    let sp = fldr.createFolder("Student Programs")
    studentPrgmFldr.push([sp.getId()])
  })
  
  folderList.getRange(2,3,studentPrgmFldr.length,1).setValues(studentPrgmFldr)
}



