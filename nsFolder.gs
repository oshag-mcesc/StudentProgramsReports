/**
 * @file This file contains functions for creating folders in Google Drive.
 */

/**
 * @namespace nsFolders
 * @description This namespace groups functions related to folder creation.
 */
const nsFolders = (()=>{
  /**
   * Creates folders in Google Drive based on a list of school names.
   * 
   * @returns {boolean} True if folders were created successfully, false otherwise.
   */
  const createFolders_ = ()=>{
    // TODO: Implement a more robust way to retrieve the tab name dynamically.
    // This could involve user input, configuration files, etc.
    const {ss, tab} = getSheetTab(nsSettings.tabs.DistrictProgramLists.name)
    
    //FOR TESTING hard coded to get to Column H
    const schools = tab.getRange("H1").getDataRegion().getValues().flat()
    //Get rid of header row
    schools.shift()
    //const schools = tab.getRange(2,8,tab.getLastRow()-1,1).getValues().flat()
    const folderInfo = [] // Array to store folder information

    // TODO: Consider implementing logic to handle folder hierarchy or user selection.
    const parentFolder = DriveApp.getFolderById(nsSettings.saveFolderID)

    schools.forEach(school => {
      const newFolder = parentFolder.createFolder(school)
      folderInfo.push([newFolder.getId(), newFolder.getName()])
    });

    // Update the spreadsheet with folder information
    // For now put it way out in ColumnmJ
    tab.getRange(2, 10, folderInfo.length, 2).setValues(folderInfo)
    return true
  }

  return{
    createFolders: createFolders_,
  }
})()


/**
 * Retrieves the spreadsheet and the specified sheet by name.
 *
 * @param {string} tabName The name of the sheet to retrieve.
 * @returns {object} An object containing the spreadsheet and sheet objects.
 */
const getSheetTab = (tabName)=>{
  const info = {}
  info.ss = SpreadsheetApp.getActiveSpreadsheet()
  info.tab   = info.ss.getSheetByName(tabName)

  return info
}


/**
 * Runs the createFolders function from the nsFolders namespace.
 */
const runCreateFolders = ()=> {
  const result = nsFolders.createFolders()
  console.log(result); // Log the result (true/false)
}