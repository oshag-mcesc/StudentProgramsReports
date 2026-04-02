// In runReports.gs - UPDATE THIS FUNCTION

const runReports = () => {
  const { ss, tab } = getSheetTab(nsSettings.tabs.K12DistrictProgramSearch.name)
  const range = tab.getRange(nsSettings.tabs.K12DistrictProgramSearch.range)

  // Get folder and district/program info
  let folderInfo = ss.getSheetByName(nsSettings.tabs.programFldrIds.name)
  let dpl = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name)

  const [schoolHeader, ...criSchools] = folderInfo.getRange(nsSettings.tabs.programFldrIds.range).getDataRegion().getValues()
  const criProgramsTemp = dpl.getRange("E2:E").getValues().flat()
  const criPrograms = criProgramsTemp.filter(value => value !== "")

  criSchools.forEach(school => {
    criPrograms.forEach(program => {
      tab.getRange("C1").setValue(school[0])
      tab.getRange("C2").setValue(program)
      if (range.getFilter()) { range.getFilter().remove() }
      let filter = range.createFilter()
      let cri1 = SpreadsheetApp.newFilterCriteria()
        .whenTextContains(school[0])
        .build()

      let cri2 = SpreadsheetApp.newFilterCriteria()
        .whenTextEqualTo(program)
        .build()

      filter
        .setColumnFilterCriteria(7, cri1)
        .setColumnFilterCriteria(4, cri2)

      // Check to see if student count is greater than 0
      let count = tab.getRange(nsSettings.tabs.K12DistrictProgramSearch.checkRange).getValue()
      if (count > 0) {
        let fileName = school[0] + " - " + program + " - " + nsHelpers.getPreviousMonthName()
        
        // UPDATED: Pass the specific sheet instead of relying on active sheet
        exportSheetAsPDF(tab, fileName, school[1])  // ← Changed this line!
      }

    })

  });

  return true
}


const runReportTest = () => {
  // Show a 3-second popup with the title "Status" and the message "Task started".
  SpreadsheetApp.getActiveSpreadsheet().toast('Reports are a reporting!!', 'Started!!', -1);
  let rslt = runReports()
  if (rslt) {
    // Show a 3-second popup with the title "Status" and the message "Task started".
    SpreadsheetApp.getActiveSpreadsheet().toast('Reports have been reported!!', 'DONE', -1);
  }
  console.log(rslt);

}







