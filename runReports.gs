const runReports = ()=> {
  const {ss, tab}=getSheetTab(nsSettings.tabs.K12DistrictProgramSearch.name)
  const range = tab.getRange(nsSettings.tabs.K12DistrictProgramSearch.range)
  
  // TODO: HAS to be a better way here!!
  //HARD CODED for demonstration!!
  let dpl = ss.getSheetByName(nsSettings.tabs.DistrictProgramLists.name)

  const [schoolHeader, ...criSchools] = dpl.getRange("H1").getDataRegion().getValues()
  const criProgramsTemp = dpl.getRange("E2:E").getValues().flat()
  const criPrograms = criProgramsTemp.filter(value => value !=="")

  criSchools.forEach(school => {
    criPrograms.forEach(program =>{
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
        

        //check to see if student count is greater than 0, is so export PDF
        let count = tab.getRange(nsSettings.tabs.K12DistrictProgramSearch.checkRange).getValue()
        if(count >1 ){
          let fileName = school[0] + " - " + program + " - " + nsHelpers.getCurrentMonthName()
          exportCurrentSheetAsPDF(fileName,school[1])
        }
        
    })
    
  });

  return true
}




const runReportTest = ()=> {
  let rslt = runReports()
  console.log(rslt);
  
}