const getIt = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let tab = ss.getSheetByName(nsSettings.tabs.k12.Oct)

  let data = tab.getRange(nsSettings.extractInfo.range).getDataRegion().getValues()
  data.shift()
  let newData = nsHelpers.keepCols(data, nsSettings.extractInfo.colsKeep)

  let rslt = ss.getSheetByName(nsSettings.tabs.studInfo.name)
  rslt.getRange(1, 1, newData.length, newData[0].length).setValues(newData)


}

const updateStudentDays = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  try {
    let newData = []
    const k12tabs = nsSettings.tabs.k12
    for (const tabName in k12tabs) {
      const [headers,...tabData] = ss.getSheetByName(k12tabs[tabName]).getRange(nsSettings.extractInfo.range).getDataRegion().getValues()
      let smallData = nsHelpers.keepCols(tabData, nsSettings.extractInfo.colsKeep)
      newData = newData.concat(smallData)
    }

    newData.unshift(nsSettings.tabs.studInfo.headers)
    let rslt = ss.getSheetByName(nsSettings.tabs.studInfo.name)
    rslt.clearContents()
    rslt.getRange(1, 1, newData.length, newData[0].length).setValues(newData)

  }
  catch (err) {
    console.log(err);

  }
}


const tester3 = () => {
  // Example usage:
  const originalArray = [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
    [11, 12, 13, 14, 15]
  ]
  let rslt = []
  for (let i = 0; i < 2; i++) {
    let arr = nsHelpers.keepCols(originalArray, [i + 1, i + 2])
    console.log(arr);

    console.log(arr[0])

    rslt = rslt.concat(arr)
  }
  console.log(rslt);


}

const datePrac = ()=>{
  const monthName = new Date().toLocaleString('default', { month: 'long' });
console.log(monthName);

}












