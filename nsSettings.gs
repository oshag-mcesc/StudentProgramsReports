 /**
 * @namespace nsSettings
 * @description A namespace containing settings for various tabs.
 */
const nsSettings = (() => {
  let ns = {};

  /**
   * @memberof nsSettings
   * @description Settings for tabs.
   */
  ns.tabs = {
    /**
     * @memberof nsSettings.tabs
     * @description Settings for K-12.
     */
    k12: {
      AugSep: 'Aug-Sep K-12',
      Oct: 'Oct K-12',
      Nov: 'Nov K-12',
      Dec: 'Dec K-12',
      Jan: 'Jan K-12',
      Feb: 'Feb K-12',
      Mar: 'Mar K-12',
      Apr: 'Apr K-12',
      May: 'May/Final K-12'
    },

    /**
     * @memberof nsSettings.tabs
     * @description Settings for PS tabs.
     */
    ps: {
      AugSep: 'Aug-Sep PS',
      Oct: 'Oct PS',
      Nov: 'Nov PS',
      Dec: 'Dec PS',
      Jan: 'Jan PS',
      Feb: 'Feb PS',
      Mar: 'Mar PS',
      Apr: 'Apr PS',
      May: 'May/Final PS'
    },
    studInfo : {
      name:"Student_Info",
      headers : ["SSID","Name","Days"]
    },
    DistrictProgramLists:{
      name:"District & Program Lists"
    },
    K12DistrictProgramSearch:{
      name: "K-12 District/Program Search",
      range: "A3:Q",
      checkRange:"G1"
    },
    programFldrIds:{
      name:"foldrIds",
      range:"A1"
    }
  }

  ns.extractInfo = {
    range : "A1",
    colsKeep:[0,1,19]
  }

  ns.saveFolderID = "13L0urjlK0b2YfthCIPEtQZ6AFDQRbmG7"

  return ns;
})();