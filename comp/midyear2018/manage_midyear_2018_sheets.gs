function manage_midyear_2018_sheets() 
{
  /**
   * NOTE: You cannot allow access for IMPORTRANGE through app script. After spreadsheet creation,
   * I will manually access each spreadsheet and allow access.
   *
   * Abbreviations used (case-insensitive):
   *    eib -- Enterprise Interface Builder
   *    cidx -- Column Index (column in array)
   *    cont -- Continuation
   *    cpy - Copy
   *    det -- Details
   **/

  // Confirm user wants to run script
  // var ui = SpreadsheetApp.getUi();
  // var response = ui.alert("Please check cells C1 and C2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
  // if (response !== ui.Button.OK) {
  //  return;
  // }

  // Spreadsheet ID to Spawn Folder ID mapping
  var mapping = {"1J2F8DOLjs3hDahvGjGG5NTZP0MRkCoB0SsvRXZT0VEI" : "0B8RZqzfVtu2lcWdadkNyTEVzc00", // Allie Kline
                 "1NdIKZub80_20RyIwEST4PQVkWN3XYcAHMlwsRvfWzao" : "0B8RZqzfVtu2lSURRbUFaZ2VCMlU", // Atte Lahtiranta
                 "1Evz483b11WfWQQQLotFhQsC565kPW3mVHweA6Rj5K8c" : "0B1f8ZpGaVGpdU0JKMTIxcDkxWEU", // Bob Toohey
                 "1kO49XI5KsEtQtRGGH2WWhp6MCKcazVexcSax_bpXIH0" : "0B8RZqzfVtu2lNGYtZ3dveTl5dGs", // Holly Hess Groos
                 "1LjHOeTP9955OisrH-SyMExzUQSO2OYSzomziMxixTK4" : "0B8RZqzfVtu2lTUEyajB3UEIwNkU", // Jeffrey Bonforte
                 "1rFycbs2N3mRW0r8YpPaSAufLbYoUUBuWhw416rlJbn8" : "0B8RZqzfVtu2lM3JaeGdZanh2aHM", // John DeVine
                 "1GoWFNJ7HKV6Soj4LNGi2LjuiL3Uhcixt6mknc71RWzg" : "0B8RZqzfVtu2lelhwYzI0WHIzcXc", // Julie Jacobs
                 "1qJOC3zJqi-Kat0ZT1uQe8TOYb5po4mLyOS-oYFyFNVk" : "0B8RZqzfVtu2lVHBKM2tOSFBfdWc", // Mark Roszkowski
                 "1XStLZ_Nbo7Xgb-EzBUWYmQxuM1SYuX2edeXsSz6eb6w" : "0B8RZqzfVtu2lTVZlS3VmbkpkRjA", // Ralf Jacob
                 "170UtoblflH8534UygPYrU3xTqtbG5Afznwyqp3wnC74" : "0B8RZqzfVtu2lS0MyV1h4RER2b1k", // Simon Khalaf
                 "1CpSWOfHC-IJXEqGQnQnVestO5aEPZ1FEjV80unFYslQ" : "0B8RZqzfVtu2lVFNkY3dJVmFTb0k", // Tim Mahlman
                 "1Mg_zyaDCBmB4iOWYHPNmVd-wnaWZ7w1e7ID6Puwo_jU" : "0B8RZqzfVtu2lZzlvR0JjQ1VRWVE", // Timothy Lemmon
                 "1XR4LrU5ZmtvHV7T1H2_CCVENpzAdXR2JQPRKV1vbC88" : "0B1f8ZpGaVGpdS3BZSV9uYkIzTHc"} // VP file

  // Promotion Tracker Name to Importrange Key mapping
  var trackername_to_importkey = 
  {
    "Alex Wallace - 2018 Mid Year Promotions Tracker" : "d23792434bda017203935434b94d33f2",
    "Atte Lahtiranta - 2018 Mid Year Promotions Tracker" : "d23792434bda0141e4519e2eb94d80ed",
    "Brian Silver - 2018 Mid Year Promotions Tracker" : "d23792434bda016248bb7b2db94d84ec",
    "Bob Toohey - 2018 Mid Year Promotions Tracker" : "d23792434bda010fc5b4691db94d0adf",
    "Dave McDowell - 2018 Mid Year Promotions Tracker" : "d23792434bda01ba7a41391ab94d3ddc",
    "Geoff Reiss - 2018 Mid Year Promotions Tracker" : "d23792434bda01a0f6a57521b94da3e2",
    "Guru Gowrappan - 2018 Mid Year Promotions Tracker" : "04b33945885801cd8d9b6da4f5a065bb",
    "Jared Grusd - 2018 Mid Year Promotions Tracker" : "d23792434bda01e23314a515b94d4ad8",
    "Jeff Bonforte - 2018 Mid Year Promotions Tracker" : "d23792434bda012592125325b94d5ee5",
    "Jeff D'Onofrio - 2018 Mid Year Promotions Tracker" : "d23792434bda01fb84692d30b94d94ee",
    "Jeff Lucas - 2018 Mid Year Promotions Tracker" : "3ba032e998ed01e00f10a848a9a45a95",
    "Jen Vescio - 2018 Mid Year Promotions Tracker" : "9446ef399a4801c8e87b9b7f3978b7bc",
    "Joanna Lambert - 2018 Mid Year Promotions Tracker" : "ea87f606fdcf01b9388bad82139cc43c",
    "Julie Jacobs - 2018 Mid Year Promotions Tracker" : "d23792434bda0191d9c77a12b94db9d5",
    "Kelly Hirano - 2018 Mid Year Promotions Tracker" : "d23792434bda01cdfac5a021b94dcae2",
    "Mark Roszkowski - 2018 Mid Year Promotions Tracker" : "d23792434bda01e657f3801ab94d7cdc",
    "Natalie Ravitz - 2018 Mid Year Promotions Tracker" : "d23792434bda01bdf6a49218b94dc6da",
    "Rohit Chandra - 2018 Mid Year Promotions Tracker" : "d23792434bda01c534757b24b94db3e4",
    "Rose Tsou - 2018 Mid Year Promotions Tracker" : "d23792434bda01d71e9c0923b94d8ae3",
    "Stuart Flint - 2018 Mid Year Promotions Tracker" : "d23792434bda01e8dec2431cb94d08de",
    "Tenni Theurer - 2018 Mid Year Promotions Tracker" : "d23792434bda012e8aa60626b94d00e6",
    "Tim Lemmon - 2018 Mid Year Promotions Tracker" : "d23792434bda014edffa7016b94df2d8",
    "Tim Mahlman - 2018 Mid Year Promotions Tracker" : "d23792434bda01b9fb51311bb94d18dd",
    "Vanessa Wittman - 2018 Mid Year Promotions Tracker" : "d23792434bda01c7662f741db94d13df",
    "GForce - 2018 Mid Year Promotions Tracker" : "99146c88756001d861a460dda9b0562e"
  }

  var TML_SSID_L2_L3_FILE = "1MKJ5DJv7sIcIe1GyqzhmdAgj53qZOA6iYSsrjHjZgBI";
  var SSID_PROMO_RESPONSES = "1H1qmum4XB4H-pz0W4RqFUghAGza--kdhQ1CUy5PBR6A";
  var FOLDER_ID_L2_L3_TRACKERS = "19c0UebDUC-aZlsQRZOL_0ywM1RrClKhK";
  var SHN_SR_MRG_AND_BELOW = "Sr Mgr and below";
  var IR_RANGE_SR_MGR_AND_BELOW = "\"" + SHN_SR_MRG_AND_BELOW + "!A:L\"";
  var IR_KEY_COL_INDEX_SR_MGR_AND_BELOW = "Col13";
  var SHN_DIR_AND_ABOVE = "Director and above";
  var IR_RANGE_DIR_AND_ABOVE = "\"" + SHN_DIR_AND_ABOVE + "!A:M\"";
  var IR_KEY_COL_INDEX_DIR_AND_ABOVE = "Col14";

  // Google file ids
  var ss = SpreadsheetApp.getActive();
  var FOLDER_L2_L3_TRACKERS = DriveApp.getFolderById(FOLDER_ID_L2_L3_TRACKERS);

  var L2_FILE_SSID = ss.getId();
  var EMPLOYEES_SHN = "Comp Review - EE Data";
  var SPAWN_L3_FILE_TAB_SHN = "Spawn L3 File";
  var SPAWN_FOLDER_ID = mapping[L2_FILE_SSID];
  var TML_SPAWN_FILE_SSID = "1KsmcRf7P1GhkS4wardO-QvNdkTgbGZmZNGsRJMyU-IA";
  var TML_SPAWN_FILE_SHN = "Sheet1";

  function create_L2_L3_files()
  {
    var sheet_lookuptable = ss.getSheetByName(SHN_SR_MRG_AND_BELOW);

    NUM_L2_L3_FILES = 25;
    for (var i = 0; i < NUM_L2_L3_FILES; i++)
    {
      trackername = Object.keys(trackername_to_importkey)[i];
      ir_key = trackername_to_importkey[trackername];
      var new_ssid = create_L2_L3_file(trackername, ir_key);
      sheet_lookuptable.getRange(i+2,4,1,1).setValue(new_ssid);
    }
  }

  function create_L2_L3_file(filename, L2_L3_ir_key)
  {
    var file_tml_cpy = DriveApp.getFileById(TML_SSID_L2_L3_FILE).makeCopy(filename, FOLDER_L2_L3_TRACKERS);
    var sr_mgr_below = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(SHN_SR_MRG_AND_BELOW);
    var dir_above = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(SHN_DIR_AND_ABOVE);
    sr_mgr_below.getRange(1,1,1,1).setFormula("=QUERY(IMPORTRANGE(\"" + SSID_PROMO_RESPONSES + "\"," + IR_RANGE_SR_MGR_AND_BELOW + "),\"select * where " + IR_KEY_COL_INDEX_SR_MGR_AND_BELOW + "='" + L2_L3_ir_key + "'\")");
    dir_above.getRange(1,1,1,1).setFormula("=QUERY(IMPORTRANGE(\"" + SSID_PROMO_RESPONSES + "\"," + IR_RANGE_DIR_AND_ABOVE + "),\"select * where " + IR_KEY_COL_INDEX_DIR_AND_ABOVE + "='" + L2_L3_ir_key + "'\")");
    SpreadsheetApp.flush();
    return(file_tml_cpy.getId());
  }

  /**
   * Extract two columns from an array and convert to a dictionary
   *
   * @param a - array
   * @param key_index - array index for key field
   * @param val_index - array index for value field
   *
   * @return dictionary using key/value columns from given array
   **/
  function create_dict_from_array(a, key_index, val_index)
  {
    var NUM_ROWS = a[0].length;
    var d = {};
    for (var i = 0; i < NUM_ROWS; i++)
    {
      d.push(a[key_index][i], a[val_index][i]);
    }
    return(d);
  }
   
  create_L2_L3_files();
  
}