function spawn_L3_file() 
{
  /**
   * Abbreviations used (case-insensitive):
   *    eib -- Enterprise Interface Builder
   *    cidx -- Column Index (column in array)
   *    cont -- Continuation
   *    cpy - Copy
   *    det -- Details
   *    dt -- Date
   *    ee -- Employee(s)
   *    flr -- Folder
   *    idx -- Index
   *    pmt -- Payment(s)
   *    req -- Request
   *    sal -- Salary
   *    shn -- Sheet Name
   *    ss -- Spreadsheet
   *    ssid -- Google Spreadsheet ID
   *    tml -- Template
   *    trans -- Transition
   *    val -- Value(s)
   **/

  // Confirm user wants to run script
  // var ui = SpreadsheetApp.getUi();
  // var response = ui.alert("Please check cells C1 and C2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
  // if (response !== ui.Button.OK) {
  //  return;
  // }

  // // Spreadsheet ID to Spawn Folder ID mapping
  // var mapping = {"1RBFl7WbdSqP1jOZboWcq5dmSQP9lBWkoDA6oSDvXAho" : "0B8RZqzfVtu2lcWdadkNyTEVzc00", // TEST TO ALLIE KLINE FOLDER
  //                "0B1f8ZpGaVGpdM1pORHVLSWhHdDQ" : "0B8RZqzfVtu2lcWdadkNyTEVzc00", // Allie Kline
  //                "0B1f8ZpGaVGpdV3RKNDVFN09pY1k" : "0B8RZqzfVtu2lSURRbUFaZ2VCMlU", // Atte Lahtiranta
  //                "0B1f8ZpGaVGpdVzhRaGlWMnVPT3c" : "0B8RZqzfVtu2lNGYtZ3dveTl5dGs", // Holly Hess Groos
  //                "0B1f8ZpGaVGpdNE5zSnN1bWVYam8" : "0B8RZqzfVtu2lTUEyajB3UEIwNkU", // Jeffrey Bonforte
  //                "0B1f8ZpGaVGpdMFhTQ2ctaFg1VzA" : "0B8RZqzfVtu2lM3JaeGdZanh2aHM", // John DeVine
  //                "0B1f8ZpGaVGpdQlRhQWNYaUphZHc" : "0B8RZqzfVtu2lelhwYzI0WHIzcXc", // Julie Jacobs
  //                "0B1f8ZpGaVGpdZ2x6TGRfMDZ1QWc" : "0B8RZqzfVtu2lVHBKM2tOSFBfdWc", // Mark Roszkowski
  //                "0B1f8ZpGaVGpdMVFfaTJqR19wcU0" : "0B8RZqzfVtu2lTVZlS3VmbkpkRjA", // Ralf Jacob
  //                "0B1f8ZpGaVGpdZHJXRTQ4ZnlaWk0" : "0B8RZqzfVtu2lS0MyV1h4RER2b1k", // Simon Khalaf
  //                "0B1f8ZpGaVGpdUi10S1FCd01qaFE" : "0B8RZqzfVtu2lVFNkY3dJVmFTb0k", // Tim Mahlman
  //                "0B1f8ZpGaVGpdLXl5WEI3b3psVWM" : "0B8RZqzfVtu2lZzlvR0JjQ1VRWVE"} // Timothy Lemmon
  var mapping = {"1RBFl7WbdSqP1jOZboWcq5dmSQP9lBWkoDA6oSDvXAho" : "0B8RZqzfVtu2lcWdadkNyTEVzc00"} // TEST TO ALLIE KLINE FOLDER

  // Google file ids
  var ss = SpreadsheetApp.getActive();

  var L2_FILE_SSID = ss.getId();
  var EMPLOYEES_SHN = "Comp Review - EE Data";
  var SPAWN_L3_FILE_TAB_SHN = "SPAWN L3 FILE";
  var SPAWN_FOLDER_ID = mapping[L2_FILE_SSID];
  var TML_SPAWN_FILE_SSID = "1KsmcRf7P1GhkS4wardO-QvNdkTgbGZmZNGsRJMyU-IA";
  var TML_SPAWN_FILE_SHN = "Sheet1";

  // Set folder where spawned file will be created
  var folder = DriveApp.getFolderById(SPAWN_FOLDER_ID);

  // Load generate spawn tab from L2 spreadsheet
  var sheet_spawn_tab = ss.getSheetByName(SPAWN_L3_FILE_TAB_SHN);

  // Load sheet with employees for given L2 spreadsheet
  var ees = ss.getSheetByName(EMPLOYEES_SHN);

  // Identify specific first and last rows to extract
  var FIRST_ROW_EXTRACTED = 7; // First row of table, excluding column headers
  var LAST_ROW_EXTRACTED = ees.getLastRow(); // Dynamically grab last row of table

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = ees.getLastColumn(); // Dynamically grab last column of table
    
  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var L3_CIDX = 5 - 1;

  function create_L3_file()
  {
    // Get L3 to make file for
    var L3_name = sheet_spawn_tab.getRange(1, 3, 1, 1).getValue();

    // Create array of payments to fill-in EIB
    var values_ees_under_L3 = values_ees.filter(function(row) { return row[L3_CIDX] == L3_name });
    var NUM_ROWS_TO_WRITE = values_ees_under_L3.length;
    var NUM_COLS_TO_WRITE = values_ees_under_L3[0].length;

    // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
    var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
    var filename = L3_name + " - " + datetimestamp;

    // Make copy of EIB template in folder, open copy, and write new values
    var file_tml_cpy = DriveApp.getFileById(TML_SPAWN_FILE_SSID).makeCopy(filename, folder);
    var sheet_new_L3_file = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(TML_SPAWN_FILE_SHN);
    sheet_new_L3_file.getRange(5, 1, 1, 1).setValue(L3_name);
    sheet_new_L3_file.getRange(7, 1, NUM_ROWS_TO_WRITE, NUM_COLS_TO_WRITE).setValues(values_ees_under_L3);
    SpreadsheetApp.flush()

    // Save new EIB as Excel file, and delete GSheet version
    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + file_tml_cpy.getId() + "&exportFormat=xlsx";
    var params = {
      method : "get",
      headers : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions : true
    };
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(filename + ".xlsx");
    var excel_new_eib = folder.createFile(blob);
    file_tml_cpy.setTrashed(true);   

    // Write unique URL for new EIB file in spreadsheet for easy reference
    sheet_spawn_tab.getRange(7, 3, 1, 1).setValue("https://drive.google.com/drive/folders/" + SPAWN_FOLDER_ID);
    sheet_spawn_tab.getRange(8, 3, 1, 1).setValue(datetimestamp);
    sheet_spawn_tab.getRange(9, 3, 1, 1).setValue(L3_name);
    sheet_spawn_tab.getRange(10, 3, 1, 1).setValue("https://drive.google.com/file/d/"+excel_new_eib.getId()+"/view");
  }
   
  create_L3_file();
}