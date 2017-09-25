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

  // Google file ids
  var ss = SpreadsheetApp.getActive();

  var L2_FILE_SSID = ss.getId();
  var EMPLOYEES_SHN = "Comp Review - EE Data w/TTC";
  var SPAWN_L3_FILE_TAB_SHN = "Spawn L3 File";
  var SPAWN_FOLDER_ID = mapping[L2_FILE_SSID];
  var TML_SPAWN_FILE_SSID = "1kWeN7K4qOhQL4b_ELTKjrvewIM-OTpwk5wSkebAsX3U";
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
  var L2_CIDX = 4 - 1;
  var L3_CIDX = 5 - 1;

  // Identify location of columns with formulas
  var FINAL_NEW_SALARY_SECTION_START_CIDX = 47;
  var FINAL_NEW_SALARY_SECTION_END_CIDX = 57;
  var FINAL_INCREASE_REASON_CIDX = 71;
  var SALARY_INCREASE_INPUTTED_CIDX = 98;

  // Lengths of formula sections
  var NUM_FINAL_NEW_SALARY_SECTION = FINAL_NEW_SALARY_SECTION_END_CIDX - FINAL_NEW_SALARY_SECTION_START_CIDX + 1;
  var NUM_FINAL_INCREASE_REASON = 1;
  var NUM_SALARY_INCREASE_INPUTTED = 1;

  function create_L3_file()
  {
    // Get L3 to make file for
    var L3_name = sheet_spawn_tab.getRange(1, 3, 1, 1).getValue();

    // Create array of employee data values to fill-in spreadsheet
    var values_ees_under_L3 = values_ees.filter(function(row) { return row[L3_CIDX] == L3_name });
    var NUM_ROWS_TO_WRITE = values_ees_under_L3.length;
    var NUM_COLS_TO_WRITE = values_ees_under_L3[0].length;

    // Create array of formulas to fill-in spreadsheet
    var formulas_final_new_salary_section = ees.getRange(FIRST_ROW_EXTRACTED, FINAL_NEW_SALARY_SECTION_START_CIDX, NUM_ROWS_TO_WRITE, NUM_FINAL_NEW_SALARY_SECTION).getFormulas();
    var formulas_final_increase_reason = ees.getRange(FIRST_ROW_EXTRACTED, FINAL_INCREASE_REASON_CIDX, NUM_ROWS_TO_WRITE, NUM_FINAL_INCREASE_REASON).getFormulas();
    var formulas_salary_increase_inputted = ees.getRange(FIRST_ROW_EXTRACTED, SALARY_INCREASE_INPUTTED_CIDX, NUM_ROWS_TO_WRITE, NUM_SALARY_INCREASE_INPUTTED).getFormulas();

    // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
    var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
    var filename = L3_name + " - " + datetimestamp;

    // Make copy of EIB template in folder, open copy, and write new values
    var file_tml_cpy = DriveApp.getFileById(TML_SPAWN_FILE_SSID).makeCopy(filename, folder);
    var sheet_new_L3_file = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(TML_SPAWN_FILE_SHN);
    sheet_new_L3_file.getRange(5, 1, 1, 1).setValue(L3_name);
    sheet_new_L3_file.getRange(7, 1, NUM_ROWS_TO_WRITE, NUM_COLS_TO_WRITE).setValues(values_ees_under_L3);
    SpreadsheetApp.flush()
    sheet_new_L3_file.getRange(FIRST_ROW_EXTRACTED, FINAL_NEW_SALARY_SECTION_START_CIDX, NUM_ROWS_TO_WRITE, NUM_FINAL_NEW_SALARY_SECTION).setFormulas(formulas_final_new_salary_section);
    SpreadsheetApp.flush()
    sheet_new_L3_file.getRange(FIRST_ROW_EXTRACTED, FINAL_INCREASE_REASON_CIDX, NUM_ROWS_TO_WRITE, NUM_FINAL_INCREASE_REASON).setFormulas(formulas_final_increase_reason);
    SpreadsheetApp.flush()
    sheet_new_L3_file.getRange(FIRST_ROW_EXTRACTED, SALARY_INCREASE_INPUTTED_CIDX, NUM_ROWS_TO_WRITE, NUM_SALARY_INCREASE_INPUTTED).setFormulas(formulas_salary_increase_inputted);
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