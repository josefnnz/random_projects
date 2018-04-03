function create_change_sup_org_eib() 
{
   // Confirm user wants to run script
   var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Please check cells B1 and B2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
   if (response !== ui.Button.OK) {
    return;
   }

  // Google file ids
  var MAIN_FOLDER_ID = "1biWufWRpeZFVExDOX6knrOcTcprFmHS-"; // Folder: 
  var DATA_SSID = "1lv3TDQnTgx5XgWOWJYavgodLIl-WXSU8awrzn7pNMuQ"; // File: 
  var DATA_SHN = "GENERATE_EIB_FOR_SUP_ORG_CHANGES_HERE"; // Sheet with eeid, transition bonus amt, pmt amts, pay dates, etc.
  var EIB_TML_SSID = "1hkhFGT5aphqbKJzZdXVtLWqjDFBsnlo4XUafLau3IzQ"; // File: 
  var EIB_TML_SHN = "Change Job"; // 

  // Set folder where EIB will be created (also holds payment details ss and eib template)
  var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);

  // Load sheet with impacted employees -- may or may not include ees non-eligible for salary continuation
  var ees = SpreadsheetApp.openById(DATA_SSID).getSheetByName(DATA_SHN);

  // Identify specific first and last rows and columns to extract
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetValues(1, 2, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetValues(2, 2, 1, 1);
  var FIRST_COL_EXTRACTED = 1 * ees.getSheetValues(3, 2, 1, 1);
  var LAST_COL_EXTRACTED = 1 * ees.getSheetValues(4, 2, 1, 1);

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = LAST_COL_EXTRACTED - FIRST_COL_EXTRACTED + 1;
    
  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, FIRST_COL_EXTRACTED, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var WORKER_EEID_CIDX = 1 - 1;
  var WORKER_WID_CIDX = 2 - 1;
  var WORKER_PREFERRED_NAME_CIDX = 3 - 1;
  var WORKER_TYPE_CIDX = 4 - 1;
  var WORKER_SUB_TYPE_CIDX = 5 - 1;
  var WORKER_ACTIVE_STATUS_CIDX = 6 - 1;
  var SUP_ORG_NAME_CIDX = 7 - 1;
  var SUP_ORG_INHERITED_FLG_CIDX = 8 - 1;
  var SUP_ORG_WID_CIDX = 9 - 1;
  var SUP_ORG_MGR_DEFAULT_ORG_CIDX = 10 - 1;
  var SUP_ORG_MGR_DEFAULT_ORG_INHERITED_CIDX = 11 - 1;
  var SUP_ORG_MGR_DEFAULT_ORG_WID_CIDX = 12 - 1;

  // EIB Constants
  var NUM_COLS_IN_EIB = 52;

  // Constants
  var TODAYS_DATE = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  function create_full_eib()
  {
    // Create array of payments to fill-in EIB
    var eib_array = create_eib_array();
    var NUM_ROWS_TO_WRITE = eib_array.length;
    var NUM_COLS_TO_WRITE = eib_array[0].length;

    // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
    var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
    var filename = "Change_Supervisory_Organization - " + datetimestamp;

    // Make copy of EIB template in folder, open copy, and write new values
    var file_tml_cpy = DriveApp.getFileById(EIB_TML_SSID).makeCopy(filename, folder);
    var sheet_new_eib = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(EIB_TML_SHN);
    sheet_new_eib.getRange(6, 1, NUM_ROWS_TO_WRITE, NUM_COLS_TO_WRITE).setValues(eib_array);
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
    ees.getRange(5, 2, 1, 1).setValue(datetimestamp);
    ees.getRange(6, 2, 1, 1).setValue("https://drive.google.com/drive/folders/"+MAIN_FOLDER_ID);
    ees.getRange(7, 2, 1, 1).setValue("https://drive.google.com/file/d/"+excel_new_eib.getId()+"/view");
  }
  
  function create_eib_array() 
  {
    // Create empty 2D array to hold payments
    var eib_array = [];
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Get required fields
      var sskey = row + 1;
      var worker_wid = curr[WORKER_WID_CIDX];
      var effective_date = TODAYS_DATE;
      var change_job_reason = "Change_Supervisory_Organization";
      var sup_org_wid = curr[SUP_ORG_MGR_DEFAULT_ORG_WID_CIDX];

      eib_array.push(create_eib_row(sskey, worker_wid, effective_date, change_job_reason, sup_org_wid)); 
    }
    return eib_array;
  }
   
  /**
   * Fill-in EIB payment row with required fields
   *
   * @param sskey -- Spreadsheet Key
   * @param worker_wid -- WID for Employee or Contingent Worker
   * @param effective_date -- Effective date of supervisory org change
   * @param change_job_reason -- Change_Job_Subcategory_ID for Change Job transactions
   * @param sup_org_wid -- WID of new supervisory org
   *
   * @return formatted array to fill EIB payment row with given details
   **/
  function create_eib_row(sskey, worker_wid, effective_date, change_job_reason, sup_org_wid)
  {
    // https://stackoverflow.com/questions/1295584/most-efficient-way-to-create-a-zero-filled-javascript-array
    var row = Array.apply(null, Array(NUM_COLS_IN_EIB)).map(String.prototype.valueOf, "");
    row[1] = sskey;
    row[2] = worker_wid;
    row[4] = effective_date;
    row[5] = change_job_reason;
    row[6] = sup_org_wid;
    return row;
  }

  /**
   * Add necessary leading zeros to make number given number of digits long
   *
   * @param str_num -- number cast as a string
   * @param digits -- desired length of number formatted string
   *
   * @return number as string with desired number of digits
   **/
  function add_leading_zeros(str_num, digits)
  {
    return str_num.length < digits ? add_leading_zeros("0" + str_num, digits) : str_num;
  }
    
  create_full_eib();
}