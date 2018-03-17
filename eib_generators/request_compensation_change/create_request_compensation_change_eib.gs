function create_request_compensation_change_eib() 
{
  // Array.apply(null, Array(10)).map(String.prototype.valueOf, "")
  // there are 82 columns on the Request Compensation Change EIB
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
   var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Please check cells M1 and M2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
   if (response !== ui.Button.OK) {
    return;
   }

  // Google file ids
  var MAIN_FOLDER_ID = "1TiCrf4J79XQ2Jk264JI1F4uoG-SxuvoY"; // Folder: 
  var DATA_SSID = "1m46zKbNTIi64vTWnLE9jvcA42W7iEBZ6nBHEB6AdQ-4"; // File: 
  var DATA_SHN = "data"; // Sheet with eeid, transition bonus amt, pmt amts, pay dates, etc.
  var EIB_TML_SSID = "1pgA8-BrsSHoPyp44ttPCCYmAGDPcBHy4F2sBrCeEOek"; // File: 
  var EIB_TML_SHN = "Request Compensation Change"; // 

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
  var EEID_CIDX = 1 - 1;
  var EFFECTIVE_DATE_CIDX = 2 - 1;
  var REQ_COMP_CHANGE_REASON_CODE_CIDX = 3 - 1;
  var SALARY_OR_HOURLY_PLAN_CIDX = 4 - 1;
  var SALARY_OR_HOURLY_AMOUNT_CIDX = 5 - 1;
  var LOCAL_CURRENCY_CIDX = 6 - 1;
  var ANNUAL_OR_HOURLY_FREQUENCY_CIDX = 7 - 1;
  var BONUS_PLAN_CIDX = 8 - 1;
  var PERCENT_OR_AMOUNT_BASED_BONUS_PLAN_CIDX = 9 - 1;
  var TARGET_BONUS_PERCENT_CIDX = 10 - 1;
  var TARGET_BONUS_AMOUNT_CIDX = 11 - 1;
  var IS_ON_INDIVIDUAL_TARGET_CIDX = 12 - 1;

  // EIB Constants
  var NUM_COLS_IN_EIB = 82;

  function create_full_eib()
  {
    // Create array of payments to fill-in EIB
    var eib_array = create_eib_array();
    var NUM_ROWS_TO_WRITE = eib_array.length;
    var NUM_COLS_TO_WRITE = eib_array[0].length;

    // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
    var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
    var filename = "Request_Compensation_Change - " + datetimestamp;

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
    ees.getRange(3, 14, 1, 1).setValue(datetimestamp);
    ees.getRange(4, 14, 1, 1).setValue("https://drive.google.com/file/d/"+excel_new_eib.getId()+"/view");
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
      var eeid = add_leading_zeros(new String(curr[EEID_CIDX]), 6); // ensure EEID is 6 digits long
      var eff_date = curr[EFFECTIVE_DATE_CIDX];
      var reason_code = curr[REQ_COMP_CHANGE_REASON_CODE_CIDX];
      var comp_plan = curr[SALARY_OR_HOURLY_PLAN_CIDX];
      var comp_plan_amt = curr[SALARY_OR_HOURLY_AMOUNT_CIDX];
      var currency = curr[LOCAL_CURRENCY_CIDX];
      var freq = curr[ANNUAL_OR_HOURLY_FREQUENCY_CIDX];
      var bonus_plan = curr[BONUS_PLAN_CIDX];
      var pct_or_amt_based = curr[PERCENT_OR_AMOUNT_BASED_BONUS_PLAN_CIDX];
      var tgt_bonus_pct = curr[TARGET_BONUS_PERCENT_CIDX];
      var tgt_bonus_amt = curr[TARGET_BONUS_AMOUNT_CIDX];
      var is_on_individ_tgt = curr[IS_ON_INDIVIDUAL_TARGET_CIDX];
      var individamt = (pct_or_amt_based == "Amount") ? tgt_bonus_amt : "";
      var individpct = (pct_or_amt_based == "Percent" && is_on_individ_tgt == "Y") ? tgt_bonus_pct : "";
      var pctassigned = (pct_or_amt_based == "Percent" && is_on_individ_tgt == "N") ? 1 : "";

      eib_array.push(create_eib_row(sskey, eeid, eff_date, reason_code, comp_plan, comp_plan_amt, currency, freq, bonus_plan, individamt, individpct, pctassigned)); 
    }
    return eib_array;
  }
   
  /**
   * Fill-in EIB payment row with required fields
   *
   * @param sskey -- Spreadsheet Key
   * @param eeid -- Employee ID in 6 digit text format (i.e. has leading zeros)
   * @param effdate -- Effective Date of payment
   * @param reason -- Request Compensation Change reason code
   * @param compplan -- Salary or Hourly plan
   * @param amt -- Annualized Base for Salaried EEs or Hourly Rate for Hourly EEs (values based on 100% FTE)
   * @param currency -- local currency
   * @param freq -- Annual or Hourly frequency
   * @param bonusplan -- bonus plan
   * @param individamt -- individual bonus amount for amount based plans
   * @param individpct -- individual bonus percent for percent based plans
   * @param pctassigned -- percentage of default bonus percent assigned to EE (set to 1 for EEs on target)
   *
   * @return formatted array to fill EIB payment row with given details
   **/
  function create_eib_row(sskey, eeid, effdate, reason, compplan, amt, currency, freq, bonusplan, individamt, individpct, pctassigned)
  {
    // https://stackoverflow.com/questions/1295584/most-efficient-way-to-create-a-zero-filled-javascript-array
    var row = Array.apply(null, Array(NUM_COLS_IN_EIB)).map(String.prototype.valueOf, "");
    row[1] = sskey; // Spreadsheet Key*
    row[2] = eeid; // Employee*
    row[4] = effdate; // Compensation Change Date*
    row[5] = reason; //Reason*
    // Base section
    var is_there_a_base_change = compplan.trim() !== "";
    if (is_there_a_base_change)
    {
      row[11] = "Y"; // Replace
      row[12] = "1"; // Row ID*
      row[13] = compplan; // Pay Plan
      row[14] = amt; // Amount*
      row[17] = currency; // Currency
      row[18] = freq; // Frequency
    }
    // Bonus section
    var is_there_a_bonus_change = bonusplan.trim() !== "";
    if (is_there_a_bonus_change)
    {
      row[44] = "Y"; // Replace
      row[45] = "1"; // Row ID*
      row[46] = bonusplan; // Bonus Plan
      row[47] = individamt; // Individual Target Amount
      row[48] = individpct; // Individual Target Percent
      row[50] = pctassigned; // Percent Assigned
    }
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