function create_request_one_time_payment_eib() 
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
   var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Please check cells C1 and C2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
   if (response !== ui.Button.OK) {
    return;
   }

  // EIB reference ids
  var TRANS_BONUS_PMT_CODE = "OTP_TransBonus"; // Transition Bonus Pmt Plan ID
  var SAL_CONT_PMT_CODE = "OTP_SalContinuation"; // Salary Continuation Pmt Plan ID
  var USD_CURRENCY_ID = "USD"; // Currency ID for US Dollars

  // Google file ids
  var MAIN_FOLDER_ID = "0B2QuBirnXYjxSW9sN2ltWDc2dVk"; // Folder: redwood_salary_continuation
  var PMT_DET_SSID = "1kL9SwTA887XgsvQPMia1OoY_-oWFAYKLPg52ugdambY"; // File: CIC Payroll - US Regular Employees
  var PMT_DET_SHN = "pay_continuation_details"; // Sheet with eeid, transition bonus amt, pmt amts, pay dates, etc.
  var EIB_TML_REQ_ONE_TIME_PMT_SSID = "1zr7OsgYBXRbA9DKwxR3hAdeF4ToY23VMKPtv1ZDrvNA"; // File: Request_One-Time_Payment - GSheet
  var EIB_TML_REQ_ONE_TIME_PMT_SHN = "Request One Time Payment"; // Sheet to input pay details

  // Set folder where EIB will be created (also holds payment details ss and eib template)
  var folder = DriveApp.getFolderById(MAIN_FOLDER_ID);

  // Load sheet with impacted employees -- may or may not include ees non-eligible for salary continuation
  var ees = SpreadsheetApp.openById(PMT_DET_SSID).getSheetByName(PMT_DET_SHN);

  // Identify specific first and last rows to extract
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetValues(1, 3, 1, 1); //NEEDTOUPDATE
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetValues(2, 3, 1, 1); //NEEDTOUPDATE

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 59; // Columns A to BG -- NEEDTOUPDATE
    
  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var EEID_CIDX = 2 - 1; //NEEDTOUPDATE
  var TRANS_FLAG_CIDX = 5 - 1; //NEEDTOUPDATE
  var TRANS_BONUS_AMT_CIDX = 6 - 1; //NEEDTOUPDATE
  var PMT_AMT_CIDX = 13 - 1; //NEEDTOUPDATE
  var NUM_PMTS_CIDX = 10 - 1; //NEEDTOUPDATE
  var FIRST_PAY_DATE_CIDX = 16 - 1; //NEEDTOUPDATE
  var LAST_PAY_DATE_CIDX = 59 - 1; //NEEDTOUPDATE

  function create_full_eib()
  {
    // Create array of payments to fill-in EIB
    var pmts = create_pmts_array();
    var NUM_ROWS_TO_WRITE = pmts.length;
    var NUM_COLS_TO_WRITE = pmts[0].length;

    // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
    var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
    var filename = "Request_One-Time_Payment - " + datetimestamp;

    // Make copy of EIB template in folder, open copy, and write new values
    var file_tml_cpy = DriveApp.getFileById(EIB_TML_REQ_ONE_TIME_PMT_SSID).makeCopy(filename, folder);
    var sheet_new_eib = SpreadsheetApp.openById(file_tml_cpy.getId()).getSheetByName(EIB_TML_REQ_ONE_TIME_PMT_SHN);
    sheet_new_eib.getRange(6, 1, NUM_ROWS_TO_WRITE, NUM_COLS_TO_WRITE).setValues(pmts);
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
    ees.getRange(4, 5, 1, 1).setValue(datetimestamp);
    ees.getRange(5, 5, 1, 1).setValue("https://drive.google.com/file/d/"+excel_new_eib.getId()+"/view");
  }
  
  function create_pmts_array() 
  {
    // Create empty 2D array to hold payments
    var pmts = [];
    var sskey = 1; // Initiate a spreadsheet key value for EIB
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Extract required fields
      var eeid = add_leading_zeros(new String(curr[EEID_CIDX]), 6); // ensure EEID is 6 digits long
      var trans_flag = curr[TRANS_FLAG_CIDX];
      var trans_bonus_amt = curr[TRANS_BONUS_AMT_CIDX];
      var pmt_amt = curr[PMT_AMT_CIDX];
      var num_pmts = curr[NUM_PMTS_CIDX];
      var first_pay_date = curr[FIRST_PAY_DATE_CIDX];

      if (trans_flag === "Y")
      {
        // Add transition bonus payment if applicable -- transition bonus paid on first continued pay date
        pmts.push(create_eib_row(new String(sskey), eeid, Utilities.formatDate(first_pay_date, Session.getScriptTimeZone(), "yyyy-MM-dd"), TRANS_BONUS_PMT_CODE, trans_bonus_amt, USD_CURRENCY_ID));
        sskey++; // Increment spreadsheet key
      }

      // Add salary continuation payments for allotted number of payments
      for (var i = FIRST_PAY_DATE_CIDX; i <= LAST_PAY_DATE_CIDX; i++)
      {
        pay_date = curr[i];
        if (!pay_date)
        {
          break; // Break loop if pay date is NULL -- already covered all required payments
        }
        pay_date = Utilities.formatDate(pay_date, Session.getScriptTimeZone(), "yyyy-MM-dd");
        pmts.push(create_eib_row(new String(sskey), eeid, pay_date, SAL_CONT_PMT_CODE, pmt_amt, USD_CURRENCY_ID));
        sskey++; // Increment spreadsheet key
      }    
    }
    return pmts;
  }
   
  /**
   * Fill-in EIB payment row with required fields
   *
   * @param sskey -- Spreadsheet Key
   * @param eeid -- Employee ID in 6 digit text format (i.e. has leading zeros)
   * @param effdate -- Effective Date of payment
   * @param pmtcode -- Workday One-Time_Payment_Plan_ID value. Separate codes for continuation vs. transition bonus payments
   * @param amt -- Payment amount
   * @param currency -- Payment currency
   *
   * @return formatted array to fill EIB payment row with given details
   **/
  function create_eib_row(sskey, eeid, effdate, pmtcode, amt, currency)
  {
    // All EIB columns: 
    // Fields, Spreadsheet Key, Employee position, Effective Date, Employee Visibility Date, Reason, 
    // One Time Payment Plan, Amount, Percent, Currency, Comment, Do Not Pay
    return ["", sskey, eeid, "", effdate, "", "", pmtcode, amt, "", currency, "", ""];
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
    
  //var prompt_text = "Are the first and last row employee names correct? If not, please fix the values in cells C1 and C2. \n\n"
  //                  + "";
  //var ui = SpreadsheetApp.getUi();
  //var result = ui.prompt(prompt_text, ui.ButtonSet.YES_NO_CANCEL);
    
  create_full_eib();
}