function create_term_notices()
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
   *    mth -- Month(s)
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

  // Google file ids
  var TERM_NOTICE_FOLDER_ID = "0B2QuBirnXYjxbGtseGZVbWhWWG8"; // Folder: california_change_of_status_documents
  var TERM_NOTICE_NON_TRANSITION_TEMPLATE_ID = "1EQa0S3uXW2Dtdx6k9YbvKcdO5DX6myiQgxiQYfCj6mA"; // File: template_term_notice_nontransition
  var TERM_NOTICE_TRANSITION_TEMPLATE_ID = "1eV4o0RM4m1LdgPHbzlYCFdSMWUOrZn1nUevDJ1OToDY"; // File: template_term_notice_transition
  var RIFS_SSID = "1s-fOV7IZ4ow-N6GTpqtiXg9j9VJR7nvQ3xMkTRvXsqw"; // File: Impacted Yahoos
  var RIFS_SHN = "create_docs"; // Sheet containing RIF'd employees to create docs for

  // Set folder where California Change of Status documents will be created
  var folder = DriveApp.getFolderById(TERM_NOTICE_FOLDER_ID);

  // Load sheet with impacted employees -- may or may not include non-eligible for CIC
  var ees = SpreadsheetApp.openById(RIFS_SSID).getSheetByName(RIFS_SHN);

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
  var EEID_CIDX = 1 - 1;
  var LEGAL_FIRST_NAME_CIDX = 3 - 1;
  var LEGAL_LAST_NAME_CIDX = 4 - 1;
  var LAST_DAY_OF_WORK_CIDX = 6 - 1;
  var SEPARATION_DATE_CIDX = 7 - 1;
  var COBRA_MTHS_CIDX = 10 - 1;
  var SALARY_CONT_MTHS_CIDX = 11 - 1;
  var SEVERANCE_CONT_MTHS_CIDX = 12 - 1; // Salary continuation minus notice period
  var TRANS_BONUS_AMT_CIDX = 13 - 1;
  var OATH_L2_CIDX = 15 - 1;
  var ADDRESS_LINE_1_CIDX = 16 - 1;
  var ADDRESS_LINE_2_CIDX = 17 - 1;
  var ADDRESS_LINE_3_CIDX = 18 - 1;
  var USA_STATE_ISO_CODE_CIDX = 20 - 1;
  var ADEA_FLAG_CIDX = 26 - 1;
  var CIC_ELIGIBILITY_FLAG_CIDX = 27 - 1;
  
  function mail_merge() 
  {
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Only continue for CIC eligible employees
      var is_eligible_for_cic = curr[CIC_ELIGIBILITY_FLAG_CIDX];
      if (is_eligible_for_cic) 
      {
        // Extract required fields
        var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMMM d, yyyy");
        var eeid = curr[EEID_CIDX];
        var legal_first_name = curr[LEGAL_FIRST_NAME_CIDX];
        var legal_last_name = curr[LEGAL_LAST_NAME_CIDX];
        var full_legal_name = legal_first_name + " " + legal_last_name;
        var last_day_of_work = Utilities.formatDate(curr[LAST_DAY_OF_WORK_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy"); 
        var separation_date = Utilities.formatDate(curr[SEPARATION_DATE_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy");
        var cobra_mths = curr[COBRA_MTHS_CIDX];
        var salary_cont_mths = curr[SALARY_CONT_MTHS_CIDX];
        var severance_cont_mths = curr[SEVERANCE_CONT_MTHS_CIDX];
        var trans_bonus_amt = curr[TRANS_BONUS_AMT_CIDX];
        var is_ee_with_transition = (trans_bonus_amt) ? true : false;
        var oath_L2 = curr[OATH_L2_CIDX];
        var address_line_1 = curr[ADDRESS_LINE_1_CIDX];
        var address_line_2 = curr[ADDRESS_LINE_2_CIDX];
        var address_line_3 = curr[ADDRESS_LINE_3_CIDX];        
        var usa_state = curr[USA_STATE_ISO_CODE_CIDX];
        var adea_flag = (curr[ADEA_FLAG_CIDX]) ? "Over40" : "Under40";

        // Copy the template
        var filename = "TNRC - " + adea_flag + " - " + oath_L2 + " - " + usa_state + " - " + full_legal_name + " (" + eeid + ")";
        var TERM_NOTICE_TEMPLATE_ID = (is_ee_with_transition) ? TERM_NOTICE_TRANSITION_TEMPLATE_ID : TERM_NOTICE_NON_TRANSITION_TEMPLATE_ID;
        var file_new_ee_doc = DriveApp.getFileById(TERM_NOTICE_TEMPLATE_ID).makeCopy(filename, folder);
      
        // Fill-in copy with employee details
        var doc_new_ee_doc = DocumentApp.openById(file_new_ee_doc.getId());
        var body = doc_new_ee_doc.getBody();
        body.replaceText("<<today>>", today);
        body.replaceText("<<full_legal_name>>", full_legal_name);
        body.replaceText("<<address_line_1>>", address_line_1); //NEEDTOUPDATE
        if (address_line_2)
        {
          body.replaceText("<<address_line_2>>", address_line_2); //NEEDTOUPDATE
          body.replaceText("<<address_line_3>>", address_line_3);
        } 
        else
        {
          body.replaceText("<<address_line_2>>", address_line_3); //NEEDTOUPDATE
          body.replaceText("<<address_line_3>>", "");
          // var rangeElement = body.findText("<<address_line_2>>");
          // var startOffset = rangeElement.getStartOffset();
          // var endOffset = rangeElement.getEndOffsetInclusive();
          // rangeElement.getElement().asText().deleteText(startOffset, endOffset);
        }
        body.replaceText("<<address_line_3>>", address_line_3); //NEEDTOUPDATE
        body.replaceText("<<legal_first_name>>", legal_first_name);
        body.replaceText("<<last_day_of_work>>", last_day_of_work);
        body.replaceText("<<separation_date>>", separation_date);
        if (is_ee_with_transition) { body.replaceText("<<transition_bonus_amt>>", trans_bonus_amt); }
        body.replaceText("<<continuation_months_minus_notice>>", severance_cont_mths);
        body.replaceText("<<continuation_months_plus_notice>>", salary_cont_mths);
        body.replaceText("<<cobra_months>>", cobra_mths);
        doc_new_ee_doc.saveAndClose();
      
        // Save memo as pdf and delete Google Doc version
        var pdf_version = folder.createFile(file_new_ee_doc.getAs("application/pdf"));
        pdf_version.setName(filename);
        file_new_ee_doc.setTrashed(true);
      }
    }
  }
  mail_merge();
}