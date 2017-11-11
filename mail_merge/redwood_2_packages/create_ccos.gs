function create_ccos()
{
  // Constants
  var EMPTY_STRING = "";
  var CIC = "CIC"
  var OATH = "Oath"
  var TRANS = "Trans"
  var NON_TRANS = "NonTrans"
  var WARN = "WARN"
  var NON_WARN = "NonWARN"
  var NON_CA = "NonCA"

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var SEPARATION_AGREEMENT_URL_CIDX = 1 - 1;
  var CALIFORNIA_CHANGE_OF_STATUS_URL_CIDX = 2 - 1;
  var EEID_CIDX = 3 - 1;
  var FULL_LEGAL_NAME_CIDX = 4 - 1;
  var LEGAL_FIRST_NAME_CIDX = 5 - 1;
  var HOME_ADDRESS_LINE_1_CIDX = 6 - 1;
  var HOME_ADDRESS_LINE_2_CIDX = 7 - 1;
  var SSN_CIDX = 8 - 1;
  var CALIFORNIA_FLAG_CIDX = 9 - 1;
  var ADEA_FLAG_CIDX = 10 - 1;
  var L2_ORGNAME_CIDX = 11 - 1;
  var TRANSITION_LENGTH_DAYS_CIDX = 12 - 1;
  var NOTIFICATION_DATE_CIDX = 13 - 1;
  var LAST_DAY_OF_WORK_CIDX = 14 - 1;
  var SEPARATION_DATE_CIDX = 15 - 1;
  var SEVERANCE_PLAN_CIDX = 16 - 1;
  var TRANSITION_FLAG_CIDX = 17 - 1;
  var WARN_FLAG_CIDX = 18 - 1;
  var SEPARATION_AGREEMENT_TEMPLATE_CIDX = 19 - 1;
  var DATE_OF_SEPARATION_AGREEMENT_CIDX = 20 - 1;
  var OATH_PLAN_BASE_COMP_PAYOUT_AMOUNT_CIDX = 21 - 1;
  var OATH_PLAN_NUM_WEEKS_OF_BASE_COMP_PAYOUT_CIDX = 22 - 1;
  var OATH_PLAN_NUM_MONTHS_OF_COBRA_CIDX = 23 - 1;
  var CIC_PLAN_SALARY_CONTINUATION_MONTHS_INCLUDING_NOTICE_PERIOD_CIDX = 24 - 1;
  var CIC_PLAN_SALARY_CONTINUATION_MONTHS_MINUS_NOTICE_PERIOD_CIDX = 25 - 1;
  var CIC_PLAN_NUM_MONTHS_OF_COBRA_CIDX = 26 - 1;
  var TRANSITION_BONUS_AMOUNT_CIDX = 27 - 1;
  var OAB_OR_SIP_CIDX = 28 - 1;
  var RETENTION_FLAG_CIDX = 29 - 1;
  var L2_CIDX = 30 - 1;
  var WORK_LOCATION_CIDX = 31 - 1;

  // Google file ids
  var CALIFORNIA_CHANGE_OF_STATUS_FOLDER_ID = "14OCM7eMv9NNuMlcIibk4cY3LCpBJNe-i"; // Folder: separation_agreements
  var CALIFORNIA_CHANGE_OF_STATUS_TMPL_ID = "14nQr1kWLkMPeRjpoHgKAjZtFendGpFacaP6M13xS9JQ";
  var RIFS_SSID = "1YaUHm_G5O72Twd9EfApfUwAAcLVPWhDRgFGi48OLLZ4"; // File: Project R2 - USA Calculations and Agreement Generator
  var RIFS_SHN = "Calcs";

  // Set folder where Separation Agreements will be created
  var folder = DriveApp.getFolderById(CALIFORNIA_CHANGE_OF_STATUS_FOLDER_ID);

  // Load sheet with impacted employees
  var ees = SpreadsheetApp.openById(RIFS_SSID).getSheetByName(RIFS_SHN);

  // Starting and Ending Rows of Table
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetValues(1, 2, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetValues(2, 2, 1, 1);

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 31; // Columns A - AC

  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  function mail_merge()
  {
    // Confirm user wants to run script
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Please check cells B1 and B2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) 
    {
     return;
    }

    for (var row = 0; row < NUM_EES; row++)
    {
      // Extract current employee
      var curr = values_ees[row];

      if (curr[CALIFORNIA_FLAG_CIDX] == NON_CA)
      {
        ees.getRange(row+FIRST_ROW_EXTRACTED, 2, 1, 1).setValue("NonCA");
        SpreadsheetApp.flush();
        continue;
      }

      // Extract required fields
      var notification_date = curr[NOTIFICATION_DATE_CIDX];
      var full_legal_name = curr[FULL_LEGAL_NAME_CIDX];
      var ssn = curr[SSN_CIDX];
      var separation_date = curr[SEPARATION_DATE_CIDX];
      var eeid = curr[EEID_CIDX];

      var sep_agmt_tmpl = curr[SEPARATION_AGREEMENT_TEMPLATE_CIDX]
      var adea_flag = curr[ADEA_FLAG_CIDX];
      var L2 = curr[L2_CIDX];
      var L2_orgname = curr[L2_ORGNAME_CIDX];
      var california_flag = curr[CALIFORNIA_FLAG_CIDX];
      var work_location = curr[WORK_LOCATION_CIDX];

      // Copy the template
      var filename = work_location + " - " + sep_agmt_tmpl + " - " + adea_flag + " - " + L2 + " - " + california_flag + " - " + full_legal_name + " (" + eeid + ")";
      var file_new_ee_doc = DriveApp.getFileById(CALIFORNIA_CHANGE_OF_STATUS_TMPL_ID).makeCopy(filename, folder);

      // Fil-in copy with employee details
      var doc_new_ee_doc = DocumentApp.openById(file_new_ee_doc.getId());
      var body = doc_new_ee_doc.getBody();

      body.replaceText("<<notification_date>>", notification_date);
      body.replaceText("<<full_legal_name>>", full_legal_name);
      body.replaceText("<<social_security_number>>", ssn);
      body.replaceText("<<separation_date>>", separation_date);
      doc_new_ee_doc.saveAndClose();

      // Save memo as pdf and delete Google Doc version
      var pdf_version = folder.createFile(file_new_ee_doc.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_ee_doc.setTrashed(true);

      // Write unique URL for new file
      ees.getRange(row+FIRST_ROW_EXTRACTED, 2, 1, 1).setValue("https://drive.google.com/a/oath.com/file/d/" + pdf_version.getId() + "/view?usp=sharing");
      SpreadsheetApp.flush();
    }
  }
  mail_merge();
}