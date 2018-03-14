function create_legal_contracts()
{

  // Confirm user wants to run script
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Please check cells B1 and B2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to prevent the script from running.", ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
   return;
  }

  // Google file ids
  var BACKUP_GDOCS_FOLDER_ID = "1pOg0LhYZ404_ov-pKansAsSoypoPhSVO"; // Folder: 
  var PDFS_FOLDER_ID = "1rgl4tEGuwq09UXSHvJTYxxklyDyVHSuj"; // Folder:
  var TMPL_ID = "1Q1gQVLH9T4XtJB_1a64cD6PVHsBXueddaO-zPq2LD9s"; // Template: 
  var SSID = "1hXeyTOPkDN7N1ZjgyUcl1ojW5IafOzc_7mFGPYJ2MRE"; // Spreadsheet: 
  var SHN = "create_contracts"; // Sheet Name: 

  // Get folder
  var backup_gdocs_folder = DriveApp.getFolderById(BACKUP_GDOCS_FOLDER_ID);
  var pdfs_folder = DriveApp.getFolderById(PDFS_FOLDER_ID);

  // Get sheet
  var sheet_ees = SpreadsheetApp.openById(SSID).getSheetByName(SHN);

  // Identify first and last rows to extract
  var FIRST_ROW_EXTRACTED = 1 * sheet_ees.getSheetValues(1, 2, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * sheet_ees.getSheetValues(2, 2, 1, 1);

  // Identify first column to extract
  var FIRST_COL_EXTRACTED = 1;

  // Identify number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 11; // Columns B to K

  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = sheet_ees.getRange(FIRST_ROW_EXTRACTED, FIRST_COL_EXTRACTED, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var CONTRACT_URL_CIDX = 1 - 1;
  var DATE_OF_CONTRACT_CREATION_CIDX = 2 - 1;
  var LEGAL_FIRST_NAME_CIDX = 3 - 1;
  var EFFECTIVE_DATE_OF_CHANGE_CIDX = 4 - 1;
  var NEW_SALARY_CIDX = 5 - 1;
  var OLD_BONUS_PCT_CIDX = 6 - 1;
  var NEW_BONUS_PCT_CIDX = 7 - 1;
  var DATE_TO_RETURN_SIGNED_CONTRACT_CIDX = 8 - 1;
  var ENTITY_NAME_CIDX = 9 - 1;
  var LEGAL_FULL_NAME_CIDX = 10 - 1;
  var EEID_CIDX = 11 - 1;

  function mail_merge() 
  {
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Get required fields
      var date_of_contract_creation = curr[DATE_OF_CONTRACT_CREATION_CIDX];
      var legal_first_name = curr[LEGAL_FIRST_NAME_CIDX];
      var effective_date = curr[EFFECTIVE_DATE_OF_CHANGE_CIDX];
      var new_salary = curr[NEW_SALARY_CIDX];
      var old_bonus_pct = curr[OLD_BONUS_PCT_CIDX];
      var new_bonus_pct = curr[NEW_BONUS_PCT_CIDX];
      var date_to_return_contract = curr[DATE_TO_RETURN_SIGNED_CONTRACT_CIDX];
      var legal_entity = curr[ENTITY_NAME_CIDX];
      var legal_full_name = curr[LEGAL_FULL_NAME_CIDX];
      var eeid = curr[EEID_CIDX];

      // Create filename for statement
      var filename = "Change of Employment Terms for " + legal_full_name + " (" + eeid + ")";

      // Copy statement template gdoc. Open new copy.
      var file_tmpl_copy = DriveApp.getFileById(TMPL_ID).makeCopy(filename, backup_gdocs_folder)
      var doc_tmpl_copy = DocumentApp.openById(file_tmpl_copy.getId());
      var body = doc_tmpl_copy.getBody();

      // Merge data fields in statement, as applicable
      body.replaceText("<<DATE_OF_CONTRACT_CREATION>>", date_of_contract_creation);
      body.replaceText("<<LEGAL_FIRST_NAME>>", legal_first_name);
      body.replaceText("<<EFFECTIVE_DATE_OF_CHANGE>>", effective_date);
      body.replaceText("<<NEW_SALARY>>", new_salary);
      body.replaceText("<<OLD_BONUS_PCT>>", old_bonus_pct);
      body.replaceText("<<NEW_BONUS_PCT>>", new_bonus_pct);
      body.replaceText("<<DATE_TO_RETURN_SIGNED_CONTRACT>>", date_to_return_contract);
      body.replaceText("<<ENTITY_NAME>>", legal_entity);
      body.replaceText("<<LEGAL_FULL_NAME>>", legal_full_name);
      body.replaceText("<<EMPLOYEE_ID>>", eeid);

      // Save completed gdoc into backup folder of gdocs
      doc_tmpl_copy.saveAndClose();

      // Save gdoc as pdf in the corresponding Region + L2 folder
      var pdf_version = pdfs_folder.createFile(file_tmpl_copy.getAs("application/pdf"));
      pdf_version.setName(filename);

      // Write unique URL for new file and the folder holding new file
      sheet_ees.getRange(row+FIRST_ROW_EXTRACTED, 1, 1, 1).setValue("https://drive.google.com/a/oath.com/file/d/" + pdf_version.getId() + "/view?usp=sharing");
      SpreadsheetApp.flush();
    }
  }
  mail_merge();
}
