function ad_hoc_memo_generation() 
{

  // google file ids
  var RIF_SPREADSHEET_ID = "1oHZbpp-YNX3CiSmW2tLuJFRvoddr3p5P-cv10E40ToE"; // file: Master List - All L3s
  var RIF_SHEET_NAME = "MAKE_NEW_MEMOS_HERE";
  var MEMO_FOLDER_ID = "0B2QuBirnXYjxZ0N0WElaRnFsM0k"; // folder: Project Oath - Notification Memos
  var MEMO_TERMINATION_TEMPLATE_ID = "1bJNSA2BV3kg2StbYssYSBCLj93GefwufEb2MWpOeCig";
  var MEMO_TRANSITION_TEMPLATE_ID = "1hrqXWsuXDsGSc0g3RvdVO2ViF3MXMh6rMrExa5r_ihs";

  // set memo folder where all l3 folder will be created
  var folder_destination = DriveApp.getFolderById(MEMO_FOLDER_ID);

  // load sheet with employees with termination or transition from RIF spreadsheet
  var sheet_rifs = SpreadsheetApp.openById(RIF_SPREADSHEET_ID).getSheetByName(RIF_SHEET_NAME);
  // var range_last_row_index = sheet_rifs.getLastRow();
  // var range_last_col_index = sheet_rifs.getLastColumn();
    
  // identify rows to extract
  var FIRST_ROW_EXTRACTED = 1 * sheet_rifs.getSheetValues(1, 3, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * sheet_rifs.getSheetValues(2, 3, 1, 1);
  var NUM_ROWS_EXTRACTED = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
    
  // extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_employees = sheet_rifs.getRange(FIRST_ROW_EXTRACTED, 26, NUM_ROWS_EXTRACTED, 11).getValues();
  var NUM_EMPLOYEES = values_employees.length;

  // field array indices
  var IS_ON_TRANSITION_COL_INDEX = 0;
  var TRANSITION_BONUS_AMOUNT_COL_INDEX = 1;
  var LAST_DAY_OF_WORK_COL_INDEX = 2;
  var FIRST_NAME_COL_INDEX = 3;
  var LAST_NAME_COL_INDEX = 4;
  var USER_ID_COL_INDEX = 5;
  var L2_NAME_COL_INDEX = 6;
  var L3_NAME_COL_INDEX = 7;
  var OFFICE_COL_INDEX = 8;
  var COUNTRY_COL_INDEX = 9;
  var EMPLOYEE_TYPE_COL_INDEX = 10;

  // mail merge status column index in sheet
  var NOTES_COL_INDEX = 37;
    
  function mail_merge() 
  {
    // values array starts indexing at [0][0]
    for (var row = 0; row < NUM_EMPLOYEES; row++) 
    {
      // extract current employee
      var curr = values_employees[row];
      
      // skip (1) non-US, (2) L2/L3s, (3) FTC employees
      // only generate memos for US L4+ employees
      var country = curr[COUNTRY_COL_INDEX];
      var l2_name = curr[L2_NAME_COL_INDEX];
      var l3_name = curr[L3_NAME_COL_INDEX];
      var employee_type = curr[EMPLOYEE_TYPE_COL_INDEX];
      if (country != "United States of America" || l2_name === "" || l3_name === "" || employee_type === "Employee - Fixed Term Contract") 
      {
        Logger.log(FIRST_ROW_EXTRACTED+row);
        sheet_rifs.getRange(FIRST_ROW_EXTRACTED+row, NOTES_COL_INDEX).setValue("we are not generating memos for non-US employees NOR L2s NOR L3s NOR Fixed Term Contract employees");
        continue;
      }
      
      // extract required fields
      var is_on_transition = curr[IS_ON_TRANSITION_COL_INDEX];
      var transition_bonus_amount = curr[TRANSITION_BONUS_AMOUNT_COL_INDEX];
      var last_day_of_work = Utilities.formatDate(curr[LAST_DAY_OF_WORK_COL_INDEX], Session.getScriptTimeZone(), "MMMMM d, yyyy");
      var first_name = curr[FIRST_NAME_COL_INDEX];
      var last_name = curr[LAST_NAME_COL_INDEX];
      var user_id = curr[USER_ID_COL_INDEX];
      var office = curr[OFFICE_COL_INDEX];
      var employee_type = curr[EMPLOYEE_TYPE_COL_INDEX];
      
      // make file. do not make file if file already exists
      var filename = office + "_" + l2_name + "_" + l3_name + "_" + first_name + " " + last_name + " (" + user_id + ")";
      var subfiles = folder_destination.getFilesByName(filename);
      if (subfiles.hasNext())
      {
        sheet_rifs.getRange(FIRST_ROW_EXTRACTED+row, NOTES_COL_INDEX).setValue("memo already exists. id: " + subfiles.next().getId());
      }
      else
      {
        // copy memo template. select template based on transition/termination type
        var doc_template_id = (is_on_transition) ? MEMO_TRANSITION_TEMPLATE_ID : MEMO_TERMINATION_TEMPLATE_ID;
        var file_new_memo = DriveApp.getFileById(doc_template_id).makeCopy(filename, folder_destination);
      
        // create new memo doc
        var doc_new_memo = DocumentApp.openById(file_new_memo.getId());
        var body = doc_new_memo.getBody();
        
        body.replaceText("<<first_name>>", first_name);
        body.replaceText("<<last_name>>", last_name);
        body.replaceText("<<last_day_of_work>>", last_day_of_work);
        if (is_on_transition) { body.replaceText("<<transition_bonus_amount>>", transition_bonus_amount) }      
        
        doc_new_memo.saveAndClose();
      
        // save memo as pdf in drive root directory. delete memo as google doc
        var pdf_version = folder_destination.createFile(file_new_memo.getAs("application/pdf"));
        pdf_version.setName(filename);
        file_new_memo.setTrashed(true);
        
        sheet_rifs.getRange(FIRST_ROW_EXTRACTED+row, NOTES_COL_INDEX).setValue("memo created. id: " + pdf_version.getId());
      }
    }
  }
   
    
  //var prompt_text = "Are the first and last row employee names correct? If not, please fix the values in cells C1 and C2. \n\n"
  //                  + "";
  //var ui = SpreadsheetApp.getUi();
  //var result = ui.prompt(prompt_text, ui.ButtonSet.YES_NO_CANCEL);
    
  mail_merge();
  
}