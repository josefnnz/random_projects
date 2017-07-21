function create_california_change_of_status_documents()
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

  // Google file ids
  var CA_STATUS_CHANGE_FOLDER_ID = "0B2QuBirnXYjxNGRMMEhJV0RHTzg"; // Folder: california_change_of_status_documents
  var CA_STATUS_CHANGE_TEMPLATE_ID = "10rH6rHEvrFroq6U5JO4TI3IEcxpm19XNc3Qt-A7QWjs"; // File: template_california_change_of_status
  var RIFS_SSID = "1s-fOV7IZ4ow-N6GTpqtiXg9j9VJR7nvQ3xMkTRvXsqw"; // File: Impacted Yahoos
  var RIFS_SHN = "create_docs"; // Sheet containing RIF'd employees to create docs for

  // Set folder where California Change of Status documents will be created
  var folder = DriveApp.getFolderById(CA_STATUS_CHANGE_FOLDER_ID);

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
  var LEGAL_FIRST_NAME_CIDX = 3 - 1; //NEEDTOUPDATE
  var LEGAL_LAST_NAME_CIDX = 4 - 1; //NEEDTOUPDATE
  var NOTIFICATION_DATE_CIDX = 8 - 1; //NEEDTOUPDATE
  var SEPARATION_DATE_CIDX = 12 - 1; //NEEDTOUPDATE
  var USA_STATE_CIDX = 28 - 1; //NEEDTOUPDATE
  var SOCIAL_SECURITY_NUMBER_CIDX = 29 - 1; //NEEDTOUPDATE

  function mail_merge() 
  {
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Only continue for California employees
      if (curr[USA_STATE_CIDX] === "California") 
      {
        // Extract required fields
        var notification_date = Utilities.formatDate(curr[NOTIFICATION_DATE_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy");
        var full_legal_name = curr[LEGAL_FIRST_NAME_CIDX] + " " + curr[LEGAL_LAST_NAME_CIDX];
        var social_security_number = curr[SOCIAL_SECURITY_NUMBER_CIDX];
        var separation_date = Utilities.formatDate(curr[SEPARATION_DATE_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy");

        // Copy the template
        var filename = full_legal_name;
        var file_new_ee_doc = DriveApp.getFileById(CA_STATUS_CHANGE_TEMPLATE_ID).makeCopy(filename, folder);
      
        // Fill-in copy with employee details
        var doc_new_ee_doc = DocumentApp.openById(file_new_ee_doc.getId());
        var body = doc_new_ee_doc.getBody();
        body.replaceText("<<notification_date>>", notification_date);
        body.replaceText("<<full_legal_name>>", full_legal_name);
        body.replaceText("<<social_security_number>>", social_security_number);
        body.replaceText("<<separation_date>>", separation_date);
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