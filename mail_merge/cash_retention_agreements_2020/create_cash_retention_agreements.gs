// identify rows and columns to extract
var FIRST_ROW_EXTRACTED = 34;
var LAST_ROW_EXTRACTED = 100;
var NUM_ROWS_EXTRACTED = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;

var FIRST_COLUMN_EXTRACTED = 1;
var LAST_COLUMN_EXTRACTED = 12;
var NUM_COLUMNS_EXTRACTS = LAST_COLUMN_EXTRACTED - FIRST_COLUMN_EXTRACTED;

// set google file ids
var SPREADSHEET_ID_CASH_RETENTION = "1ArvOJq-jObY_NUcnn_e6T8Gfs0oFG35F6x8HXH9LGzQ"; // file: 2020 Transitional LTI Cash Retentions (with HR) - HR2050151
var SHEET_NAME_CASH_RETENTION_EMPLOYEES = "Transition Cash Retention";
var FOLDER_ID_CASH_RETENTION_AGREEMENTS = "1l1s2FgpEixQqSr_MiZb4DcL0vbGc7Cx4"; // folder: HR2050151 - 2020 Cash Retention Agreements
var TEMPLATE_ID_USA_RETENTION_2_INSTALLMENTS = "1mXjUDU98a3yM_hrdRjnWH7ktP-lGZHnSp_fNMXZBqds";
var TEMPLATE_ID_INTL_RETENTION_2_INSTALLMENTS = "15YwJK895F2eXBAtpXyYch9QIK5b9rw4eZrhcCPmfCRo";

// get folder where all L2/L3 folders will be created
var folder_cash_retention_agreements = DriveApp.getFolderById(FOLDER_ID_CASH_RETENTION_AGREEMENTS);

// load sheet with employees receiving cash retention agreements
var sheet_cash_retention_employees = SpreadsheetApp.openById(SPREADSHEET_ID_CASH_RETENTION).getSheetByName(SHEET_NAME_CASH_RETENTION_EMPLOYEES);

// extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
var values_employees = sheet_cash_retention_employees.getRange(FIRST_ROW_EXTRACTED, FIRST_COLUMN_EXTRACTED, NUM_ROWS_EXTRACTED, NUM_COLUMNS_EXTRACTS).getValues();
var NUM_EMPLOYEES = values_employees.length;

// field array indices
var COLUMN_INDEX_PREFERRED_NAME = 0;
var COLUMN_INDEX_MANAGER_NAME = 1;
var COLUMN_INDEX_LETTER_DATE = 2;
var COLUMN_INDEX_PERIOD_1_START = 3;
var COLUMN_INDEX_PERIOD_1_END = 4;
var COLUMN_INDEX_PERIOD_2_START = 5;
var COLUMN_INDEX_PERIOD_2_END = 6;
var COLUMN_INDEX_CURRENCY = 7;
var COLUMN_INDEX_TOTAL_AMOUNT = 8;
var COLUMN_INDEX_PAYMENT_1 = 9;
var COLUMN_INDEX_PAYMENT_2 = 10;
var COLUMN_INDEX_EMPLOYEE_ID = 11;
var COLUMN_INDEX_COUNTRY = 12;
var COLUMN_INDEX_L2 = 13;
var COLUMN_INDEX_L3 = 14;
var COLUMN_INDEX_L4 = 15;
var COLUMN_INDEX_FILENAME = 16;

// mail merge status column index in sheet
var NOTES_COL_INDEX = 37;

function mail_merge() 
{
  // values array starts indexing at [0][0]
  for (var row = 0; row < NUM_EMPLOYEES; row++) 
  {
    // extract current employee
    var curr = values_employees[row];
    
    // extract required fields
    var preferred_name = curr[COLUMN_INDEX_PREFERRED_NAME];
    var manager_name = curr[COLUMN_INDEX_MANAGER_NAME];
    var letter_date = curr[COLUMN_INDEX_LETTER_DATE];
    var period_1_start = curr[COLUMN_INDEX_PERIOD_1_START];
    var period_1_end = curr[COLUMN_INDEX_PERIOD_1_END];
    var period_2_start = curr[COLUMN_INDEX_PERIOD_2_START];
    var period_2_end = curr[COLUMN_INDEX_PERIOD_2_END];
    var currency = curr[COLUMN_INDEX_CURRENCY];
    var total_amount = curr[COLUMN_INDEX_TOTAL_AMOUNT];
    var payment_1 = curr[COLUMN_INDEX_PAYMENT_1];
    var payment_2 = curr[COLUMN_INDEX_PAYMENT_2];
    var employee_id = curr[COLUMN_INDEX_EMPLOYEE_ID];
    var country = curr[COLUMN_INDEX_COUNTRY];
    var L2 = curr[COLUMN_INDEX_L2];
    var L3 = curr[COLUMN_INDEX_L3];
    var L4 = curr[COLUMN_INDEX_L4];
    var filename = curr[COLUMN_INDEX_FILENAME];
    // var last_day_of_work = Utilities.formatDate(curr[LAST_DAY_OF_WORK_COL_INDEX], Session.getScriptTimeZone(), "MMMMM d, yyyy");

    // make file copy of template in l3 folder. create new l3 folder if inexistent
    if (L2 === "") { L2 = "Direct Report to Guru Gowrappan"; }
    var L2_subfolders = folder_cash_retention_agreements.getFoldersByName(L2);
    var L2_folder = (L2_subfolders.hasNext()) ? L2_subfolders.next() : folder_memos.createFolder(L2);
    
    if (L2 === "") { L3 = "Direct Report to Guru Gowrappan"; }
    else if (L3 === "") { L3 = "Direct Report to " + L2; }
    var L3_subfolders = L2_folder.getFoldersByName(L3);
    var folder_destination = (L3_subfolders.hasNext()) ? L3_subfolders.next() : L2_folder.createFolder(L3);
    
    // make file -- do not make file if file already exists
    var subfiles = folder_destination.getFilesByName(filename);
    if (subfiles.hasNext())
    {
      sheet_cash_retention_employees.getRange(FIRST_ROW_EXTRACTED+row, NOTES_COL_INDEX).setValue("Agreement already exists. Id: " + subfiles.next().getId());
    }
    else
    {
      // copy cash retention agreement template. select template based on USA or international employee.
      var doc_template_id = (country === "United States of America") ? TEMPLATE_ID_USA_RETENTION_2_INSTALLMENTS : TEMPLATE_ID_INTL_RETENTION_2_INSTALLMENTS;
      var file_new_agreement = DriveApp.getFileById(doc_template_id).makeCopy(filename, folder_destination);
    
      // create new agreement doc
      var doc_new_agreement = DocumentApp.openById(file_new_agreement.getId());
      var body = doc_new_agreement.getBody();

      body.replaceText("{{Preferred Name}}", preferred_name);
      body.replaceText("{{Manager Name}}", manager_name);
      body.replaceText("{{Letter Date}}", letter_date);
      body.replaceText("{{Period 1 Start}}", period_1_start);
      body.replaceText("{{Period 1 End}}", period_1_end);
      body.replaceText("{{Period 2 Start}}", period_2_start);
      body.replaceText("{{Period 2 End}}", period_2_end);
      body.replaceText("{{Currency}}", currency);
      body.replaceText("{{Total Amount}}", total_amount);
      body.replaceText("{{Payment 1}}", payment_1);
      body.replaceText("{{Payment 2}} ", payment_2);
      body.replaceText("{{Emp ID}}", employee_id);

      doc_new_agreement.saveAndClose();
    
      // save memo as pdf in drive root directory. delete memo as google doc
      var pdf_version = folder_destination.createFile(file_new_agreement.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_agreement.setTrashed(true);
      
      sheet_cash_retention_employees.getRange(FIRST_ROW_EXTRACTED+row, NOTES_COL_INDEX).setValue("Agreement created. Id: " + pdf_version.getId());
    }
  }
}