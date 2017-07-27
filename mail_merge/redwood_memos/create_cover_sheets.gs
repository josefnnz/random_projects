// filter for rows of array: https://stackoverflow.com/questions/40849369/how-to-filter-an-array-of-arrays-google-app-script

function run_this() 
{

  // google file ids
  var SPREADSHEET_ID = "1zJ4uqrUe_jasPyZTFtF0kka98XeQVny_ZIYnZBWF98s"; // file: Generate Cover Sheets
  var SHEET_NAME = "notifying_mgrs";
  var COVER_SHEET_FOLDER_ID = "0B2QuBirnXYjxWGZJWDI3Sl90Yms"; // folder: Project Oath - Notification Memos
  var COVER_SHEET_TEMPLATE_ID = "1n6FAIxtaaH4U31FDV2T1B-QZjvVuGzYVR-itAWTxo2Y";

  // load sheet / cover sheet folder / cover sheet template
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var folder_destination = DriveApp.getFolderById(COVER_SHEET_FOLDER_ID);
  var file_cover_sheet_template = DriveApp.getFileById(COVER_SHEET_TEMPLATE_ID);
      
  // identify rows to extract
  var FIRST_ROW_EXTRACTED = sheet.getRange(1, 2).getValue();
  var LAST_ROW_EXTRACTED = sheet.getRange(2, 2).getValue();
  if (FIRST_ROW_EXTRACTED > LAST_ROW_EXTRACTED) { throw new Error("First Row Index (cell B1) must be less than or equal to the last row index (cell B2)"); }
  var NUM_ROWS_EXTRACTED = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;

  // extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_employees = sheet.getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_EXTRACTED, 18).getValues();
  var NUM_EMPLOYEES = values_employees.length;

  // field array indices, i.e. spreadsheet column index minus 1
  var EMP_ID_INDEX = 1 - 1;
  var EMP_NAME_INDEX = 2 - 1;
  var EMP_LEGAL_FIRST_NAME_INDEX = 3 - 1;
  var EMP_LEGAL_LAST_NAME_INDEX = 4 - 1;
  var EMP_TYPE_INDEX = 5 - 1;
  var MGR_EMP_ID_INDEX = 6 - 1;
  var MGR_NAME_INDEX = 7 - 1;
  var MGR_EMAIL_INDEX = 8 - 1;
  var MGR_OFFICE_INDEX = 9 - 1;
  var MGR_COUNTRY_INDEX = 10 - 1;
  var EMP_TRANSITION_INDEX = 11 - 1;
  var EMP_WFH_FLAG_INDEX = 12 - 1;
  var EMP_OFFICE_INDEX = 13 - 1;
  var EMP_COUNTRY_INDEX = 14 - 1;
  var EMP_VISA_HOLDER_FLAG_INDEX = 15 - 1;
  var EMP_VISA_STATUS_FLAG_INDEX = 16 - 1;
  var EMP_LEGAL_HOLD_FLAG_INDEX = 17 - 1;
  var EMP_WORK_PHONE_NUMBER = 18 - 1;
  
  // default table styles
  var headerStyle = {};  
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  //headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#336600';  
  //headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
  
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;  
  //cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  var centerText = {};
  centerText[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  
  // mail merge status column index in sheet
  var NOTES_COL_INDEX = 37;
  
  /*
   * Extract unique list of notifying manager employee ids from spreadsheet
   *
   * @return array of unique notifying manager employee ids
   */
  function unique_mgrs()
  {
    // extract notifying manager employee id column from spreadsheet and convert to 1-D array
    var array_2d_mgr_ids = sheet.getRange(FIRST_ROW_EXTRACTED, MGR_EMP_ID_INDEX+1, NUM_ROWS_EXTRACTED, 1).getValues();
    var array_mgr_ids = [];
    for (var i = 0; i < array_2d_mgr_ids.length; i++)
    {
      array_mgr_ids.push(array_2d_mgr_ids[i][0]);
    }
    
    // remove duplicate ids    
    var unique = [];
    for (var i = 0; i < array_mgr_ids.length; i++)
    {
      curr = array_mgr_ids[i];
      if (array_mgr_ids.indexOf(curr) == i)
      {
        unique.push(curr);
      }
    }
    
    // remove unnecessary or missing notifying manager ids
    unique.splice(unique.indexOf("L2/L3"), 1);
    unique.splice(unique.indexOf("non-USA"), 1);
    unique.splice(unique.indexOf("FTC"), 1);
    unique.splice(unique.indexOf("HR"), 1);
    unique.splice(unique.indexOf("#N/A"), 1);
    return unique;
  }
  
  function make_cover_sheets()
  {
    // get unique manager ids to iterate thru
    var mgrs = unique_mgrs();
    var mgrs = [sheet.getRange(2, 11).getValue()];
    for (var i = 0; i < mgrs.length; i++)
    {
      // filter out rows with given notifying manager employee id
      var mgr_emp_id = mgrs[i];
      
      var emps = values_employees.filter(function(row) { return row[MGR_EMP_ID_INDEX] === mgr_emp_id });
      
      var mgr_name = emps[0][MGR_NAME_INDEX];
      var mgr_office = emps[0][MGR_OFFICE_INDEX];

      var filename = mgr_office + "_" + mgr_name;
      var file_new_memo = file_cover_sheet_template.makeCopy(filename, folder_destination);
      
      // create new memo doc
      var doc_new_memo = DocumentApp.openById(file_new_memo.getId());
      var body = doc_new_memo.getBody();
      
      body.replaceText("<<notifying_manager>>", mgr_name);
      body.replaceText("<<notifying_manager_location>>", mgr_office);
      
      var tbl = body.appendTable();
      var tr = tbl.appendTableRow();
      tr.appendTableCell("Employee").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Cell #").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Location").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Transition").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Visa").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Legal Hold").getChild(0).asParagraph().setAttributes(headerStyle);
      
      tbl.setColumnWidth(0, 165);
      tbl.setColumnWidth(1, 110);
      tbl.setColumnWidth(2, 185)
      tbl.setColumnWidth(3, 70);
      tbl.setColumnWidth(4, 70);
      tbl.setColumnWidth(5, 70);
      
      headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
      tbl.setAttributes({"HORIZONTAL_ALIGNMENT" : "CENTER"});
            
      for (var j = 0; j < emps.length; j++)
      {
        var row = emps[j];
        var preferred_name = row[EMP_NAME_INDEX];
        var legal_first_name = row[EMP_LEGAL_FIRST_NAME_INDEX];
        var legal_last_name = row[EMP_LEGAL_LAST_NAME_INDEX];
        var legal_name = legal_first_name + " " + legal_last_name;
        var emp_office = row[EMP_OFFICE_INDEX];
        var emp_transition = row[EMP_TRANSITION_INDEX];
        var emp_wfh_flag = row[EMP_WFH_FLAG_INDEX];
        var emp_visa = row[EMP_VISA_HOLDER_FLAG_INDEX];
        var emp_legal_hold_flag = row[EMP_LEGAL_HOLD_FLAG_INDEX];
        var emp_work_phone_number = row[EMP_WORK_PHONE_NUMBER];
        var name = (preferred_name.search(legal_first_name) === -1) ? legal_name + " (" + preferred_name + ")" : legal_name;
        emp_office = (emp_wfh_flag === "WFH") ? "Work From Home" : emp_office.replace("US - ","");
        emp_transition = (emp_transition === "") ? "N" : emp_transition + " days"; 
        emp_visa = (emp_visa === "Y") ? "Y" : "N";
        emp_legal_hold_flag = (emp_legal_hold_flag === "Y") ? "Y" : "N";
        emp_work_phone_number = (emp_work_phone_number === "") ? "" : emp_work_phone_number;

        tr = tbl.appendTableRow();
        tr.appendTableCell(name).setAttributes(cellStyle);
        tr.appendTableCell(emp_work_phone_number).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
        tr.appendTableCell(emp_office).setAttributes(cellStyle);
        tr.appendTableCell(emp_transition).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
        tr.appendTableCell(emp_visa).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
        tr.appendTableCell(emp_legal_hold_flag).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
      }
           
      doc_new_memo.saveAndClose();
    
      // save memo as pdf in drive root directory. delete memo as google doc
      var pdf_version = folder_destination.createFile(file_new_memo.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_memo.setTrashed(true);
    }
  }
  
  make_cover_sheets();
}
