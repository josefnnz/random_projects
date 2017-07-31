function create_adea_forms()
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
  // var ui = SpreadsheetApp.getUi();
  // var response = ui.alert("Please check cells C1 and C2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
  // if (response !== ui.Button.OK) {
  //  return;
  // }

  // Google file ids
  var ADEA_FOLDER_ID = "0B2QuBirnXYjxd0xBbXpVZkVJWTg"; // Folder: adea_documents
  var ADEA_TEMPLATE_ID = "1IT5D7rZzlBk9eAt66vJz7gU7FMqWnFnJ2pDJ2tVy_RI"; // File: template_adea
  var ADEA_DATA_SSID = "1PjWyLTJjMUjtYAN39pzDxbU0bhsIyR6fSlTRmlZsvFQ"; // File: DATAFILE_adea_data
  var ADEA_DATA_SHN = "FORMATTED_ADEA_DATA"; // Sheet containing complete USA employee selection pool used
  var NOTICE_RANGES_SHN = "OATH_L2_NOTICE_RANGES"; // Sheet containing notice range dates by Oath L2 / L3 orgs

  // Set folder where ADEA documents will be created
  var folder = DriveApp.getFolderById(ADEA_FOLDER_ID);

  // Load ADEA template
  var file_adea_template = DriveApp.getFileById(ADEA_TEMPLATE_ID);

  // Load Spreadsheet with complete USA employee selection pool and Oath L2 notice range dates
  var ees = SpreadsheetApp.openById(ADEA_DATA_SSID);
  
  /** -------------------- LOAD NOTICE RANGE DATES SHEET -------------------- **/

  // Load sheet with notice range dates
  // var values_notice_ranges = ees.getSheetByName(NOTICE_RANGES_SHN).getRange(2, 1, 13, 5).getValues();
  var values_notice_ranges = ees.getSheetByName(NOTICE_RANGES_SHN).getRange(2, 1, 13, 5).getValues();
  var NUM_OATH_L2S = values_notice_ranges.length;

  // Array column indices for required fields
  var L2_or_L3_CIDX = 1 - 1;
  var OATH_L2_NAME_CIDX = 2 - 1;
  var OATH_L2_ORGNAME_CIDX = 3 - 1;
  var FIRST_NOTICE_DATE_CIDX = 4 - 1;
  var LAST_NOTICE_DATE_CIDX = 5 - 1;

  /** -------------------- LOAD SELECTION POOL SHEET -------------------- **/

  // Identify specific first and last rows to extract from selection pool sheet
  var FIRST_ROW_EXTRACTED = 5; //NEEDTOUPDATE
  var LAST_ROW_EXTRACTED = 7060; //NEEDTOUPDATE

  // Identify total number of rows and columns to extract from selection pool sheet
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 6; // Columns A to BG -- NEEDTOUPDATE

  // Extract range of USA selection pool employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_groups = ees.getSheetByName(ADEA_DATA_SHN).getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();

  // Array column indices for required fields
  var ADEA_OATH_L2_NAME_CIDX = 1 - 1;
  var ADEA_OATH_L2_ORGNAME_CIDX = 2 - 1;
  var CURRENT_JOB_TITLE_CIDX = 3 - 1;
  var AGE_CIDX = 4 - 1;
  var NOT_SELECTED_CIDX = 5 - 1;
  var SELECTED_CIDX = 6 - 1;
  
  // default table styles
  var headerStyle = {};  
  headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  // headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#336600';  
  // headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';

  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  cellStyle[DocumentApp.Attribute.BOLD] = false;  
  //cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  var centerText = {};
  centerText[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  centerText[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  centerText[DocumentApp.Attribute.FONT_SIZE] = 10;

  function mail_merge() 
  {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMMM d, yyyy");
    today = "August 2, 2017"
    for (var row_notice_range = 0; row_notice_range < NUM_OATH_L2S; row_notice_range++)
    {
      curr_notice_range = values_notice_ranges[row_notice_range];
      var curr_L2 = curr_notice_range[OATH_L2_NAME_CIDX];
      var curr_L2_orgname = curr_notice_range[OATH_L2_ORGNAME_CIDX];
      var first_notice_date = Utilities.formatDate(curr_notice_range[FIRST_NOTICE_DATE_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy");
      var last_notice_date = Utilities.formatDate(curr_notice_range[LAST_NOTICE_DATE_CIDX], Session.getScriptTimeZone(), "MMMMM d, yyyy");
      var notice_range = (first_notice_date == last_notice_date) ? "on " + first_notice_date : "between " + first_notice_date + " and " + last_notice_date;

      // Copy template
      var filename = "ADEA - " + curr_L2 + " - " + curr_L2_orgname;
      var file_new_memo = file_adea_template.makeCopy(filename, folder);
      
      // Create new memo doc
      var doc_new_memo = DocumentApp.openById(file_new_memo.getId());
      var body = doc_new_memo.getBody();
      
      body.replaceText("<<oath_L2_org_name>>", curr_L2_orgname);
      body.replaceText("<<notification_range>>", notice_range);
      body.replaceText("<<date_data_prepared>>", today);

      var tbl = body.appendTable();
      var tr = tbl.appendTableRow();
      tr.appendTableCell("Current Job Title").setBackgroundColor("#DCE6F1").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Age").setBackgroundColor("#DCE6F1").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Not Selected").setBackgroundColor("#DCE6F1").getChild(0).asParagraph().setAttributes(headerStyle);
      tr.appendTableCell("Selected").setBackgroundColor("#DCE6F1").getChild(0).asParagraph().setAttributes(headerStyle);
      
      tbl.setColumnWidth(0, 310);
      tbl.setColumnWidth(1, 50);
      tbl.setColumnWidth(2, 50)
      tbl.setColumnWidth(3, 50);

      tbl.setAttributes({"HORIZONTAL_ALIGNMENT" : "CENTER"});

      // Get current L2s groups
      var curr_L2_groups = values_groups.filter(function(row) { return row[ADEA_OATH_L2_ORGNAME_CIDX] == curr_L2_orgname });
      var NUM_GROUPS = curr_L2_groups.length;
      
      for (var row_grp = 0; row_grp < NUM_GROUPS; row_grp++)
      {
        var curr_grp = curr_L2_groups[row_grp];
        var job_title = curr_grp[CURRENT_JOB_TITLE_CIDX];
        var age = curr_grp[AGE_CIDX];
        var not_selected = curr_grp[NOT_SELECTED_CIDX];
        var selected = curr_grp[SELECTED_CIDX];

        tr = tbl.appendTableRow();
        tr.appendTableCell(job_title).setAttributes(cellStyle);
        tr.appendTableCell(age).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
        tr.appendTableCell(not_selected).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
        tr.appendTableCell(selected).setAttributes(cellStyle).getChild(0).asParagraph().setAttributes(centerText);
      }

      doc_new_memo.saveAndClose();
    
      // save memo as pdf in drive root directory. delete memo as google doc
      var pdf_version = folder.createFile(file_new_memo.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_memo.setTrashed(true);
    }
  }
  mail_merge();
}