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
  var ADEA_FOLDER_ID = "1pNSx5uA5BpNiLyetwEjhJu8TY7-S6BhH"; // Folder: adea_forms
  var ADEA_TEMPLATE_ID = "1rT3vwxB81FTe04KEIu8e2wZBDKuDCpZttAO5fSTwP1I"; // File: Template - ADEA
  var ADEA_DATA_SSID = "1YaUHm_G5O72Twd9EfApfUwAAcLVPWhDRgFGi48OLLZ4"; // File: Project R2 - USA Calculations and Agreement Generator
  var ADEA_DATA_SHN = "ADEA"; // Sheet containing complete USA employee selection pool used

  // Set folder where ADEA documents will be created
  var folder = DriveApp.getFolderById(ADEA_FOLDER_ID);

  // Load ADEA template
  var file_adea_template = DriveApp.getFileById(ADEA_TEMPLATE_ID);

  // Load Spreadsheet with complete USA employee selection pool and Oath L2 notice range dates
  var ees = SpreadsheetApp.openById(ADEA_DATA_SSID);

  // L2s to generate ADEA forms for
  var L2s = ["Allie Kline", "Atte Lahtiranta", "Jeffrey Bonforte", "John DeVine", "Julie Jacobs", "Simon Khalaf", "Tim Mahlman", "Timothy Lemmon"];

  // Mappings from L2 to L2 Orgname and L2 to factors used to determine termination selections
  var L2_to_orgname = {"Allie Kline"      : "Marketing and Communications",
                       "Atte Lahtiranta"  : "Technology",
                       "Jeffrey Bonforte" : "Communications, Search, and Data",
                       "John DeVine"      : "Sales and Customer Operations",
                       "Julie Jacobs"     : "Legal and Corporate Services",
                       "Simon Khalaf"     : "Media Brands and Products",
                       "Tim Mahlman"      : "AdTech Platforms",
                       "Timothy Lemmon"   : "Business Operations"};

  var L2_to_factors = {"Allie Kline"      : "job function, skillset, and the direction of the business going forward",
                       "Atte Lahtiranta"  : "job function, location, performance, and the direction of the business going forward",
                       "Jeffrey Bonforte" : "job function, location, performance, and the direction of the business going forward",
                       "John DeVine"      : "job function, location, performance, and the direction of the business going forward",
                       "Julie Jacobs"     : "job function, location, performance, and the direction of the business going forward",
                       "Simon Khalaf"     : "job function, location, performance, and the direction of the business going forward",
                       "Tim Mahlman"      : "job function, location, performance, and the direction of the business going forward",
                       "Timothy Lemmon"   : "job function and the direction of the business going forward"};

  // Identify specific first and last rows to extract from selection pool sheet
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetByName(ADEA_DATA_SHN).getSheetValues(2, 2, 1, 1); // Cell B2
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetByName(ADEA_DATA_SHN).getSheetValues(3, 2, 1, 1); // Cell B3

  // Identify total number of rows and columns to extract from selection pool sheet
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 6; // Columns A to F

  // Extract range of USA selection pool employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_groups = ees.getSheetByName(ADEA_DATA_SHN).getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();

  // Array column indices for required fields
  var L2_CIDX = 1 - 1;
  var L2_ORGNAME_CIDX = 2 - 1;
  var JOB_TITLE_CIDX = 3 - 1;
  var AGE_CIDX = 4 - 1;
  var NOT_SELECTED_CIDX = 5 - 1;
  var SELECTED_CIDX = 6 - 1;
  
  // default table styles
  var CELL_FONT_SIZE = 10;
  var CELL_FONT_FAMILY = "Calibri";
  var CELL_PADDING_TOP = 2;
  var CELL_PADDING_BOTTOM = 2;

  var headerStyle = {};  
  headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  headerStyle[DocumentApp.Attribute.FONT_FAMILY] = CELL_FONT_FAMILY;
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = CELL_FONT_SIZE;
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.PADDING_TOP] = CELL_PADDING_TOP;
  headerStyle[DocumentApp.Attribute.PADDING_BOTTOM] = CELL_PADDING_BOTTOM;

  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = CELL_FONT_FAMILY;
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = CELL_FONT_SIZE;
  cellStyle[DocumentApp.Attribute.BOLD] = false;  
  cellStyle[DocumentApp.Attribute.PADDING_TOP] = CELL_PADDING_TOP;
  cellStyle[DocumentApp.Attribute.PADDING_BOTTOM] = CELL_PADDING_BOTTOM;
  
  var centerText = {};
  centerText[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  centerText[DocumentApp.Attribute.FONT_FAMILY] = CELL_FONT_FAMILY;
  centerText[DocumentApp.Attribute.FONT_SIZE] = CELL_FONT_SIZE;
  centerText[DocumentApp.Attribute.PADDING_TOP] = CELL_PADDING_TOP;
  centerText[DocumentApp.Attribute.PADDING_BOTTOM] = CELL_PADDING_BOTTOM;

  function mail_merge() 
  {
    // Confirm user wants to run script
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Please check cell B1 and confirm you entered the correct date. Please check cells B2 and B3 and confirm they capture the first and last rows on the spreadsheet are referenced. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) 
    {
     return;
    }

    var date_data_prepared = Utilities.formatDate(ees.getSheetByName(ADEA_DATA_SHN).getRange(1, 2).getValue(), Session.getScriptTimeZone(), "MMMMM d, yyyy"); // Cell B1
    for (var i = 0; i < L2s.length; i++)
    {
      var curr_L2 = L2s[i];
      var curr_L2_orgname = L2_to_orgname[curr_L2];
      var curr_factors = L2_to_factors[curr_L2];

      // Copy template
      var filename = "ADEA - " + curr_L2 + " - " + curr_L2_orgname;
      var file_new_memo = file_adea_template.makeCopy(filename, folder);
      
      // Create new memo doc
      var doc_new_memo = DocumentApp.openById(file_new_memo.getId());
      var body = doc_new_memo.getBody();
      
      body.replaceText("<<L2_orgname>>", curr_L2_orgname);
      body.replaceText("<<factors_used_to_determine_termination>>", curr_factors);
      body.replaceText("<<date_data_prepared>>", date_data_prepared);

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
      var curr_L2_groups = values_groups.filter(function(row) { return row[L2_ORGNAME_CIDX] == curr_L2_orgname });
      var NUM_GROUPS = curr_L2_groups.length;
      
      for (var row_grp = 0; row_grp < NUM_GROUPS; row_grp++)
      {
        if (row_grp % 500 == 0)
        {
          doc_new_memo.saveAndClose();
          doc_new_memo = DocumentApp.openById(file_new_memo.getId());
          body = doc_new_memo.getBody();
          tbl = body.findElement(DocumentApp.ElementType.TABLE).getElement();
        }
        var curr_grp = curr_L2_groups[row_grp];
        var job_title = curr_grp[JOB_TITLE_CIDX];
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
      //var pdf_version = folder.createFile(file_new_memo.getAs("application/pdf"));
      //pdf_version.setName(filename);
      //file_new_memo.setTrashed(true);
    }
  }
  mail_merge();
}