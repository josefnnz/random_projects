function onOpen() 
{ 
  // Create button to launch data validation script in spreadsheet toolbar
  try
  {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Set Data Validations')
      .addItem('set_data_validations', 'set_data_validations')
      .addToUi(); 
  } 
  catch (e)
  {
  	// Log the error
  	Logger.log(e)
  }
}

function set_data_validations() 
{
  // Set constants
  var CINDEX_MAIN_FUNCT_AREA_ENTRY = 16; // Numeric column index for "New: Main Functional Area" selection column
  var CINDEX_SUB_TEAM_ENTRY = 17; // Numeric column index for "New: Sub-Team" selection column
  var CINDEX_MAIN_FUNCT_AREA_LIST = 21; // Numeric column index for "Main Functional Area Data Validation Values" field
  var CINDEX_SUB_TEAM_LIST = 22; // Numeric column index for "New Sub-Team Data Validation Values" field

  // Get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Get active spreadsheet Josef - Emp Roster  
  
  // Get data from "Copy of Roster" tab
  var sheet_roster = ss.getSheetByName("Copy of Roster"); // Get tab "Copy of Roster"
  var range_roster = sheet_roster.getDataRange();
  
  // Get sheets for Main Functional Area data validation Sub-Team data validation tabs
  var sheet_dv_main_funct_area = ss.getSheetByName("DVNewFunctArea");
  var sheet_dv_sub_team = ss.getSheetByName("DVNewSubTeam");
  
  // Start at Row 2 and set data validations
  for (var i = 2; i < sheet_roster.getLastRow(); i++)
  {
    // Get Main Functional Area and Sub Team data validation ranges
    var range_dv_main_funct_area = sheet_dv_main_funct_area.getRange("DVNewFunctArea!" + i + ":" + i);
    var range_dv_sub_team = sheet_dv_sub_team.getRange("DVNewSubTeam!" + i + ":" + i);
    // Set data validation ranges into cells
    var dv_main_funct_area = SpreadsheetApp.newDataValidation().requireValueInRange(range_dv_main_funct_area);
    var dv_sub_team = SpreadsheetApp.newDataValidation().requireValueInRange(range_dv_sub_team);
    sheet_roster.getRange(i, CINDEX_MAIN_FUNCT_AREA_ENTRY).setDataValidation(dv_main_funct_area);
    sheet_roster.getRange(i, CINDEX_SUB_TEAM_ENTRY).setDataValidation(dv_sub_team);
  }
}
