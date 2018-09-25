/**
 * Abbreviations used (case-insensitive):
 *
 * eib -- enterprise interface builder
 * cidx -- column index in array (i.e. starts at ZERO)
 * shn -- sheet name
 * ssid -- spreadsheet
 * tml -- template
**/

// Identify specific first and last rows and columns to extract and number of rows/columns to extract
var FIRST_ROW_EXTRACTED = 5;
var LAST_ROW_EXTRACTED = 774;
var FIRST_COL_EXTRACTED = 2;
var LAST_COL_EXTRACTED = 26;
var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
var NUM_COLS_TO_EXTRACT = LAST_COL_EXTRACTED - FIRST_COL_EXTRACTED + 1;

// Google file ids
var FOLDER_ID_MAIN = "1C7r8RfsGBT5jk2-_wVWfOD5A0msIg59a"; // Folder: 
var SSID_DATA = "1xJlu13lq0S6-WW7oelRE-CPmJXocbTn3MRAlcCM-_rQ"; // File: 
var SHN_DATA = "Data"; // Sheet with data to transform into EIB
var SSID_EIB_TML = "1sm-Vn4n2e89pzPe_-M8KzDCx5jjxKDvPlIWtn0ToY7Y"; // File: 
var SHN_EIB_TAB_CHANGE_JOB = "Change Job"; // Sheet name for Change Job tab in Change_Job EIB 
var SHN_EIB_TAB_PROPOSE_COMP_CHANGE = "Propose Compensation"; // Sheet name for Request Compensation Change tab in Change_Job EIB

// Constants
var NUM_COLS_IN_CHANGE_JOB_TAB = 52;
var NUM_COLS_IN_PROPOSE_COMP_CHANGE_TAB = 103;

// Indices for data array
var indices = 
{
  "Employee_ID" : 0,
  "Employee_Name" : 1,
  "Effective_Date" : 2,
  "Reason" : 3,
  "Job_Code" : 4,
  "Job_Profile_Name" : 5,
  "Business_Title" : 6,
  "Job_Category" : 7,
  "Job_Level" : 8,
  "Pay_Rate_Type_USA_Specific_Values" : 9,
  "Work_Country" : 10,
  "Work_Office" : 11,
  "Compensation_Package" : 12,
  "Compensation_Grade" : 13,
  "Compensation_Grade_Profile" : 14,
  "Hourly_Salary_Plan" : 15,
  "Annual_Rate_Salary_Rate" : 16,
  "Hourly_Rate" : 17,
  "Currency" : 18,
  "Frequency" : 19,
  "Bonus_Plan" : 20,
  "Amount_Percent_Based" : 21,
  "Target_Bonus_Percent" : 22,
  "Target_Bonus_Amount" : 23,
  "Is_On_Individual_Target_Yes_No" : 24
}

function create_full_eib()
{
  // Create Change Job and Propose Compensation Change arrays
  var eib_arrays = create_promo_eib_arrays();
  array_eib_change_job = eib_arrays[0];
  array_eib_propose_comp_change = eib_arrays[1];
  var NUM_ROWS_TO_WRITE_CHANGE_JOB = array_eib_change_job.length;
  var NUM_COLS_TO_WRITE_CHANGE_JOB = array_eib_change_job[0].length;
  var NUM_ROWS_TO_WRITE_PROPOSE_COMP_CHANGE = array_eib_propose_comp_change.length;
  var NUM_COLS_TO_WRITE_PROPOSE_COMP_CHANGE = array_eib_propose_comp_change[0].length;

  // Create filename -- append current datetime in format yyyy-MM-dd HH_MM PDT
  var datetimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH_mm") + " PDT";
  var filename = "HR1403375 - 2018 Mid-Year Promotions - " + datetimestamp;

  // Get folder
  var folder = DriveApp.getFolderById(FOLDER_ID_MAIN);

  // Make copy of EIB template in folder, open copy, and write new values
  var file_tml_cpy = DriveApp.getFileById(SSID_EIB_TML).makeCopy(filename, folder);
  var ss_new_eib = SpreadsheetApp.openById(file_tml_cpy.getId());
  var sheet_new_eib_change_job = ss_new_eib.getSheetByName(SHN_EIB_TAB_CHANGE_JOB);
  var sheet_new_eib_propose_comp_change = ss_new_eib.getSheetByName(SHN_EIB_TAB_PROPOSE_COMP_CHANGE);
  sheet_new_eib_change_job.getRange(6, 1, NUM_ROWS_TO_WRITE_CHANGE_JOB, NUM_COLS_TO_WRITE_CHANGE_JOB).setValues(array_eib_change_job);
  sheet_new_eib_propose_comp_change.getRange(6, 1, NUM_ROWS_TO_WRITE_PROPOSE_COMP_CHANGE, NUM_COLS_TO_WRITE_PROPOSE_COMP_CHANGE).setValues(array_eib_propose_comp_change);
  SpreadsheetApp.flush()

  // Save new EIB as Excel file, and delete GSheet version
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + file_tml_cpy.getId() + "&exportFormat=xlsx";
  var params = 
  {
    method : "get",
    headers : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions : true
  };
  var blob = UrlFetchApp.fetch(url, params).getBlob();
  blob.setName(filename + ".xlsx");
  var excel_new_eib = folder.createFile(blob);
  file_tml_cpy.setTrashed(true);
}


function create_promo_eib_arrays()
{
  // Get sheet with promotions data
  var values_promos = SpreadsheetApp.openById(SSID_DATA).getSheetByName(SHN_DATA).getRange(FIRST_ROW_EXTRACTED, FIRST_COL_EXTRACTED, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();

  // Create empty 2D arrays for Change Job and Request Compensation Change EIB tabs
  var array_eib_change_job = [];
  var array_eib_propose_comp_change = [];
  var sskey = 1;
  for (var i = 0; i < values_promos.length; i++)
  {
    // Add details here
    var curr = values_promos[i];
    array_eib_change_job.push(create_change_job_tab_eib_row(sskey, curr)); // add job data to Change Job tab
    array_eib_propose_comp_change.push(create_propose_comp_change_tab_eib_row(sskey, curr)); // add comp data to Propose Compensation Change tab
    sskey++; // increment spreadsheet key
  }
  return [array_eib_change_job, array_eib_propose_comp_change];
}

function create_change_job_tab_eib_row(sskey, promo)
{
  var row = Array.apply(null, Array(NUM_COLS_IN_CHANGE_JOB_TAB)).map(String.prototype.valueOf, "");
  row[1] = sskey; // spreadsheet key
  row[2] = promo[indices["Employee_ID"]]; // employee id
  row[4] = promo[indices["Effective_Date"]]; // effective date in format YYYY-MM-DD
  row[5] = promo[indices["Reason"]]; // change job subcategory id
  row[14] = promo[indices["Job_Code"]]; // job profile id
  row[16] = promo[indices["Business_Title"]]; // business title
  return row;
}

function create_propose_comp_change_tab_eib_row(sskey, promo)
{
  var row = Array.apply(null, Array(NUM_COLS_IN_PROPOSE_COMP_CHANGE_TAB)).map(String.prototype.valueOf, "");
  row[1] = sskey; // spreadsheet key
  row[5] = promo[indices["Compensation_Package"]]; // compensation package
  row[6] = promo[indices["Compensation_Grade"]]; // compensation grade
  row[7] = promo[indices["Compensation_Grade_Profile"]];
  row[10] = "Y"; // overwrite all existing Hourly/Salary plan assignments
  row[11] = 1; // identify row for new Hourly/Salary plan assignment
  row[12] = promo[indices["Hourly_Salary_Plan"]]; // Hourly/Salary plan
  var is_hourly_employee = promo[indices["Hourly_Salary_Plan"]] === "Hourly Plan";
  row[13] = is_hourly_employee ? promo[indices["Hourly_Rate"]] : promo[indices["Annual_Rate_Salary_Rate"]]; // set comp plan amount (hourly or annual rate)
  row[14] = promo[indices["Currency"]]; // comp plan currency
  row[15] = promo[indices["Frequency"]]; // comp plan frequency (annual or hourly)
  row[48] = "Y"; // overwrite all existing Bonus Plan assignments
  row[49] = 1; // identify row for new Bonus Plan assignment
  row[50] = promo[indices["Bonus_Plan"]]; // Bonus plan
  var is_on_individual_target = promo[indices["Is_On_Individual_Target_Yes_No"]] === "Yes";
  var is_amount_based_plan = promo[indices["Amount_Percent_Based"]] === "Amount_Based";
  if (!is_on_individual_target)
  {
    row[55] = 1; // set employee's target bonus to default target value
  }
  else if (is_amount_based_plan)
  {
    row[51] = promo[indices["Target_Bonus_Amount"]]; // set individual target bonus amount
  } 
  else
  {
    row[52] = promo[indices["Target_Bonus_Percent"]]; // set individual target bonus percent
  }
  return row;
}
