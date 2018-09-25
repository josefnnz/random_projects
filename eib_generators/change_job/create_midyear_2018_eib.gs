/**
 * Abbreviations used (case-insensitive):
 *
 * eib -- enterprise interface builder
 * cidx -- column index in array (i.e. starts at ZERO)
 * shn -- sheet name
 * ssid -- spreadsheet
 * tml -- template
**/

// Google file ids
var FOLDER_ID_MAIN = "PLACEHOLDER"; // Folder: 
var SSID_DATA = "PLACEHOLDER"; // File: 
var SHN_DATA = "PLACEHOLDER"; // Sheet with data to transform into EIB
var SSID_EIB_TML = "PLACEHOLDER"; // File: 
var SHN_EIB_TAB_CHANGE_JOB = "PLACEHOLDER"; // Sheet name for Change Job tab in Change_Job EIB 
var SHN_EIB_TAB_REQ_COMP_CHANGE = "PLACEHOLDER"; // Sheet name for Request Compensation Change tab in Change_Job EIB

// Constants
var NUM_COLUMN_HEADERS = 2;
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


function create_promo_eib_arrays()
{
  // Get sheet with promotions data
  var values_promos = SpreadsheetApp.openById(SSID_DATA).getSheetByName(SHN_DATA).getDataRange();

  // Create empty 2D arrays for Change Job and Request Compensation Change EIB tabs
  var array_eib_change_job = [];
  var array_eib_propose_comp_change = [];
  var sskey = 1;
  for (var i = 0+NUM_COLUMN_HEADERS; i < values_promos.length; i++)
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
