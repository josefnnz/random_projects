var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetsCount = ss.getNumSheets();
var sheets = ss.getSheets();

var tabs_in_order = 
[
  "Overview",
  "Hire Employee",
  "Propose Compensation for Hire",
  "Update ID Information",
  "Edit Government IDs",
  "Edit Passports and Visas",
  "Edit License",
  "Edit Custom IDs",
  "Edit Assign Organization",
  "Assign Pay Group",
  "Review Payroll Interface",
  "Review Payroll Interface Event",
  "Request One Time Payment",
  "Request One Time Payment for Referral",
  "Request Stock Grant",
  "Create Workday Account",
  "Assign Matrix Organization",
  "Change Personal Information",
  "Create Provisioning Event",
  "Create Benefit Life Event",
  "Maintain Employee Contracts Sub Business Process",
  "Edit Service Dates",
  "Remove Retiree Status",
  "Check Position Budget",
  "Assign Costing Allocation",
  "Edit Background Check",
  "Add Academic Appointment",
  "Create Workday Account Sub Business Process for Academic Affiliate",
  "Manage Professional Affiliation Sub Business Process for Academic Affiliate",
  "Manage Education Sub Business Process for Academic Affiliate",
  "Manage Instructor Eligibility Sub Business Process for Academic Affiliate",
  "Assign Employee Collective Agreement",
  "Manage Employee Probation Period Sub Business Process",
  "Emergency Contacts",
  "Onboarding Setup",
  "Student Employment Eligibility",
  "Manage Union Membership",
  "Edit Notice Periods"
]

var hidden_sheets = 
[
	"Add Academic Appointment",
	"Assign Costing Allocation",
	"Assign Employee Collective Agreement",
	"Assign Matrix Organization",
    "Change Personal Information",
	"Check Position Budget",
	"Create Benefit Life Event",
	"Create Provisioning Event",
	"Create Workday Account",
	"Create Workday Account Sub Business Process for Academic Affiliate",
	"Edit Background Check",
    "Edit Custom IDs",
    "Edit Government IDs",
	"Edit License",
	"Edit Notice Periods",
    "Edit Passports and Visas",
	"Emergency Contacts",
	"Maintain Employee Contracts Sub Business Process",
	"Manage Education Sub Business Process for Academic Affiliate",
	"Manage Employee Probation Period Sub Business Process",
	"Manage Instructor Eligibility Sub Business Process for Academic Affiliate",
	"Manage Professional Affiliation Sub Business Process for Academic Affiliate",
	"Manage Union Membership",
	"Onboarding Setup",
	"Remove Retiree Status",
	"Request One Time Payment",
	"Request One Time Payment for Referral",
	"Request Stock Grant",
	"Review Payroll Interface",
	"Review Payroll Interface Event",
	"Student Employment Eligibility",
    "Update ID Information"
]

function onOpen() 
{ 
 // Try New Google Sheets method
  try
  {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Spreadsheet Cleanup')
      .addItem('Unhide All Sheets', 'showSheets')
      .addItem('Hide Sheets', 'hideSheets')
      .addItem('Reorder Sheets', 'reorderSheets')
      .addToUi(); 
  } 
  catch (e)
  {
  	// Log the error
  	Logger.log(e)
  }
  finally
  {
  	// Use old Google Spreadsheet method
    var items = 
    [
      {name: 'Hide Sheets', functionName: 'hideSheets'},
      {name: 'Unhide All Sheets', functionName: 'showSheets'},
      {name: 'Reorder Sheets', functionName: 'reorderSheets'},
    ];
    ss.addMenu('Spreadsheet Cleanup', items);
  }
}

function hideSheets() 
{
  for (var i = 0; i < sheetsCount; i++)
  {
  	var sheet = sheets[i];
  	var sheetName = sheet.getName().toString();
    var sheetNameLength = sheetName.length;
    var substring_hidden_sheets_name = hidden_sheets.map(function(s) { return s.substring(0,sheetNameLength) });
  	if (substring_hidden_sheets_name.indexOf(sheetName) !== -1)
  	{
  		sheet.hideSheet();
  	}
  }
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("Sheets Hidden","Review hidden sheets",ui.ButtonSet.OK);
}

function showSheets() 
{
  for (var i = 0; i < sheetsCount; i++)
  {
  	var sheet = sheets[i];
  	sheet.showSheet();
  }

  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("All sheets unhidden","Review unhidden sheets",ui.ButtonSet.OK);
}

function reorderSheets()
{
  for(var i = 0; i < sheetsCount; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName().toString();
    var sheetNameLength = sheetName.length;
    var substring_tabs_in_order_name = tabs_in_order.map(function(s) { return s.substring(0,sheetNameLength) });
    var order = substring_tabs_in_order_name.indexOf(sheetName);
    Logger.log("Sheet: " + sheetName + "  Order: " + (order+1).toString());
    if (order !== -1)
    {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(order+1);
    }
  }
}