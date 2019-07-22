function create_bring_to_target_statements()
{
  // Google file ids
  var GOOGLE_ID_FOLDER_IVAN_MARKMAN = "1OhZUxrkJ3p8eGJW8GMQdxGx8Tcdt3NFO"; // Folder: Ivan Markman - 2019 Bring to Target Statements
  var GOOGLE_ID_FOLDER_ROSE_TSOU = "1HwbJ46ZIeZMQXB8KjLa-cJpeBrlbdwr5"; // Folder: Rose Tsou - 2019 Bring to Target Statements
  var GOOGLE_ID_FOLDER_KELLY_LIANG = "1IH6IsH1Ku0A-XYIPeXpfalwhLsKIAitv"; // Folder: Kelly Liang - 2019 Bring to Target Statements
  var GOOGLE_ID_FOLDER_JOANNA_LAMBERT = "1OF9oMk4sQoAlyAShrvLv7iI9-qJyYY5W"; // Folder: Joanna Lambert - 2019 Bring to Target Statements
  var GOOGLE_ID_HOURLY_TEMPLATE = "1PLQJabq_zSTmj_Li4wN6lFHvjnav4bVA4nQ3tsJPNnc";
  var GOOGLE_ID_SALARIED_TEMPLATE = "1mmroG97XTDEOwyG5h3jivctqxFJiYaLLNbg6Tl9JwNQ";
  var GOOGLE_ID_SS_MAIL_MERGE_DATA_SOURCE = "18bAU9MnCKGdsK8GPHAXLkIluXd1Syb0RmPIckOne5js"; // File: Mail Merge Data Source - 2019 Bring to Target Statements
  var SHN_MAIL_MERGE_DATA_SOURCE = "MailMergeData";

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var PREFERRED_FIRST_NAME_CIDX = 1 - 1;
  var PREFERRED_FULL_NAME_CIDX = 2 - 1;
  var EMPLOYEE_ID_CIDX = 3 - 1;
  var HOURLY_OR_SALARIED_CIDX = 4 - 1;
  var LOCAL_CURRENCY_CODE_CIDX = 5 - 1;
  var CURRENT_BASE_SALARY_CIDX = 6 - 1;
  var CURRENT_HOURLY_RATE_CIDX = 7 - 1;
  var CURRENT_BONUS_TARGET_PCT_CIDX = 8 - 1;
  var CURRENT_BONUS_TARGET_AMT_CIDX = 9 - 1;
  var CURRENT_TTC_CIDX = 10 - 1;
  var NEW_BASE_SALARY_CIDX = 11 - 1;
  var NEW_HOURLY_RATE_CIDX = 12 - 1;
  var NEW_BONUS_TARGET_PCT_CIDX = 13 - 1;
  var NEW_BONUS_TARGET_AMT_CIDX = 14 - 1;
  var NEW_TTC_CIDX = 15 - 1;
  var TTC_INCREASE_PCT_CIDX = 16 - 1;
  var DIRECT_MGR_CIDX = 17 - 1;
  var L2_CIDX = 18 - 1;
  var L3_CIDX = 19 - 1;
  var L4_CIDX = 20 - 1;

  // Set folders where statements will be created
  var folder_ivan_markman = DriveApp.getFolderById(GOOGLE_ID_FOLDER_IVAN_MARKMAN);
  var folder_rose_tsou = DriveApp.getFolderById(GOOGLE_ID_FOLDER_ROSE_TSOU);
  var folder_kelly_liang = DriveApp.getFolderById(GOOGLE_ID_FOLDER_KELLY_LIANG);
  var folder_joanna_lambert = DriveApp.getFolderById(GOOGLE_ID_FOLDER_JOANNA_LAMBERT);

  // Set mail merge templates
  var gdoc_template_hourly = DriveApp.getFileById(GOOGLE_ID_HOURLY_TEMPLATE);
  var gdoc_template_salaried = DriveApp.getFileById(GOOGLE_ID_SALARIED_TEMPLATE);

  // Load sheet with bring-to-target employees
  var ees = SpreadsheetApp.openById(GOOGLE_ID_SS_MAIL_MERGE_DATA_SOURCE).getSheetByName(SHN_MAIL_MERGE_DATA_SOURCE);

  // Starting and Ending Rows of Table
  var FIRST_COL_OF_DATA = 1;
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetValues(1, 22, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetValues(2, 22, 1, 1);

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 20; // Columns U - AN

  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, FIRST_COL_OF_DATA, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  function mail_merge()
  {
    // Confirm user wants to run script
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Please check cells W1 and W2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) 
    {
     return;
    }

    for (var row = 0; row < NUM_EES; row++)
    {
      // Extract current employee
      var curr = values_ees[row];

      // Extract required fields
      var preferred_first_name = curr[PREFERRED_FIRST_NAME_CIDX];
      var preferred_full_name = curr[PREFERRED_FULL_NAME_CIDX];
      var employee_id = curr[EMPLOYEE_ID_CIDX];
      var hourly_or_salaried = curr[HOURLY_OR_SALARIED_CIDX];
      var local_currency_code = curr[LOCAL_CURRENCY_CODE_CIDX];
      var current_base_salary = curr[CURRENT_BASE_SALARY_CIDX];
      var current_hourly_rate = curr[CURRENT_HOURLY_RATE_CIDX];
      var current_bonus_target_pct = curr[CURRENT_BONUS_TARGET_PCT_CIDX];
      var current_bonus_target_amt = curr[CURRENT_BONUS_TARGET_AMT_CIDX];
      var current_ttc = curr[CURRENT_TTC_CIDX];
      var new_base_salary = curr[NEW_BASE_SALARY_CIDX];
      var new_hourly_rate = curr[NEW_HOURLY_RATE_CIDX];
      var new_bonus_target_pct = curr[NEW_BONUS_TARGET_PCT_CIDX];
      var new_bonus_target_amt = curr[NEW_BONUS_TARGET_AMT_CIDX];
      var new_ttc = curr[NEW_TTC_CIDX];
      var ttc_increase_pct = curr[TTC_INCREASE_PCT_CIDX];
      var direct_manager = curr[DIRECT_MGR_CIDX];
      var L2 = curr[L2_CIDX];
      var L3 = curr[L3_CIDX];
      var L4 = curr[L4_CIDX];

      // Set correct folder and statement template for employee
      var folder = (L2 === "Rose Tsou") ? folder_rose_tsou : folder_ivan_markman;
      var folder = (L2 === "Kelly Liang") ? folder_kelly_liang : folder;
      var folder = (L2 === "Joanna Lambert") ? folder_joanna_lambert : folder;
      var template = (hourly_or_salaried === "Hourly") ? gdoc_template_hourly : gdoc_template_salaried;

      // Copy the template
      var filename = L3 + " - " + L4 + " - " + direct_manager + " - " + preferred_full_name + " (" + employee_id + ") Bonus Target Change";
      var file_new_ee_doc = template.makeCopy(filename, folder);

      // Fil-in copy with employee details
      var doc_new_ee_doc = DocumentApp.openById(file_new_ee_doc.getId());
      var body = doc_new_ee_doc.getBody();

      body.replaceText("<Preferred_First_Name>", preferred_first_name);
      body.replaceText("<Preferred_Full_Name>", preferred_full_name);
      body.replaceText("<Employee_ID>", employee_id);
      body.replaceText("<Current_Base_Salary>", current_base_salary);
      body.replaceText("<Current_Hourly_Rate>", current_hourly_rate);
      body.replaceText("<Current_Bonus_Target_%>", current_bonus_target_pct);
      body.replaceText("<Current_Bonus_Target_Amount>", current_bonus_target_amt);
      body.replaceText("<Current_TTC>", current_ttc);
      body.replaceText("<New_Base_Salary>", new_base_salary);
      body.replaceText("<New_Hourly_Rate>", new_hourly_rate);
      body.replaceText("<New_Bonus_Target_%>", new_bonus_target_pct);
      body.replaceText("<New_Bonus_Target_Amount>", new_bonus_target_amt);
      body.replaceText("<New_TTC>", new_ttc);
      body.replaceText("<TTC_Increase_%>", ttc_increase_pct);

      doc_new_ee_doc.saveAndClose();

      // Save memo as pdf and delete Google Doc version
      var pdf_version = folder.createFile(file_new_ee_doc.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_ee_doc.setTrashed(true);

      // Write unique URL for new file
      ees.getRange(row+FIRST_ROW_EXTRACTED, 41, 1, 1).setValue("https://drive.google.com/a/oath.com/file/d/" + pdf_version.getId() + "/view?usp=sharing");
      SpreadsheetApp.flush();
    }
  }

  mail_merge();
}