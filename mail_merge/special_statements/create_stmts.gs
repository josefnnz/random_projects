function create_stmts()
{

  // Confirm user wants to run script
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Please check cells B1 and B2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to prevent the script from running.", ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
   return;
  }

  // Google file ids
  var BACKUP_GDOCS_FOLDER_ID = "1D6iyw3fPMVo6Erep3PZlQ4WfzPx6p1bm"; // Folder: 
  var TMPL_ID = "1ipPSqP08lenNyRFAYHkJfaqHKYRyLLHydk-H24aDLLc"; // Template: 
  var SSID = "1xTNM6Od7GA39UQagU659fLB4Mluau8_RyLlVSK-7CGw"; // Spreadsheet: 
  var SHN = "create_stmts"; // Sheet Name: 

  // Get folder
  var backup_gdocs_folder = DriveApp.getFolderById(BACKUP_GDOCS_FOLDER_ID);

  // Get sheet
  var sheet_ees = SpreadsheetApp.openById(SSID).getSheetByName(SHN);

  // Identify first and last rows to extract
  var FIRST_ROW_EXTRACTED = 1 * sheet_ees.getSheetValues(1, 2, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * sheet_ees.getSheetValues(2, 2, 1, 1);

  // Identify number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 40; // Columns C to AP

  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = sheet_ees.getRange(FIRST_ROW_EXTRACTED, 3, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var EEID_CIDX = 1 - 1;
  var PREFERRED_FULL_NAME_CIDX = 2 - 1;
  var LEGAL_FULL_NAME_CIDX = 3 - 1;
  var LEGAL_FIRST_NAME_CIDX = 4 - 1;
  var MANAGER_EEID_CIDX = 5 - 1;
  var MANAGER_PREFERRED_NAME_CIDX = 6 - 1;
  var MANAGER_WORK_EMAIL_CIDX = 7 - 1;
  var L2_CIDX = 8 - 1;
  var WORK_LOCATION_COUNTRY_CIDX = 9 - 1;
  var WORK_LOCATION_REGION_CIDX = 10 - 1;
  var IS_PROMO_CIDX = 11 - 1;
  var CURRENT_JOB_PROFILE_CIDX = 12 - 1;
  var CURRENT_JOB_LEVEL_CIDX = 13 - 1;
  var NEW_JOB_PROFILE_CIDX = 14 - 1;
  var NEW_JOB_LEVEL_CIDX = 15 - 1;
  var LOCAL_CURRENCY_CIDX = 16 - 1;
  var IS_HOURLY_EE_CIDX = 17 - 1;
  var BASE_SALARY_DEC_CIDX = 18 - 1;
  var BONUS_TARGET_PCT_DEC_CIDX = 19 - 1;
  var TTC_DEC_CIDX = 20 - 1;
  var HOURLY_RATE_DEC_CIDX = 21 - 1;
  var BASE_SALARY_JAN_CIDX = 22 - 1;
  var BONUS_TARGET_PCT_JAN_CIDX = 23 - 1;
  var TTC_JAN_CIDX = 24 - 1;
  var HOURLY_RATE_JAN_CIDX = 25 - 1;
  var BASE_PCT_INC_DEC_TO_JAN_CIDX = 26 - 1;
  var MAKE_WHOLE_INC_CIDX = 27 - 1;
  var BASE_SALARY_AFTER_MERIT_CIDX = 28 - 1;
  var BONUS_TARGET_PCT_AFTER_MERIT_CIDX = 29 - 1;
  var TTC_AFTER_MERIT_CIDX = 30 - 1;
  var HOURLY_RATE_AFTER_MERIT_CIDX = 31 - 1;
  var BASE_PCT_INC_JAN_TO_AFTER_MERIT_CIDX = 32 - 1;
  var OVERALL_BASE_INC_CIDX = 33 - 1;
  var MERIT_EFFECTIVE_DATE_CIDX = 34 - 1;
  var IS_AWARDED_EQUITY_CIDX = 35 - 1;
  var EQUITY_VALUE_CIDX = 36 - 1;
  var EMPLOYMENT_AGMT_DATE_CIDX = 37 - 1;
  var OAB_OR_SIP_CIDX = 38 - 1;
  var HRA_RECEIVER_CIDX = 39 - 1;
  var DATE_DEADLINE_TO_RETURN_EMPLOYMENT_AGMT_CIDX = 40 - 1;
  var SIGNATURE_NAME_CIDX = 41 - 1;
  var LEGAL_ENTITY_CIDX = 42 - 1;

  // Create mapping of Region + L2 to Folder ID. Statements will be saved in regional L2 folders.
  var region_L2_folder_ids = {"EMEA - Allie Kline" : "1DlsDzzIkB5tZ-sjT1a__7HBweiO-koRg",
                              "EMEA - Atte Lahtiranta" : "1oeRVR_pMCzOmgeA2OkNIVJyzNV1t2XJW",
                              "EMEA - Bob Toohey" : "1uXsHKa3sUuUAHOMYyHxYZpWgNhGSFclS",
                              "EMEA - John DeVine" : "10C27d1ZpZ9rLa48PBVgjr6o74mz_rYjm",
                              "EMEA - Julie Jacobs" : "10viqPfwwcsGoW0b47jzvUhw-HOdszFrr",
                              "EMEA - Rohit Chandra" : "1iwaqAEbjgK4rpgp9NRKzMFSHU_--DVA-",
                              "EMEA - Simon Khalaf" : "1PtDT2Qf03gOa9TWILe4pRJAKF_e6JtQg",
                              "EMEA - Vanessa Wittman" : "1MdA-3q3Hd3nf9iz5XNLVcIvmDNx3jSi5",
                              "NonEMEA - Allie Kline" : "1xR14p8rcEWdhrCUsYoDo0HUnVEDdbzCZ",
                              "NonEMEA - Jeffrey Bonforte" : "1w_smks3K83smkzRyHIjfxtHCoP2_4Wow",
                              "NonEMEA - John DeVine" : "1VmWu9zjarAgWtA1R_NMEBawRmFqizG33",
                              "NonEMEA - Julie Jacobs" : "1uUopRThBuVxoqKF6udbNjPnheS5ZzdM5",
                              "NonEMEA - Ralf Jacob" : "1Cf0jca4yunNI1P3A1FwNMnx8TOOhOoRA",
                              "NonEMEA - Rohit Chandra" : "1pfVWaVDPYn-ttqgm43Q22ALaa72tXHOu",
                              "NonEMEA - Simon Khalaf" : "1eBM4tk6hygTOFyH4Y7lnFRRuiun725Jx",
                              "NonEMEA - Vanessa Wittman" : "1WRLqH8pxXhHiW5j0M89kkZz_Vc7vYEZW"}

  function mail_merge() 
  {
    for (var row = 0; row < NUM_EES; row++) 
    {
      // Extract current employee
      var curr = values_ees[row];

      // Get required fields
      var ee_preferred_full_name = curr[PREFERRED_FULL_NAME_CIDX];
      var eeid = curr[EEID_CIDX];
      var mgr_preferred_full_name = curr[MANAGER_PREFERRED_NAME_CIDX];
      var mgr_work_email = curr[MANAGER_WORK_EMAIL_CIDX];
      var emea_nonemea = (curr[WORK_LOCATION_REGION_CIDX].toLowerCase() === "EMEA".toLowerCase()) ? "EMEA" : "NonEMEA";
      var L2 = curr[L2_CIDX];
      var is_promo = curr[IS_PROMO_CIDX] === "Y";
      var curr_job_profile = curr[CURRENT_JOB_PROFILE_CIDX];
      var curr_job_level = curr[CURRENT_JOB_LEVEL_CIDX];
      var new_job_profile = curr[NEW_JOB_PROFILE_CIDX];
      var new_job_level = curr[NEW_JOB_LEVEL_CIDX];
      var is_hourly = curr[IS_HOURLY_EE_CIDX] === "Y";
      var salary_dec = curr[BASE_SALARY_DEC_CIDX];
      var bonus_pct_dec = curr[BONUS_TARGET_PCT_DEC_CIDX];
      var ttc_dec = curr[TTC_DEC_CIDX];
      var hourly_rt_dec = curr[HOURLY_RATE_DEC_CIDX];
      var salary_jan = curr[BASE_SALARY_JAN_CIDX];
      var bonus_pct_jan = curr[BONUS_TARGET_PCT_JAN_CIDX];
      var ttc_jan = curr[TTC_JAN_CIDX];
      var hourly_rt_jan = curr[HOURLY_RATE_JAN_CIDX];
      var salary_merit = curr[BASE_SALARY_AFTER_MERIT_CIDX];
      var bonus_pct_merit = curr[BONUS_TARGET_PCT_AFTER_MERIT_CIDX];
      var ttc_merit = curr[TTC_AFTER_MERIT_CIDX];
      var hourly_rt_merit = curr[HOURLY_RATE_AFTER_MERIT_CIDX];
      var merit_eff_date = curr[MERIT_EFFECTIVE_DATE_CIDX];
      var base_pct_inc_jan = curr[BASE_PCT_INC_DEC_TO_JAN_CIDX];
      var base_pct_inc_merit = curr[BASE_PCT_INC_JAN_TO_AFTER_MERIT_CIDX];
      var make_whole_inc = curr[MAKE_WHOLE_INC_CIDX];
      var overall_base_inc = curr[OVERALL_BASE_INC_CIDX];
      var is_awarded_equity = curr[IS_AWARDED_EQUITY_CIDX] === "Y";
      var equity_amt = curr[EQUITY_VALUE_CIDX];
      var is_non_usa = curr[WORK_LOCATION_COUNTRY_CIDX].toLowerCase() !== "United States of America".toLowerCase()
      var legal_first_name = curr[LEGAL_FIRST_NAME_CIDX];
      var entity = curr[LEGAL_ENTITY_CIDX];
      var legal_full_name = curr[LEGAL_FULL_NAME_CIDX];

      // Get Google ID for Region+L2 folder to save statement into
      var target_folder = DriveApp.getFolderById(region_L2_folder_ids[emea_nonemea + " - " + L2]);

      // Create filename for statement
      var filename = mgr_preferred_full_name + " (" + mgr_work_email + ") - Rewards Statement for " + ee_preferred_full_name + " (" + eeid + ")";

      // Copy statement template gdoc. Open new copy.
      var file_tmpl_copy = DriveApp.getFileById(TMPL_ID).makeCopy(filename, backup_gdocs_folder)
      var doc_tmpl_copy = DocumentApp.openById(file_tmpl_copy.getId());
      var body = doc_tmpl_copy.getBody();

      // Get elements for Promo, Base/TTC, and Equity sections. Elements needed to remove inapplicable sections of statement.
      var tables = body.getTables();
      var table_promo = tables[0];
      var table_base_ttc = tables[1];
      var tables_equity = tables[2];
      var tablerow_hourly_rates = table_base_ttc.getRow(4); // Get row with hourly rates
      var header_promo = body.findText("PROMOTION INFORMATION").getElement();
      var header_base_ttc = body.findText("TOTAL TARGET CASH").getElement();
      var header_equity = body.findText("EQUITY AWARD INFORMATION").getElement();
      var range_element_horizontal_rule_promo = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
      var range_element_horizontal_rule_base_ttc = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE, range_element_horizontal_rule_promo);
      var range_element_horizontal_rule_equity = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE, range_element_horizontal_rule_base_ttc);
      var element_horizontal_rule_promo = range_element_horizontal_rule_promo.getElement();
      var element_horizontal_rule_base_ttc = range_element_horizontal_rule_base_ttc.getElement();
      var element_horizontal_rule_equity = range_element_horizontal_rule_equity.getElement();
      var footnote_non_usa_bonus_equity = body.findText("Except as required by local or regional law").getElement(); // Get non-USA footnote
      var footnote_equity = body.findText("Reflects target value of your Verizon equity award on the grant date").getElement(); // Get equity footnote

      if (!is_promo)
      {
        // Remove promotion section if employee is not receiving a promotion
        header_promo.removeFromParent();
        element_horizontal_rule_promo.removeFromParent();
        table_promo.removeFromParent();
      }

      if (!is_hourly)
      {
        // Remove row of hourly rates in Base/TTC section if employee is not an hourly employee
        tablerow_hourly_rates.removeFromParent();
      }

      if (!is_awarded_equity)
      {
        // Remove equity section if employee is not receiving equity
        header_equity.removeFromParent();
        element_horizontal_rule_equity.removeFromParent();
        tables_equity.removeFromParent();
        footnote_equity.removeFromParent();
      }

      if (!is_non_usa)
      {
        // Remove non-USA legal footnote if employee is located in the USA
        footnote_non_usa_bonus_equity.removeFromParent();
      }

      var paragraphs = body.getParagraphs();
      
      for (var k = 0, line_breaks = 0; k < paragraphs.length; k++)
      {
        if (paragraphs[k].findElement(DocumentApp.ElementType.PAGE_BREAK))
        {
          break;
        }
        if(paragraphs[k].getText() !== "") 
        {
          line_breaks = 0;
        } 
        else 
        {
          if (line_breaks === 0) 
          {
            line_breaks++;
          } 
          else if (!paragraphs[k].isAtDocumentEnd()) 
          {
              paragraphs[k].removeFromParent();
          }
        }
      }

      if (!is_non_usa)
      {
        for (k; k<paragraphs.length; k++)
        {
          if (!paragraphs[k].isAtDocumentEnd())
          {
            paragraphs[k].removeFromParent();
          }
        }
      }

      // Merge data fields in statement, as applicable
      body.replaceText("<<EMPLOYEE_FULL_PREFERRED_NAME>>", ee_preferred_full_name);
      body.replaceText("<<EMPLOYEE_ID>>", eeid);
      if (is_promo)
      {
        // Merge in data specific to an employee receiving a promotion
        body.replaceText("<<PROMO_EFFECTIVE_DATE>>", merit_eff_date);
        body.replaceText("<<CURRENT_JOB_PROFILE>>", curr_job_profile);
        body.replaceText("<<CURRENT_JOB_LEVEL>>", curr_job_level);
        body.replaceText("<<NEW_JOB_PROFILE>>", new_job_profile);
        body.replaceText("<<NEW_JOB_LEVEL>>", new_job_level);  
      }
      body.replaceText("<<BASE_TTC_EFFECTIVE_DATE>>", merit_eff_date);
      body.replaceText("<<SALARY_DEC>>", salary_dec);
      body.replaceText("<<BONUS_PCT_DEC>>", bonus_pct_dec);
      body.replaceText("<<TTC_DEC>>", ttc_dec);
      body.replaceText("<<HOURLY_RT_DEC>>", hourly_rt_dec);
      body.replaceText("<<SALARY_JAN>>", salary_jan);
      body.replaceText("<<BONUS_PCT_JAN>>", bonus_pct_jan);
      body.replaceText("<<TTC_JAN>>", ttc_jan);
      body.replaceText("<<HOURLY_RT_JAN>>", hourly_rt_jan);
      body.replaceText("<<SALARY_MERIT>>", salary_merit);
      body.replaceText("<<BONUS_PCT_MERIT>>", bonus_pct_merit);
      body.replaceText("<<TTC_MERIT>>", ttc_merit);
      body.replaceText("<<HOURLY_RT_MERIT>>", hourly_rt_merit);
      body.replaceText("<<BASE_PCT_INC_JAN>>", base_pct_inc_jan);
      body.replaceText("<<BASE_PCT_INC_MERIT>>", base_pct_inc_merit);
      body.replaceText("<<MAKE_WHOLE>>", make_whole_inc);
      body.replaceText("<<OVERALL_BASE_INC>>", overall_base_inc);
      if (is_awarded_equity)
      {
        // Merge in data specific to an employee receiving equity
        body.replaceText("<<EQUITY_VALUE>>", equity_amt);
      }
      if (is_non_usa)
      {
        // Merge in data specific to a non-USA employee receiving the legal employment agreement
        body.replaceText("<<LEGAL_FIRST_NAME>>", legal_first_name);
        body.replaceText("<<SALARY_JAN>>", salary_jan);
        body.replaceText("<<BONUS_PCT_DEC>>", bonus_pct_dec);
        body.replaceText("<<BONUS_PCT_JAN>>", bonus_pct_jan);
        body.replaceText("<<ENTITY_NAME>>", entity);
        body.replaceText("<<LEGAL_FULL_NAME>>", legal_full_name);
      }

      // Save completed gdoc into backup folder of gdocs
      doc_tmpl_copy.saveAndClose();

      // Save gdoc as pdf in the corresponding Region + L2 folder
      var pdf_version = target_folder.createFile(file_tmpl_copy.getAs("application/pdf"));
      pdf_version.setName(filename);
    }
  }
  mail_merge();
}
