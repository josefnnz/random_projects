function create_separation_agreements()
{
  // Constants
  var EMPTY_STRING = "";
  var CIC = "CIC"
  var OATH = "Oath"
  var TRANS = "Trans"
  var NON_TRANS = "NonTrans"
  var WARN = "WARN"
  var NON_WARN = "NonWARN"

  // Array column indices for required fields
  // NOTE: Array column indices do not match location on ss. SS increments indices by 1.
  //       Issue because SS indices begin at 1. But Array column indices begin at 0.
  var SEPARATION_AGREEMENT_URL_CIDX = 1 - 1;
  var CALIFORNIA_CHANGE_OF_STATUS_URL_CIDX = 2 - 1;
  var EEID_CIDX = 3 - 1;
  var FULL_LEGAL_NAME_CIDX = 4 - 1;
  var LEGAL_FIRST_NAME_CIDX = 5 - 1;
  var HOME_ADDRESS_LINE_1_CIDX = 6 - 1;
  var HOME_ADDRESS_LINE_2_CIDX = 7 - 1;
  var SSN_CIDX = 8 - 1;
  var CALIFORNIA_FLAG_CIDX = 9 - 1;
  var ADEA_FLAG_CIDX = 10 - 1;
  var L2_ORGNAME_CIDX = 11 - 1;
  var TRANSITION_LENGTH_DAYS_CIDX = 12 - 1;
  var NOTIFICATION_DATE_CIDX = 13 - 1;
  var LAST_DAY_OF_WORK_CIDX = 14 - 1;
  var SEPARATION_DATE_CIDX = 15 - 1;
  var SEVERANCE_PLAN_CIDX = 16 - 1;
  var TRANSITION_FLAG_CIDX = 17 - 1;
  var WARN_FLAG_CIDX = 18 - 1;
  var SEPARATION_AGREEMENT_TEMPLATE_CIDX = 19 - 1;
  var DATE_OF_SEPARATION_AGREEMENT_CIDX = 20 - 1;
  var OATH_PLAN_BASE_COMP_PAYOUT_AMOUNT_CIDX = 21 - 1;
  var OATH_PLAN_NUM_WEEKS_OF_BASE_COMP_PAYOUT_CIDX = 22 - 1;
  var OATH_PLAN_NUM_MONTHS_OF_COBRA_CIDX = 23 - 1;
  var CIC_PLAN_SALARY_CONTINUATION_MONTHS_INCLUDING_NOTICE_PERIOD_CIDX = 24 - 1;
  var CIC_PLAN_SALARY_CONTINUATION_MONTHS_MINUS_NOTICE_PERIOD_CIDX = 25 - 1;
  var CIC_PLAN_NUM_MONTHS_OF_COBRA_CIDX = 26 - 1;
  var TRANSITION_BONUS_AMOUNT_CIDX = 27 - 1;
  var OAB_OR_SIP_OR_SEIP_CIDX = 28 - 1;
  var RETENTION_FLAG_CIDX = 29 - 1;
  var L2_CIDX = 30 - 1;
  var WORK_LOCATION_CIDX = 31 - 1;
  var IS_SEPARATION_DATE_IN_2017_CIDX = 32 = 1;

  // Separation Agreement Type to Template Google ID
  var mapping = {"CIC - NonWARN - NonTrans"  : "1O5kxLuWHw2nd_ZAyArt0HQ4gVqTuXnNt4KaXKyXXbGk",
                 "CIC - WARN - NonTrans"     : "1hMZdS1CofMwPRQdgsD56H-v3WAyWlsVW9wBlqNY78uM",
                 "CIC - NonWARN - Trans"     : "16MhqPc-_i1N2L6y0Y_VcaWAKD3_M7HLKrH55ZwAsoDk", // CIC Transition Templates are the same
                 "CIC - WARN - Trans"        : "16MhqPc-_i1N2L6y0Y_VcaWAKD3_M7HLKrH55ZwAsoDk", // CIC Transition Templates are the same
                 "Oath - NonWARN - NonTrans" : "1CTt50ptABrfaGRVO4s9U_rx5E-8CLt0FE0W41-LJQYg",
                 "Oath - WARN - NonTrans"    : "1iFnkDBwNGMQjlx7HcLyJW3-tUCUoEEntZ29-RbRtaR4", // Oath Transition Templates are the same
                 "Oath - NonWARN - Trans"    : "1UCmLfzjnLg_qS0PQNxP-eWxOzOsbqtImgFrRybOvi1E", // Oath Transition Templates are the same
                 "Oath - WARN - Trans"       : "1UCmLfzjnLg_qS0PQNxP-eWxOzOsbqtImgFrRybOvi1E"};

  // // Separation Agreement Type to Template Google ID
  // var mapping = {CIC  + " - " + NON_TRANS + " - " + NON_WARN : "1O5kxLuWHw2nd_ZAyArt0HQ4gVqTuXnNt4KaXKyXXbGk",
  //                CIC  + " - " + NON_TRANS + " - " + WARN     : "1hMZdS1CofMwPRQdgsD56H-v3WAyWlsVW9wBlqNY78uM",
  //                CIC  + " - " + TRANS     + " - " + NON_WARN : "16MhqPc-_i1N2L6y0Y_VcaWAKD3_M7HLKrH55ZwAsoDk", // CIC Transition Templates are the same
  //                CIC  + " - " + TRANS     + " - " + WARN     : "16MhqPc-_i1N2L6y0Y_VcaWAKD3_M7HLKrH55ZwAsoDk", // CIC Transition Templates are the same
  //                OATH + " - " + NON_TRANS + " - " + NON_WARN : "1CTt50ptABrfaGRVO4s9U_rx5E-8CLt0FE0W41-LJQYg",
  //                OATH + " - " + NON_TRANS + " - " + WARN     : "1iFnkDBwNGMQjlx7HcLyJW3-tUCUoEEntZ29-RbRtaR4", // Oath Transition Templates are the same
  //                OATH + " - " + TRANS     + " - " + NON_WARN : "1UCmLfzjnLg_qS0PQNxP-eWxOzOsbqtImgFrRybOvi1E", // Oath Transition Templates are the same
  //                OATH + " - " + TRANS     + " - " + WARN     : "1UCmLfzjnLg_qS0PQNxP-eWxOzOsbqtImgFrRybOvi1E"};


  // Google file ids
  var SEPARATION_AGREEMENTS_FOLDER_ID = "1U7_gHXUqoPCCh7sh-RziqGov6fYZAlOy"; // Folder: separation_agreements
  var RIFS_SSID = "1YaUHm_G5O72Twd9EfApfUwAAcLVPWhDRgFGi48OLLZ4"; // File: Project R2 - USA Calculations and Agreement Generator
  var RIFS_SHN = "Calcs";

  // Set folder where Separation Agreements will be created
  var folder = DriveApp.getFolderById(SEPARATION_AGREEMENTS_FOLDER_ID);

  // Load sheet with impacted employees
  var ees = SpreadsheetApp.openById(RIFS_SSID).getSheetByName(RIFS_SHN);

  // Starting and Ending Rows of Table
  var FIRST_ROW_EXTRACTED = 1 * ees.getSheetValues(1, 2, 1, 1);
  var LAST_ROW_EXTRACTED = 1 * ees.getSheetValues(2, 2, 1, 1);

  // Identify total number of rows and columns to extract
  var NUM_ROWS_TO_EXTRACT = LAST_ROW_EXTRACTED - FIRST_ROW_EXTRACTED + 1;
  var NUM_COLS_TO_EXTRACT = 31; // Columns A - AC

  // Extract range of employee data starting with first employee row -- EXCLUDE HEADER ROWS
  var values_ees = ees.getRange(FIRST_ROW_EXTRACTED, 1, NUM_ROWS_TO_EXTRACT, NUM_COLS_TO_EXTRACT).getValues();
  var NUM_EES = values_ees.length;

  function mail_merge()
  {
    // Confirm user wants to run script
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Please check cells B1 and B2 and confirm they capture the first and last employees on the spreadsheet. Click 'Ok' to continue to run the script. Click 'Cancel' or exit the prompt to kill the script.", ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) 
    {
     return;
    }

    for (var row = 0; row < NUM_EES; row++)
    {
      // Extract current employee
      var curr = values_ees[row];

      // Extract required fields
      var date_of_agreement = curr[DATE_OF_SEPARATION_AGREEMENT_CIDX];
      var eeid = curr[EEID_CIDX];
      var full_legal_name = curr[FULL_LEGAL_NAME_CIDX];
      var legal_first_name = curr[LEGAL_FIRST_NAME_CIDX];
      var home_address_line_1 = curr[HOME_ADDRESS_LINE_1_CIDX];
      var home_address_line_2 = curr[HOME_ADDRESS_LINE_2_CIDX];
      var last_day_of_work = curr[LAST_DAY_OF_WORK_CIDX];
      var separation_date = curr[SEPARATION_DATE_CIDX];
      var severance_plan = curr[SEVERANCE_PLAN_CIDX];
      var has_cic_plan = severance_plan == CIC;
      var has_oath_plan = severance_plan == OATH;
      var transition_flag = curr[TRANSITION_FLAG_CIDX];
      var is_transition = transition_flag == TRANS;
      var cic_salary_cont_mths_including_notice_period = curr[CIC_PLAN_SALARY_CONTINUATION_MONTHS_INCLUDING_NOTICE_PERIOD_CIDX];
      var cic_salary_cont_mths_minus_notice_period = curr[CIC_PLAN_SALARY_CONTINUATION_MONTHS_MINUS_NOTICE_PERIOD_CIDX];
      var transition_bonus_amount = curr[TRANSITION_BONUS_AMOUNT_CIDX];
      var oath_base_compensation_payout_amount = curr[OATH_PLAN_BASE_COMP_PAYOUT_AMOUNT_CIDX];
      var oath_weeks_of_base_compensation = curr[OATH_PLAN_NUM_WEEKS_OF_BASE_COMP_PAYOUT_CIDX];
      var cic_plan_months_of_cobra_coverage = curr[CIC_PLAN_NUM_MONTHS_OF_COBRA_CIDX];
      var oath_plan_months_of_cobra_coverage = curr[OATH_PLAN_NUM_MONTHS_OF_COBRA_CIDX];
      var oab_or_sip_seip_bonus_plan = curr[OAB_OR_SIP_OR_SEIP_CIDX];
      var has_nonsales_bonus_plan = oab_or_sip_seip_bonus_plan == "OAB";
      var has_sales_nonseip_bonus_plan = oab_or_sip_seip_bonus_plan == "SIP";
      var has_seip_bonus_plan = oab_or_sip_seip_bonus_plan == "SEIP";

      var has_outstanding_retention_bonuses = curr[RETENTION_FLAG_CIDX] == "RET";

      var is_separation_date_in_2017 = curr[IS_SEPARATION_DATE_IN_2017_CIDX] == "Y";

      var sep_agmt_tmpl = curr[SEPARATION_AGREEMENT_TEMPLATE_CIDX]
      var adea_flag = curr[ADEA_FLAG_CIDX];
      var L2 = curr[L2_CIDX];
      var california_flag = curr[CALIFORNIA_FLAG_CIDX];
      var work_location = curr[WORK_LOCATION_CIDX];

      // Copy the template
      var filename = "AGMT - " + work_location + " - " + sep_agmt_tmpl + " - " + adea_flag + " - " + L2 + " - " + california_flag + " - " + full_legal_name + " (" + eeid + ")";
      var term_notice_tmpl_id = mapping[sep_agmt_tmpl];
      var file_new_ee_doc = DriveApp.getFileById(term_notice_tmpl_id).makeCopy(filename, folder);

      // Fil-in copy with employee details
      var doc_new_ee_doc = DocumentApp.openById(file_new_ee_doc.getId());
      var body = doc_new_ee_doc.getBody();

      body.replaceText("<<date_of_agreement>>", date_of_agreement);
      body.replaceText("<<full_legal_name>>", full_legal_name);
      body.replaceText("<<home_address_line_1>>", home_address_line_1);
      body.replaceText("<<home_address_line_2>>", home_address_line_2);
      body.replaceText("<<legal_first_name>>", legal_first_name);
      body.replaceText("<<last_day_of_work>>", last_day_of_work);
      body.replaceText("<<separation_date>>", separation_date);
      body.replaceText("<<employee_id>>", eeid);

      if (has_cic_plan)
      {
        body.replaceText("<<salary_continuation_months_including_notice_period>>", cic_salary_cont_mths_including_notice_period);
        body.replaceText("<<salary_continuation_months_minus_notice_period>>", cic_salary_cont_mths_minus_notice_period);
        body.replaceText("<<cic_plan_months_of_cobra_coverage>>", cic_plan_months_of_cobra_coverage);
      }
      else if (has_oath_plan)
      {
        body.replaceText("<<base_compensation_payout_amount>>", oath_base_compensation_payout_amount);
        body.replaceText("<<weeks_of_base_compensation>>", oath_weeks_of_base_compensation);
        body.replaceText("<<oath_plan_months_of_cobra_coverage>>", oath_plan_months_of_cobra_coverage);
      }

      if (is_transition)
      {
        body.replaceText("<<transition_bonus_amount>>", transition_bonus_amount);
      }

      if (has_nonsales_bonus_plan)
      {
        delete_section_tag(body, "<<corporate_bonus_section_block>>");      
        delete_section_block(body, "<<sales_bonus_section_block>>");
        delete_section_block(body, "<<seip_bonus_section_block>>");
        if (is_separation_date_in_2017)
        {
          delete_section_block(body, "<<2017_corporate_bonus_section_block>>");
        }
      }
      else if (has_sales_nonseip_bonus_plan)
      {
        delete_section_tag(body, "<<sales_bonus_section_block>>");
        delete_section_block(body, "<<corporate_bonus_section_block>>");
        delete_section_block(body, "<<2017_corporate_bonus_section_block>>");
        delete_section_block(body, "<<seip_bonus_section_block>>");
      }
      else if (has_seip_bonus_plan)
      {
        delete_section_tag(body, "<<seip_bonus_section_block>>");
        delete_section_block(body, "<<corporate_bonus_section_block>>");
        delete_section_block(body, "<<2017_corporate_bonus_section_block>>");
        delete_section_block(body, "<<sales_bonus_section_block>>");
      }

      if (has_outstanding_retention_bonuses)
      {
        delete_section_tag(body, "<<retention_bonuses_section_block>>")
      }
      else
      {
        delete_section_block(body, "<<retention_bonuses_section_block>>")
      }

      doc_new_ee_doc.saveAndClose();

      // Save memo as pdf and delete Google Doc version
      var pdf_version = folder.createFile(file_new_ee_doc.getAs("application/pdf"));
      pdf_version.setName(filename);
      file_new_ee_doc.setTrashed(true);

      // Write unique URL for new file
      ees.getRange(row+FIRST_ROW_EXTRACTED, 1, 1, 1).setValue("https://drive.google.com/a/oath.com/file/d/" + pdf_version.getId() + "/view?usp=sharing");
      SpreadsheetApp.flush();
    }
  }

  /**
    * Remove section block associated with given tag
    *
    * @param body - class Body object containing section tag
    * @param section_tag - section tag to remove
    *
    * @return deletes section block
    */
  function delete_section_block(body, section_tag) 
  {
    var range_element = body.findText(section_tag);
    // If section tag was found, remove section from document
    if (range_element) 
    {
      range_element.getElement().getParent().removeFromParent();
    }
  }

  /**
    * Remove section tag from section block. Use this function when you want to maintain the section block, and remove the mail merge tag
    *
    * @param body - class Body object containing section tag
    * @param section_tag - section tag to remove
    *
    * @return deletes section tag
    */
  function delete_section_tag(body, section_tag)
  {
    body.replaceText(section_tag, EMPTY_STRING);
  }

  mail_merge();
}
  // <<date_of_agreement>>
  // <<full_legal_name>>
  // <<home_address_line_1>>
  // <<home_address_line_2>>
  // <<legal_first_name>>
  // <<last_day_of_work>>
  // <<separation_date>>
  // <<salary_continuation_months_including_notice_period>>
  // <<salary_continuation_months_minus_notice_period>>
  // <<transition_bonus_amount>>
  // <<base_compensation_payout_amount>>
  // <<weeks_of_base_compensation>>
  // <<corporate_bonus_section_block>>
  // <<sales_bonus_section_block>>
  // <<retention_bonuses_section_block>>
  // <<oath_plan_months_of_cobra_coverage>>
  // <<cic_plan_months_of_cobra_coverage>>
  // <<employee_id>>







