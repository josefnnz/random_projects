function sendNonOvpEmails()
{
  var SSID_MASS_EMAIL_GENERATOR = "1ThDHIjT16zzG28Hbuk2DKt7XhgU7h0NWge45mTGOeJY";
  var STNM_NON_OVP_EMAILS_TAB = "Non_OVP_Emails";

  var TMPL_APPROVED_PROMO = "1gE8TjWfxsfA9hlkAbtTPx32Y0bDx3pMmWKRY3mf9QuE";
  var TMPL_DENIED_PROMO_WITH_REASON = "1h2omPaNzF5N58WEFl52osoEBZMw7JRqBFhGHiUJUmmU";
  var TMPL_DENIED_PROMO_WITHOUT_REASON = "1GZllGJX7slzShCXytc-4eTgEfI0LSAmT26tSBQUR67o";

  var FIRST_COL_OF_DATA = 1;
  var LAST_COL_OF_DATA = 9;

  // Column indices
  var L2_DECISION_CIDX = 1 - 1;
  var L2_DENIAL_REASON_CIDX = 2 - 1;
  var NOMINEE_FULL_NAME_CIDX = 3 - 1;
  var NOMINEE_EEID_CIDX = 4 - 1;
  var MGR_FULL_NAME_CIDX = 5 - 1;
  var MGR_FIRST_NAME_CIDX = 6 - 1;
  var MGR_EMAIL_CIDX = 7 - 1;
  var L2_FULL_NAME_CIDX = 8 - 1;
  var L3_FULL_NAME_CIDX = 9 - 1;
  var L3_EMAIL_CIDX = 10 - 1;
  var HRBP_FULL_NAME_CIDX = 11 - 1;
  var HRBP_EMAIL_CIDX = 12 - 1;
  var CC_LINE_CIDX = 13 - 1;
  var EMAIL_SEND_STATUS_CIDX = 14 - 1;

  // Constants
  var L2_APPROVE_DECISION_VALUE = "YES";

  // This constant is written in column E for rows for which an email
  // has been sent successfully.
  var EMAIL_SENT = "EMAIL_SENT";

  function getGoogleDocAsHTML(id)
  {
    var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
    var param = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
    };
    var html = UrlFetchApp.fetch(url,param).getContentText();
    return html;
  }

  function sendEmails() 
  {   
    var response = Browser.msgBox("You are about to send emails to EVERYONE ON SHEET " + STNM_NON_OVP_EMAILS_TAB + " (WHAAATTTT?????)."
                +"Press OK if you are confident you wont break corpmail",Browser.Buttons.OK_CANCEL);
   
    var response2 = Browser.msgBox("are you sure you're really sure? if yes, press ok. if no, press cancel and do a dance...you saved corpmail!",Browser.Buttons.OK_CANCEL);
   
    var response3 = Browser.msgBox("WOW....you are really persistent. last chance. are you really really really sure???",Browser.Buttons.OK_CANCEL);
   
    // Load sheet with nominated Non-OVP promotions
    var sheet = SpreadsheetApp.openById(SSID_MASS_EMAIL_GENERATOR).getSheetByName(STNM_NON_OVP_EMAILS_TAB);
    var FIRST_ROW_OF_DATA = 1 * sheet.getSheetValues(1, 2, 1, 1);
    var LAST_ROW_OF_DATA = 1 * sheet.getSheetValues(2, 2, 1, 1);

    var NUM_ROWS = LAST_ROW_OF_DATA - FIRST_ROW_OF_DATA + 1;
    var NUM_COLS = LAST_COL_OF_DATA - FIRST_COL_OF_DATA + 1;

    if (response == "ok" & response2 == "ok" & response3 == "ok") 
    {
      // Get HTML for email templates
      var tmpl_approved = getGoogleDocAsHTML(TMPL_APPROVED_PROMO);
      var tmpl_denied_with_reason = getGoogleDocAsHTML(TMPL_DENIED_PROMO_WITH_REASON);
      var tmpl_denied_without_reason = getGoogleDocAsHTML(TMPL_DENIED_PROMO_WITHOUT_REASON);
      var subject_line_approved_promo = "Promotion Request for EMPLOYEE_PREFERRED_NAME Approved";
      var subject_line_denied_promo = "Promotion Request for EMPLOYEE_PREFERRED_NAME Denied";

      // Get list of nominees
      var data = sheet.getRange(FIRST_ROW_OF_DATA, FIRST_COL_OF_DATA, NUM_ROWS, NUM_COLS).getValues();

      for (i = 0; i < data.length; i++) 
      {
        var row = data[i];

        var L2_decision = row[L2_DECISION_CIDX];
        var L2_denial_reason = row[L2_DENIAL_REASON_CIDX];
        var nominee = row[NOMINEE_FULL_NAME_CIDX];
        var mgr_first_name = row[MGR_FIRST_NAME_CIDX];
        var mgr_email = row[MGR_EMAIL_CIDX];
        var L2_name = row[L2_FULL_NAME_CIDX];
        // var cc_line = row[CC_LINE_CIDX];

        var message = null;
        var subject = null;
        if (L2_decision == L2_APPROVE_DECISION_VALUE)
        {
          message = tmpl_approved;
          subject = subject_line_approved_promo;
        } 
        else
        {
          subject = subject_line_denied_promo;
          if (L2_denial_reason)
          {
            message = tmpl_denied_with_reason;
          } 
          else
          {
            message = tmpl_denied_without_reason;
          }
        }

        message = message.replace("MANAGER_PREFERRED_FIRST_NAME", mgr_first_name);
        message = message.replace("EMPLOYEE_PREFERRED_NAME", nominee);
        message = message.replace("L2_PREFERRED_NAME", L2_name);
        message = message.replace("DENIAL_REASON", L2_denial_reason);
        subject = subject.replace("EMPLOYEE_PREFERRED_NAME", nominee);
       
        MailApp.sendEmail(mgr_email, subject, message, {
          htmlBody: message,
          bcc: "paquino@oath.com,sanj@oath.com"
        });
       
        // Write "EMAIL_SENT" in last column to confirm email delivery
        sheet.getRange(FIRST_ROW_OF_DATA+i, EMAIL_SEND_STATUS_CIDX+1).setValue(EMAIL_SENT);
        SpreadsheetApp.flush();
      }
    }
  }

 sendEmails(); 
}