// UPDATE THE FOLLOWING FIELDS
var FEEDBACK_FORM_LINK = "https://docs.google.com/forms/d/e/1FAIpQLSdm4tt89BfX9WhwDY6uJdoBEu3ivif3EbIvzaH3dRDkumpWrQ/viewform";
var SUBJECT = "REMINDER: Executive Promotion Feedback for Hugo Gunnarsen due at 3pm PT Today"
var TIME_NEEDED_TO_COMPLETE_SURVEY = "20-30 minutes";
var DATE_AND_TIME_OF_SURVEY_DEADLINE = "<b>3:00pm PT today, Wednesday, February 15</b>";
var EMAIL_SENDER = "Jorina";

// PORTIONS OF EMAIL INDEPENDENT OF FEEDBACK PROVIDER AND NOMINEE DETAILS
var EMAIL_START = "<html><head><meta content=\"text/html; charset=UTF-8\" http-equiv=\"content-type\"><style type=\"text/css\">@import url('https://themes.googleusercontent.com/fonts/css?kit=wAPX1HepqA24RkYW1AuHYA');ol{margin:0;padding:0}table td,table th{padding:0}.c6{border-right-style:solid;padding:2pt 2pt 2pt 2pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:middle;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;background-color:#f9cb9c;border-left-style:solid;border-bottom-width:1pt;width:210pt;border-top-color:#000000;border-bottom-style:solid}.c9{border-right-style:solid;padding:2pt 2pt 2pt 2pt;border-bottom-color:#000000;border-top-width:1pt;border-right-width:1pt;border-left-color:#000000;vertical-align:middle;border-right-color:#000000;border-left-width:1pt;border-top-style:solid;border-left-style:solid;border-bottom-width:1pt;width:210pt;border-top-color:#000000;border-bottom-style:solid}.c4{background-color:#ffffff;color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:10pt;font-family:\"Arial\";font-style:normal}.c7{background-color:#ffffff;color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-size:11pt;font-family:\"Calibri\";font-style:normal}.c5{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left;height:11pt}.c1{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:center}.c11{padding-top:0pt;padding-bottom:0pt;line-height:1.38;orphans:2;widows:2;text-align:left}.c10{color:#000000;font-weight:400;text-decoration:none;vertical-align:baseline;font-family:\"Arial\";font-style:normal}.c8{padding-top:0pt;padding-bottom:0pt;line-height:1.15;orphans:2;widows:2;text-align:left}.c3{border-spacing:0;border-collapse:collapse;margin-right:auto}.c17{background-color:#ffffff;font-family:\"Calibri\";font-weight:400}.c2{background-color:#ffffff;max-width:468pt;padding:72pt 72pt 72pt 72pt}.c0{background-color:#f9cb9c;font-size:10pt}.c15{color:inherit;text-decoration:inherit}.c13{background-color:#ffffff;color:#196ad4}.c16{height:0pt}.c12{font-size:10pt}.c14{font-size:11pt}.title{padding-top:0pt;color:#000000;font-size:26pt;padding-bottom:3pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}.subtitle{padding-top:0pt;color:#666666;font-size:15pt;padding-bottom:16pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}li{color:#000000;font-size:11pt;font-family:\"Arial\"}p{margin:0;color:#000000;font-size:11pt;font-family:\"Arial\"}h1{padding-top:20pt;color:#000000;font-size:20pt;padding-bottom:6pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h2{padding-top:18pt;color:#000000;font-size:16pt;padding-bottom:6pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h3{padding-top:16pt;color:#434343;font-size:14pt;padding-bottom:4pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h4{padding-top:14pt;color:#666666;font-size:12pt;padding-bottom:4pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h5{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;orphans:2;widows:2;text-align:left}h6{padding-top:12pt;color:#666666;font-size:11pt;padding-bottom:4pt;font-family:\"Arial\";line-height:1.15;page-break-after:avoid;font-style:italic;orphans:2;widows:2;text-align:left}</style></head>";
var TABLE_START = "<a id=\"t.c110ea5ecc0fe47b4ac6f703c36a0355ed8f5907\"></a><a id=\"t.0\"></a><table class=\"c3\"><tbody>";
var TABLE_HEADERS = "<tr class=\"c16\"><td class=\"c6\" colspan=\"1\" rowspan=\"1\"><p class=\"c1\"><span class=\"c0\">Feedback requested by L2</span></p></td><td class=\"c6\" colspan=\"1\" rowspan=\"1\"><p class=\"c1\"><span class=\"c0 c10\">Feedback Requested for:</span></p><p class=\"c1\"><span class=\"c0\">Name of Nominee (User ID)</span></p></td></tr>";
var TABLE_END = "</tbody></table><p class=\"c5\"><span class=\"c7\"></span></p>";
var EMAIL_END = "<p class=\"c11\"><span class=\"c17\">Please provide your feedback using </span><span class=\"c12 c13\"><a class=\"c15\" href=\""
+ FEEDBACK_FORM_LINK
+ "\">this form</a></span><span class=\"c7\">, which will take approximately "
+ TIME_NEEDED_TO_COMPLETE_SURVEY
+ " to complete."
+ "</span></p><p class=\"c5\"><span class=\"c7\"></span></p><p class=\"c11\"><span class=\"c7\">"
+ "If you have any questions, please contact your HRBP or the Program Team (vp-promotion@yahoo-inc.com).</span></p><p class=\"c5\"><span class=\"c7\"></span></p><p class=\"c8\"><span class=\"c4\">"
+ "Thank you,</span></p><p class=\"c8\"><span class=\"c4\">"
+ EMAIL_SENDER
+ "</span></p><p class=\"c5\"><span class=\"c10 c14\"></span></p></body></html>";

// This constant is written in column E for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT"

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
 
  var response = Browser.msgBox("You are about to send emails to EVERYONE ON SHEET " + sheet.getName().toUpperCase() + " (WHAAATTTT?????)."
             	+"Press OK if you are confident you wont break corpmail",Browser.Buttons.OK_CANCEL);
 
  var response2 = Browser.msgBox("are you sure you're really sure? if yes, press ok. if no, press cancel and do a dance...you saved corpmail!",Browser.Buttons.OK_CANCEL);
 
  var response3 = Browser.msgBox("WOW....you are really persistent. last chance. are you really really really sure???",Browser.Buttons.OK_CANCEL);
 
  if (response == "ok" & response2 == "ok" & response3 == "ok") {
	var startRow = 2;  // First row of data to process
	var startCol = 1;
	var numRows = sheet.getLastRow() - 1; // Remove header row
	var numCols = sheet.getLastColumn();
	// Fetch the range of cells from A2 to Last Row and Last Column of sheet
	var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
	// Fetch values for each row in the Range.
	var data = dataRange.getValues();
	for (i = 0; i < data.length; ++i) {
  	var row = data[i];
 	 
  	var feedbackProviderFirstName = row[1]; // 2nd column (col B)
  	var feedbackProviderEmailAddress = row[2]; // 3rd column (col C)
  	var nomineesWithLeaders = row[3].split(";"); // 4th column (col D)
  	var hasEmailAlreadyBeenSent = row[4]; // 5th column (col E)
 	 
  	var TABLE_ROWS = "";
  	for (j = 0; j < nomineesWithLeaders.length; j++) {
    	var nomineeAndLeader = nomineesWithLeaders[j].trim().split("---");
    	var nominee = nomineeAndLeader[0].trim();
    	var leader = nomineeAndLeader[1].trim();
   	 
    	var TABLE_ROW = "<tr class=\"c16\"><td class=\"c9\" colspan=\"1\" rowspan=\"1\"><p class=\"c1\"><span class=\"c4\">"
                    	+ leader
                    	+ "</span></p></td><td class=\"c9\" colspan=\"1\" rowspan=\"1\"><p class=\"c1\"><span class=\"c4\">"
                    	+ nominee
                    	+ "</span></p></td></tr>";
   	 
    	TABLE_ROWS = TABLE_ROWS + TABLE_ROW;
  	}
 	 
  	var EMAIL_INTRO = "<body class=\"c2\"><p class=\"c8\"><span class=\"c10 c12\">"
  	+ "Hi " + feedbackProviderFirstName + ","
  	+ "<p class=\"c5\"><span class=\"c7\"></span></p></span></p><p class=\"c11\"><span class=\"c7\">"
  	+ "We are reaching out because you have not yet completed feedback for the nominee listed below. Please submit your feedback by "
    + DATE_AND_TIME_OF_SURVEY_DEADLINE
    + ", as we will begin packet creation at that time. If you do not wish to provide feedback for Hugo, please let me know."
  	+ "</span><p class=\"c5\"><span class=\"c7\"></span></p>"
 	 
  	var message = EMAIL_START + EMAIL_INTRO + TABLE_START + TABLE_HEADERS + TABLE_ROWS + TABLE_END + EMAIL_END;
 	 
  	MailApp.sendEmail(feedbackProviderEmailAddress,SUBJECT,message, {
    	htmlBody: message,
    	bcc: "josefnunez@yahoo-inc.com"
  	});
 	 
  	// Write "EMAIL_SENT" in last column to confirm email delivery
  	sheet.getRange(startRow + i, 5).setValue(EMAIL_SENT);
  	SpreadsheetApp.flush();
	}
  }
}