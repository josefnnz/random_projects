function auto_response() 
{
  var form = FormApp.openByUrl('https://docs.google.com/a/oath.com/forms/d/1p2L0E42xGXHAnX8dsdFqI6AQ7rh4_JPwoLjIioekync/edit?usp=sharing_eil&ts=5995e689');
  var responses = SpreadsheetApp.openById('13JaI8P2YlhL4Z_lJ511XvzVxAUnglPmIEVwm7c3WhKQ').getSheetByName('Old doc').getRange(2,1,71,38).getValues();

  for (var i = 0; i < responses.length; i++) 
  {
    var formResponse = form.createResponse();
    var items = form.getItems();

    // Fill-in form with response from spreadsheet row
    var row = responses[i];
    var answer_index = 4; // Answers start in the 4th column of the spreadsheet
    for (var j = 0; j < items.length; j++) 
    {
    	curr = items[j];
    	var formItem = null;
    	if (curr.getType() == 'MULTIPLE_CHOICE')
    	{
    		formItem = curr.asMultipleChoiceItem();
    	}
    	else if (curr.getType() == 'SCALE')
    	{
    		formItem = curr.asScaleItem();
    	}
    	else if (curr.getType() == 'PARAGRAPH_TEXT')
    	{
    		formItem = curr.asParagraphTextItem();
    	}
    	else if (curr.getType() == 'TEXT')
    	{
    		formItem = curr.asTextItem();
    	}
    	else
    	{
    		// Skip items that do not require a response
    		continue;
    	}
    	var r = formItem.createResponse(row[answer_index]);
    	formResponse.withItemResponse(r);
    	answer_index++;
    }

    // Submit completed response. Sleep to wait for submission to complete before starting next response.
    formResponse.submit();
    Utilities.sleep(500);

  }

};