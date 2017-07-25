// This function assumes you have a Google Form with one text input for the 
// email address of the contact being unsubscribed.
// 

function unsubEmail() {
  // Gets the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  // Finds the last filled row on the spreadsheet (a.k.a. most-recent entry)
  var lastRow = sheet.getLastRow();
  // Finds the last filled column on the spreadsheet (a.k.a. email)
  var lastCol = sheet.getLastColumn()
  // Locates the value of the email cell. 
  // The +1 in the lastCol allows us to access that column later in the function.
  var email = sheet.getRange(1, 1, lastRow, lastCol + 1).getCell(lastRow,2).getValue();  
  
  // Add your HubSpot info here!
  var portalId = 'YOUR_PORTAL_ID';
  var key = 'YOUR_API_KEY';
  
  // The base URL for the HubSpot subscriptions endpoint
  var base = 'https://api.hubapi.com/email/public/v1/subscriptions/';
  
  // Builds the endpoint URL
  var url = base + email + '?portalId=' + portalId + '&hapikey=' + key;  
  
  // Unsubscribes contact from all email lists
  var bundle = {
   'unsubscribeFromAll': 'true'
  };
  
  // Various options for the request
  var options = {
    "method" : "put",
    'contentType': 'application/json',
    "payload" : JSON.stringify(bundle),
    "muteHttpExceptions" : true
  };
  
  // This is where the request is actually sent. The response is logged in the spreadsheet. 
  // If you want to see the response in the Log, make 'muteHttpExceptions' false in the above options.
  try {
    var response = UrlFetchApp.fetch(url, options); 
    
    if (response.getResponseCode() === 200) {
      Logger.log(response);
      sheet.getRange(lastRow, 3).setValue('Success! Response Code: ' + response.getResponseCode());
    }
  } catch (err) {
    Logger.log(err)
    sheet.getRange(lastRow, 3).setValue('oh no! Something went wrong.');
  }
}
