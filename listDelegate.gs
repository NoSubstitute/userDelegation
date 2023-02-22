/**
 * @OnlyCurrentDoc
 */

/**
 * This process works and gives relevant feedback to the user when the process fails.
 * I have disabled all logging to console, but keeping it in the code, for easy debugging in the future.
 */

function listGmailDelegate() {
  // Get User/Operator Info
  var userEmail = Session.getActiveUser().getEmail();
  // Get the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Set the Users sheet as the sheet we're working in
  var sheet = ss.getSheetByName("Manage");
  // var sheet = ss.getSheetByName("ManageList"); // For dev testing I use this sheet with less users to get less error content in the log
  // Log actions to the Log sheet
  var logsheet = ss.getSheetByName('Log');
  // List delegates to the Delegates sheet
  var delegatessheet = ss.getSheetByName('Delegated');
  // Get all data from the second row to the last row with data, and the last column with data
  var lastrow = sheet.getLastRow();
  var lastcolumn = sheet.getLastColumn();
  var range = sheet.getRange(2, 1, lastrow - 1, lastcolumn);
  var list = range.getValues();
  for (let i = 0; i < list.length; i++) {
    // Grab username from the first column (0), then the rest from adjoing columns and set necessary variables
    var boxEmail = list[i][0].toString();
    // For each line, try to list the user with given data, and log the result
    try {
      // Logger.log("Trying to list delegates for " + boxEmail);
      // Check to see if userEmail has service account access to boxEmail
      // If not, go straight to else to see why, or to catch if it is an account without delegates
      var serviceList = getListDelegationService_(boxEmail);
      if (serviceList.hasAccess()) {
        // Prepare the data to be included in the UrlFetccApp call
        var options = {
          "method": "GET",
          "contentType": "application/json",
          "muteHttpExceptions": true,
          "headers": {
            "Authorization": 'Bearer ' + serviceList.getAccessToken()
          }
        };
        var url = 'https://gmail.googleapis.com/gmail/v1/users/' + boxEmail +
          '/settings/delegates';
        // Run the actual API call by fetching a URL with the necessary information included
        var response = UrlFetchApp.fetch(url, options);
        // This logger would show the real response from the API call
        // Logger.log(response);
        // Logger.log("API Response: " + JSON.stringify(JSON.parse(response)));
        var delegateList = JSON.parse(response)
        // For each of the delegates list the boxEmail and delegate
        for (let j = 0; j < delegateList.delegates.length; j++) {
          const delegate = delegateList.delegates[j].delegateEmail;
          const status = delegateList.delegates[j].verificationStatus;
          // Print the delegates in the sheet Delegated, one row per delegate
          logsheet.appendRow([new Date(), userEmail, boxEmail + " is delegated to " + delegate]);
          delegatessheet.appendRow([new Date(), boxEmail, delegate]);
        }
      } else {
        // Show the error message from the API
        // Logger.log("FAIL else - " + serviceList.getLastError());
        // Convert error message to a string
        var checkElseError = serviceList.getLastError();
        var checkElse = checkElseError.toString();
        var errorInvalidRequest = "invalid_request";
        var errorInvalidGrant = "invalid_grant";
        // Check error message against the two known errors errorInvalidRequest & errorInvalidGrant and react accordingly
        if (checkElse.includes(errorInvalidRequest)) {
          // Logger.log("Error contains invalid request");
          logsheet.appendRow([new Date(), userEmail, "FAILED LIST Request - Failed to list delegates of " + boxEmail + " - check spelling"]);
        } else if (checkElse.includes(errorInvalidGrant)) {
          // Logger.log("Error contains invalid grant");
          logsheet.appendRow([new Date(), userEmail, "FAILED LIST Grant - Failed to list delegates of " + boxEmail + " - check spelling"]);
        }
        else {
          // Logger.log("Neither error message matches. No clue what's wrong. Potentially add more else if loops here.");
          logsheet.appendRow([new Date(), userEmail, "FAILED LIST - Failed to list delegates of " + boxEmail + " - reason unknown"]);
        }
      }
      // If the list fails for some reason, log the error
    } catch (err) {
      if (!(err instanceof SyntaxError)) {
        throw err; // rethrow (I don't know how to deal with unknown error types, but there also shouldn't be any.)
      }
      else if (err instanceof SyntaxError) {
        // Show the error message from the API
        // Logger.log("FAIL err - " + err);
        // Logger.log(err.name);
        // Logger.log(err.message);
        // Since I have verified that SyntaxError always means account is not delegated, print that in the log
        logsheet.appendRow([new Date(), userEmail, "NOT DELEGATED - " + boxEmail]);
      }
    } finally {
      // Logger.log("This runs regardless, and is only good for debugging, to see if the code gets this far");
      // Now the script cleans up the properties after each run, so as to avoid Properties Storage Quota errors
      PropertiesService.getScriptProperties().deleteProperty("oauth2.Gmail:"+boxEmail)
    }
  }
  SpreadsheetApp.flush();
}

/**
 * Do I really need three separate services, just because I have three different actions?
 * Create, Delete & List
 * I don't know, so keeping them all for now.
 */

function getListDelegationService_(boxEmail) {
  return OAuth2.createService('Gmail:' + boxEmail)
    // Set the endpoint URL.
    .setTokenUrl('https://oauth2.googleapis.com/token')
    // Set the private key and issuer. Values taken from secrets.gs.
    .setPrivateKey(PRIVATE_KEY)
    .setIssuer(CLIENT_EMAIL)
    // Set the name of the user to impersonate.
    .setSubject(boxEmail)
    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getScriptProperties())
    // Set the scope. This must match the scopes configured during the
    // setup of domain-wide delegation.
    .setScope(['https://www.googleapis.com/auth/gmail.settings.basic', 'https://www.googleapis.com/auth/gmail.settings.sharing']);
}
