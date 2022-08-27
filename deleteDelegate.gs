/**
 * @OnlyCurrentDoc
 */

function deleteGmailDelegate() {
  // Get User/Operator Info
  var userEmail = Session.getActiveUser().getEmail();
  // Get the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Set the Users sheet as the sheet we're working in
  var sheet = ss.getSheetByName("Manage");
  // Log actions to the Log sheet
  var logsheet = ss.getSheetByName('Log')
  // Get all data from the second row to the last row with data, and the last column with data
  var lastrow = sheet.getLastRow();
  var lastcolumn = sheet.getLastColumn();
  var range = sheet.getRange(2, 1, lastrow - 1, lastcolumn);
  var list = range.getValues();
  for (let i = 0; i < list.length; i++) {
    // Grab username from the first column (0), then the rest from adjoing columns and set necessary variables
    var boxEmail = list[i][0].toString();
    var delegatee = list[i][3].toString();
    // For each line, try to update the user with given data, and log the result.
    try {
      Logger.log("Trying to remove " + delegatee + " from " + boxEmail);
      // Check to see if userEmail has service account access to boxEmail
      var serviceDelete = getDeleteDelegationService_(boxEmail);
      if (serviceDelete.hasAccess()) {
        // Prepare the data to be included in the UrlFetccApp call
        var options = {
          "method": "DELETE",
          "contentType": "application/json",
          "muteHttpExceptions": true,
          "headers": {
            "Authorization": 'Bearer ' + serviceDelete.getAccessToken()
          }
        };
        var url = 'https://www.googleapis.com/gmail/v1/users/' + boxEmail +
          '/settings/delegates/' + delegatee;
        // Run the actual API call by fetching a URL with the necessary information included
        var response = UrlFetchApp.fetch(url, options);
        // These two lines would show the real response from the API call
        var json = response.getContentText();
        Logger.log("json getContentText: " + json);
        // Since I couldn't make heads or tails of these json responses, I found that the responsecode was more straightforward
        // Haven't run into an error code that isn't 204 or 404 yet, and if so there is a log for that
        // Logging each code inside their own loop to avoid unnecessary logging
        var responsecode = response.getResponseCode();
        var responseStatus = "Invalid delegate";
        // if (checkResponseString.includes(responseStatus)) {
        if (responsecode == "204") {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          Logger.log("responsecode: " + responsecode);
          Logger.log("Deleted delegate" + delegatee + " from " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "SUCCESSFUL DELETION - Deleted " + delegatee + " from " + boxEmail]);
        } else if (responsecode == "404") {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          Logger.log("responsecode: " + responsecode);
          Logger.log("Invalid delegate " + delegatee);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Failed to delete " + delegatee + " from " + boxEmail + " - check spelling of delegate"]);
        } else {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          Logger.log("responsecode: " + responsecode);
          Logger.log("Something else wrong " + delegatee);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Failed to delete " + delegatee + " from " + boxEmail + " - check spelling of delegate"]);
        }
      } else {
        Logger.log("FAIL else - " + serviceDelete.getLastError());
        // var checkElseError = JSON.stringify(JSON.parse(serviceDelete.getLastError()));
        var checkElseError = serviceDelete.getLastError();
        // Logger.log(checkElseError);
        var checkElse = checkElseError.toString();
        // Logger.log(checkElse);
        var errorInvalidRequest = "invalid_request";
        var errorInvalidGrant = "invalid_grant";
        if (checkElse.includes(errorInvalidRequest)) {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          Logger.log("Failed to delete delegate - check spelling of " + delegatee);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Failed to delete " + delegatee + " from " + boxEmail + " - check spelling of delegate"]);
        } else if (checkElse.includes(errorInvalidGrant)) {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          Logger.log("Failed to delete delegate - check spelling of " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Account " + boxEmail + " is invalid - check spelling"]);
        }
        else {
          Logger.log("Doesn't match either error message")
        }
      }
      // If the delete fails for some reason, log the error
    } catch (err) {
      if (!(err instanceof SyntaxError)) {
        throw err; // rethrow (I don't know how to deal with this error, even if it rarely happens.)
      }
      // else if (err instanceof SyntaxError) {
      else {
        // Show the error message from the API
        Logger.log("FAIL err - " + err);
        // var checkErrMessage = err.message;
        // var checkErrMessage = JSON.stringify(err);
        // var checkErr = checkErrMessage.toString();
        // var checkErr = checkErrMessage;
        var checkErrMessage = err.message;
        var checkErr = checkErrMessage.toString();
        var errorInvalidArgument = "Invalid argument";
        var errorUnexpectedTokenE = "Unexpected token: E";
        var errorUnexpectedToken = "Unexpected token: <";
        if (checkErr.includes(errorInvalidArgument)) {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          // Logger.log(err);
          // Logger.log(checkErrMessage);
          Logger.log(checkErr);
          Logger.log("Failed to find delegate account - check spelling of " + delegatee);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Delegate " + delegatee + " is invalid - check spelling"]);
        } else if (checkErr.includes(errorUnexpectedTokenE)) {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          // Logger.log(err);
          // Logger.log(checkErrMessage);
          Logger.log(checkErr);
          Logger.log("Failed - other reason?");
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Failed to delete " + delegatee + " from " + boxEmail + " - some other reason?"]);
        } else if (checkErr.includes(errorUnexpectedToken)) {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          // Logger.log(err);
          // Logger.log(checkErrMessage);
          Logger.log(checkErr);
          Logger.log("Failed - no delegate?");
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Failed to delete " + delegatee + " from " + boxEmail + " - Did you supply a delegate to remove?"]);
        } else {
          Logger.log("Tried to remove " + delegatee + " from " + boxEmail);
          // Logger.log(err);
          // Logger.log(checkErrMessage);
          Logger.log(checkErr);
          // Logger.log("Failed to find inbox account - check spelling of " + boxEmail);
          Logger.log("My IF INCLUDES are not working, don't know why");
          logsheet.appendRow([new Date(), userEmail, "FAILED DELETION - Account " + boxEmail + " is invalid - check spelling"]);
        }
      }
    } finally {
      // Logger.log("This runs regardless, and is only good for debugging, to see if the code gets this far")
    }
  }
  SpreadsheetApp.flush();
}

/**
 * Do I really need three separate services, just because I have three different actions?
 * Create, Delete & List
 * I don't know, so keeping them all for now.
 */

function getDeleteDelegationService_(boxEmail) {
  return OAuth2.createService('Gmail:' + boxEmail)
    // Set the endpoint URL.
    .setTokenUrl('https://oauth2.googleapis.com/token')
    // Set the private key and issuer.
    .setPrivateKey(PRIVATE_KEY)
    .setIssuer(CLIENT_EMAIL)
    // Set the name of the user to impersonate.
    .setSubject(boxEmail)
    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getScriptProperties())
    // Set the scope. This must match one of the scopes configured during the
    // setup of domain-wide delegation.
    .setScope(['https://www.googleapis.com/auth/gmail.settings.basic', 'https://www.googleapis.com/auth/gmail.settings.sharing']);
}