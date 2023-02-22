/**
 * @OnlyCurrentDoc
 */

/**
 * This process works and gives relevant feedback to the user when the process fails.
 * I have disabled all logging to console, but keeping it in the code, for easy debugging in the future.
 */

function createGmailDelegate() {
  // Get User/Operator Info
  var userEmail = Session.getActiveUser().getEmail();
  // Get the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Set the Users sheet as the sheet we're working in
  var sheet = ss.getSheetByName("Manage");
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
    var delegate = list[i][1].toString();
    // For each line, try to update the user with given data, and log the result.
    try {
      // Logger.log("Trying to let " + delegate + " read " + boxEmail);
      // Check to see if userEmail has service account access to boxEmail
      var service = getCreateDelegationService_(boxEmail);
      if (service.hasAccess()) {
        // Prepare the data to be included in the UrlFetccApp call
        var payload = {
          "delegateEmail": delegate,
          "verificationStatus": "accepted"
        }
        var options = {
          "method": "POST",
          "contentType": "application/json",
          "muteHttpExceptions": true,
          "headers": {
            "Authorization": 'Bearer ' + service.getAccessToken()
          },
          "payload": JSON.stringify(payload)
        };
        var url = 'https://www.googleapis.com/gmail/v1/users/' + boxEmail +
          '/settings/delegates';
        // Run the actual API call by fetching a URL with the necessary information included
        var response = UrlFetchApp.fetch(url, options);
        // var json = response.getContentText();
        // This logger would show the real response from the API call
        // Logger.log(json);
        // Logger.log("Tried to let " + delegate + " read " + boxEmail);
        // logsheet.appendRow([new Date(), userEmail, "ATTEMPTED DELEGATION - Tried to delegate " + boxEmail + " to " + delegate]);
        var checkResponse = JSON.stringify(JSON.parse(response));
        // Logger.log(checkResponse);
        var acceptedStatus = "accepted";
        var alreadyExists = "ALREADY_EXISTS";
        var missingDelegateEmail = "Missing delegate email"
        var delegateSameAsboxEmail = "Delegate and delegator are the same user"
        var invalidDelegate = "Invalid delegate";
        if (checkResponse.includes(acceptedStatus)) {
          // Logger.log("Delegated " + boxEmail + " to " + delegate);
          // List delegation to the Log sheet
          logsheet.appendRow([new Date(), userEmail, "SUCCESSFUL DELEGATION - Delegated " + boxEmail + " to " + delegate]);
          // List delegates to the Delegates sheet
          delegatessheet.appendRow([new Date(), boxEmail, delegate]);
        } else if (checkResponse.includes(alreadyExists)) {
          // Logger.log(delegate + " is already a delegate of " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - " + delegate + " is already a delegate of " + boxEmail]);
        } else if (checkResponse.includes(missingDelegateEmail)) {
          // Logger.log("Delegate cell for " + boxEmail + " seems to be empty);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - It seems you forgot to supply a new delegate for " + boxEmail]);
        } else if (checkResponse.includes(delegateSameAsboxEmail)) {
          // Logger.log(boxEmail + " and " + delegate + " are the same");
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - Can't delegate to same account. Check spelling of " + boxEmail + " and " + delegate]);
        } else if (checkResponse.includes(invalidDelegate)) {
          // Logger.log("It seems " + delegate + " is invalid. Check the spelling");
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - It seems " + delegate + " is invalid. Check the spelling"]);
        } else {
          // Logger.log("Failed to let " + delegate + " read " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - Failed to delegate " + boxEmail + " to " + delegate]);
        }
      } else {
        // If delegation wasn't possible for some other reason try to figure out why, and log that.
        // Logger.log("FAIL else - " + boxEmail + "," + delegate + ", " + service.getLastError());
        var checkElseError = service.getLastError();
        // Logger.log(checkElseError);
        var checkElse = checkElseError.toString();
        // Logger.log(checkElse);
        // Both these errors means that the boxEmail is invalid, but I separate them into two processes anyway.
        var errorInvalidRequest = "invalid_request";
        var errorInvalidGrant = "invalid_grant";
        if (checkElse.includes(errorInvalidRequest)) {
          // Logger.log("Tried to add " + delegatee + " to " + boxEmail);
          // Logger.log("Invalid Request - Failed to add delegate - check spelling of " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - Invalid Request - " + boxEmail + " is invalid - check spelling"]);
        } else if (checkElse.includes(errorInvalidGrant)) {
          // Logger.log("Tried to add " + delegatee + " to " + boxEmail);
          // Logger.log("Invalid Grant - Failed to add delegate - check spelling of " + boxEmail);
          logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - Invalid Grant - " + boxEmail + " is invalid - check spelling"]);
        }
      }
      // If the create fails for some unknown reason, log the error.
    } catch (err) {
      // Logger.log("FAIL err - " + boxEmail + "," + delegate + ", " + err);
      logsheet.appendRow([new Date(), userEmail, "FAIL - " + err]);
    } finally {
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

function getCreateDelegationService_(boxEmail) {
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
    // Set the scope. This must match one of the scopes configured during the
    // setup of domain-wide delegation.
    .setScope(['https://www.googleapis.com/auth/gmail.settings.basic', 'https://www.googleapis.com/auth/gmail.settings.sharing']);
}
