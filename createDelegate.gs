/**
 * @OnlyCurrentDoc
 */

function createGmailDelegate() {
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
    var delegatee = list[i][2].toString();
    // For each line, try to update the user with given data, and log the result.
    try {
      Logger.log("Trying to let " + delegatee + " read " + boxEmail);
      // Check to see if userEmail has service account access to boxEmail
      var service = getCreateDelegationService_(boxEmail);
      if (service.hasAccess()) {
        // Prepare the data to be included in the UrlFetccApp call
        var payload = {
          "delegateEmail": delegatee,
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
        Logger.log("API Response: " + JSON.stringify(JSON.parse(response)));
        Logger.log("Tried to let " + delegatee + " read " + boxEmail);
        logsheet.appendRow([new Date(), userEmail, "ATTEMPTED DELEGATION - Tried to delegate " + boxEmail + " to " + delegatee]);
        var checkResponse = JSON.stringify(JSON.parse(response));
        var acceptedStatus = "accepted";
        var alreadyExists = "ALREADY_EXISTS";
        if (checkResponse.includes(acceptedStatus)) {
          Logger.log("Delegated " + boxEmail + " to " + delegatee);
          // logsheet.appendRow([new Date(), userEmail, "SUCCESSFUL DELEGATION - Delegated " + boxEmail + " to " + delegatee]);
        } else if (checkResponse.includes(alreadyExists)) {
          Logger.log(delegatee + " is already a delegate of " + boxEmail);
          // logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - " + delegatee + " is already a delegate of " + boxEmail]);
        } else {
          Logger.log("Failed to let " + delegatee + " read " + boxEmail);
          // logsheet.appendRow([new Date(), userEmail, "FAILED DELEGATION - Failed to delegate " + boxEmail + " to " + delegatee]);
        }
      } else {
        Logger.log("FAIL else - " + boxEmail + "," + delegatee + ", " + service.getLastError());
        // logsheet.appendRow([new Date(), userEmail, "FAIL - " + boxEmail + " or " + delegatee + " is invalid - check spelling"]);
      }
      // If the create fails for some reason, log the error
    } catch (err) {
      Logger.log("FAIL err - " + boxEmail + "," + delegatee + ", " + err);
      // logsheet.appendRow([new Date(), userEmail, "FAIL - " + err]);
    }
  }
  SpreadsheetApp.flush();
}

/**
 * Do I really need three separate services, just because I have three different actions?
 * Create, Delete & List
 */

function getCreateDelegationService_(boxEmail) {
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