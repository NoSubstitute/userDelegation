/**
 * This is what I really want to do.
 * Pull info from the Sheet, and Add, Delete or Show delegates, using delegatees in columns C & D.
 * 
 * List
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/list
 * 
 * Create
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/create
 * 
 * Delete
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/delete
 * 
 */

function addDelegate() {
  // Get User/Operator Info
  var userEmail = Session.getActiveUser().getEmail()
  // Get the current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Set the Users sheet as the sheet we're working in
  var sheet = ss.getSheetByName("Delegated");
  // Log actions to the Log sheet
  var logsheet = ss.getSheetByName('Log')
  // Get all data from the second row to the last row with data, and the last column with data
  var lastrow = sheet.getLastRow();
  var lastcolumn = sheet.getLastColumn();
  var range = sheet.getRange(2, 1, lastrow - 1, lastcolumn);
  var list = range.getValues();
  for (var i = 0; i < list.length; i++) {
    // Grab username from the first column (0), then the rest from adjoing columns and set necessary variables
    var email = list[i][0].toString();
    Logger.log(email);
    var delegatee = list[i][2].toString();
    Logger.log(delegatee);
    // For each line, try to update the user with given data, and log the result.
    try {

/**
 * None of the syntaxes below works.
 * Obviously I can't read the API documentation correctly...
 * 
 * This is what the IDE says the syntax should be, but if I use it like that it's syntax error.
 * Gmail.Users.Settings.Delegates.create(resource: Gmail_v1.Gmail.V1.Schema.Delegate, userId: string): Gmail_v1.Gmail.V1.Schema.Delegate
 * 
 * I then changed it a little.
 * Gmail.Users.Settings.Delegates.create({resource: Gmail_v1.Gmail.V1.Schema.Delegate, userId: delegatee},email);
 * Gives error: ReferenceError: Gmail_v1 is not defined
 * 
 * So then I started cutting some words from the suggested syntax, to try and get something going.
 * Gmail.Users.Settings.Delegates.create({userId: delegatee}, "me");
 * Gives error "Access restricted to service accounts that have been delegated domain-wide authority"
 * 
 * Both these seem to fail for the same reason. (Found separate references to userId and delegateEmail when I googled.)
 * Gmail.Users.Settings.Delegates.create({"delegateEmail": delegatee}, email);
 * Gmail.Users.Settings.Delegates.create({userId: delegatee}, email);
 * Gives error: API call to gmail.users.settings.delegates.create failed with error: Delegation denied for admin.ns@no-substitute.net
 *  
 * Perhaps it needs to use a REST call instead, with oauth and stuff, which I don't know how to do.
 * 
 * This is what the IDE syntax text says
 * "Lists the delegates for the specified account. This method is only available to service account clients that have been delegated domain-wide authority."
 * 
 * https://developers.google.com/gmail/api/reference/rest#rest-resource:-v1.users.settings.delegates
 */
      // NONE OF THESE ACTUALLY WORK!
      // Gmail.Users.Settings.Delegates.create({userId: delegatee}, "me");
      // Gmail.Users.Settings.Delegates.create({userId: delegatee}, email);
      // Gmail.Users.Settings.Delegates.create({"delegateEmail": delegatee}, email);

      logsheet.appendRow([new Date(), userEmail, "Account: " + email + " delegated to " + delegatee]);

      // If the update fails for some reason, log the error
    } catch (err) {
      logsheet.appendRow([email, err]);
      Logger.log(err);
    }
  }
  SpreadsheetApp.flush();
}
