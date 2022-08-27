/**
 * @OnlyCurrentDoc
 */

/**
 * This is the menu for the userDelegation script.
 * 
 * Pull Accounts (user's emails) from the Sheet column A, and List their delegates, or Add/Delete delegates, using delegates in columns C & D.
 * 
 * List
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/list
 * 
 * Create
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/create
 * 
 * Delete
 * https://developers.google.com/gmail/api/reference/rest/v1/users.settings.delegates/delete
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('userDelegation')
    .addItem('List delegates', 'listGmailDelegate')
    .addSeparator()
    .addItem('Add delegates', 'createGmailDelegate')
    .addSeparator()
    .addItem('Delete delegates', 'deleteGmailDelegate')
    .addToUi();
}

/**
 * I don't know how to use this properly, yet, so not calling it anywhere, yet.
 * It's supposed to make it posible to clear the Properties stored by the three services.
 */
// function logout() {
//   var service = getListDelegatesService_()
//   service.reset();
// }
