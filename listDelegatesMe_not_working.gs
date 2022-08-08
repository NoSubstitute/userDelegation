/**
 * Lists users with delegated access to my account.
 */
function listDelegatesMe() {
  try {
    const response =
      Gmail.Users.Settings.Delegates.list('me');
    for (let i = 0; i < response.labels.length; i++) {
      const label = response.labels[i];
      Logger.log(JSON.stringify(label));
    }
  } catch (err) {
    Logger.log(err);
  }
}

/**
 * Result: I don't get a list of delegates.
 * Error message: API call to gmail.users.settings.delegates.list failed with error: Access restricted to service accounts that have been delegated domain-wide authority
 */
