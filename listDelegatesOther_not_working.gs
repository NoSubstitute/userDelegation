/**
 * Lists users with delegated access to another account.
 */
function listDelegatesOther() {
  try {
    const response =
      Gmail.Users.Settings.Delegates.list('other.user@mydomain');
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
 * Error message: API call to gmail.users.settings.delegates.list failed with error: Delegation denied for MySuperAdminHere
 */
