/**
 * Lists the user's labels, including name, type,
 * ID and visibility information.
 */
function listLabelsOther() {
  try {
    const response =
      Gmail.Users.Labels.list('other.user@mydomain');
    for (let i = 0; i < response.labels.length; i++) {
      const label = response.labels[i];
      Logger.log(JSON.stringify(label));
    }
  } catch (err) {
    Logger.log(err);
  }
}

/**
 * Result: I don't get a list of the other user's labels.
 * Error message says I'm denied because I lack delegated access.
 * API call to gmail.users.labels.list failed with error: Delegation denied for MySuperadminAccountHere
 */
