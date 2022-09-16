# userDelegation
Pull info from the Sheet, and Add, Delete or List delegates

[Copy template](https://docs.google.com/spreadsheets/d/1DQM3g39_C1y1cKvd7E7NtjBtJu_qFgai8X1GYlU3o6Q/copy) to get Sheet with Google Apps Script.

Use the _userDelegation_ script menu to list, add and delete delegates.

Add inbox accounts to list delegates for in column A of sheet _Manage_.<br>
Add a user to column C to which you wish to _Add_ the inbox in column A of sheet _Manage_.<br>
Add a user to column D which you wish to _Delete_ from the inbox in column A of sheet _Manage_.

The result of each action will be added to the sheet _Log_.<br>
When using _List_ and _Add_, the delegated accounts will also be listed in the sheet _Delegated_.<br>
The _Delete_ action only displays its result in _Log_.

## How to make it work - Yes, this is absolutely necessary to do first.
secrets.gs needs secrets from a service account with domain wide access to the necessary scopes.

[Read the wiki](https://github.com/NoSubstitute/userDelegation/wiki) how you set that up.

The secrets.gs in the repo will never have any secrets, as that would give any user acess to my domain. Which is why you also shouldn't post your secrets publically, and only give access to this sheet and script to trusted admins. Even if that's a given, it never hurts to remind people. :-)

[PRIVACY POLICY](https://tools.no-substitute.com/pp)

tl;dr - No data is sent anywhere, except between you and Google.

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.me/NoSubstitute)
