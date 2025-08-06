# Monzo Transactions Apps Script

This is a Google Apps Script that can be embedded within a Google Sheet to automatically populate it with your Monzo transactions.

## Usage

* Make a new Google Sheet
* Go to `Extensions > Apps Script`
* Replace the code in `Code.gs` with the code from [Transactions.gs](Transactions.gs)
* Create a Monzo API client in their [developer tools](https://developers.monzo.com/)
  * Go to `Clients > New OAuth Client`
  * Give the client a name and description
  * Leave Redirects URLs blank for now
  * Set to confidential
* Get the `clientId` and `clientSecret` and enter them into the top of your `Code.gs`
* Click `Deploy > New Deployment`, select type as 'Web app', leave rest as default and click Deploy
* Copy the Web app URL, go back to the Monzo developer tools and update your client's redirect URL to it
* Visit the Web app URL. A new tab should open, follow the 'Click Here' link
* Follow the steps to authorize access to your Monzo account, once successful you will see the text 'Auth Complete'
* Make sure to also authorize access within the Monzo app
* Return to the `Code.gs` editor and run the `updateTransactions` function
* You should be prompted to configure an `accountId`, enter the one you want into the top of your `Code.gs`
* Re-run `updateTransactions`, if successful you should see the latest transactions back in your Sheet
* You can go to the `Triggers` tab and add a trigger to run the `updateTransactions` function, for example create a Time-driven trigger once a day

## License

Distributed under the GNU General Public License v3.0. See `LICENSE` for more information.

## Author

Written by [Kieren Beckett](http://kierenb.net).
