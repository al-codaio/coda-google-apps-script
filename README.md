# One-way data syncs between Coda and Google Sheets
These are [Google Apps Scripts](https://developers.google.com/apps-script/overview) for syncing data between a [Coda doc](www.coda.io) and a Google Sheets file. The one-way sync scripts supports adding, deleting, and updating rows of data. List of the scripts in this repo:
1. [**coda_to_coda.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_coda.js) - Sync data from tables in one Coda doc to tables in another Coda doc ([gist](https://gist.github.com/albertc44/c3d5ebc2a9ec00a28e561ea8e64fc0c5))
2. [**sheets_to_sheets.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_sheets.js) - Sync data from one Google Sheet to another Google Sheet ([gist](https://gist.github.com/albertc44/bbae27144db5f1f75b76794d6622b3f9))
3. [**coda_to_sheets.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js) - Sync data from a table in a Coda doc to a worksheet in Google Sheets ([gist](https://gist.github.com/albertc44/ec44e1aae95730b827e6b58a7ec9a317))
4. [**sheets_to_coda.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js) - Sync data from a worksheet in Google Sheets to a table in a Coda doc ([gist](https://gist.github.com/albertc44/5fd208938870390fae6a92856e2309f3))

## What you'll need
You'll need the following to execute the scripts above that involve syncing data between Coda and Google Sheets:
* **Coda API key** - You'll see this as `YOUR_API_KEY` (get this from your Coda [account settings](https://coda.io/account))
* **Coda doc ID** - You'll see this as `YOUR_TARGET_DOC_ID` or `YOUR_SOURCE_DOC_ID`
* **Coda table ID** - You'll see this as `YOUR_TARGET_TABLE_ID` or `YOUR_SOURCE _TABLE_ID`
* **Google Sheets ID** - You'll see this as `TARGET_SHEET_ID` or `SOURCE_SHEET_ID`
* **Google Sheets worksheet name** - You'll see this as `TARGET_WORKSHEET_NAME` or `SOURCE_WORKSHEET_NAME`
* **Coda's library for Google Apps script** - `15IQuWOk8MqT50FDWomh57UqWGH23gjsWVWYFms3ton6L-UHmefYHS9Vl`

## How to use
For syncing data Coda to Coda and Sheets to Sheets, read [this tutorial](https://coda.io/@atc/how-to-sync-data-between-coda-docs-and-google-sheets-using-googl). 
For sycning data between Coda and Google Sheets, read this tutorial.

## Caveats
* You will have to set up a [time-driven installable trigger](https://developers.google.com/apps-script/guides/triggers/installable) in Google App Scripts to get the scripts to run every minute, hour, etc.
* For syncing data from a Google Sheet to Coda, you must have *edit* or *view* access to the Gogole Sheet
* You cannot sort data in the Google Sheet if syncing from Sheets to Coda (read tutorial for more detail)
* Formulas you write in Coda or Google Sheets will get lost when synced to the target

