[![Coda API 1.0.0](https://img.shields.io/badge/Coda%20API-1.0.0-orange)](https://coda.io/developers/apis/v1)

# One-way data syncs between Coda and Google Sheets
These are [Google Apps Scripts](https://developers.google.com/apps-script/overview) for syncing data between a [Coda doc](www.coda.io) and a Google Sheets file. The scripts supports adding, deleting, and updating rows of data in your target Coda doc or Google Sheet. List of the scripts in this repo:
1. [**coda_to_coda.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_coda.js) - Sync data from tables in one Coda doc to tables in another Coda doc ([gist](https://gist.github.com/albertc44/c3d5ebc2a9ec00a28e561ea8e64fc0c5))
2. [**sheets_to_sheets.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_sheets.js) - Sync data from one Google Sheet to another Google Sheet ([gist](https://gist.github.com/albertc44/bbae27144db5f1f75b76794d6622b3f9))
3. [**coda_to_sheets.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js) - Sync data from a table in a Coda doc to a worksheet in Google Sheets ([gist](https://gist.github.com/albertc44/ec44e1aae95730b827e6b58a7ec9a317))
4. [**sheets_to_coda.js**](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js) - Sync data from a worksheet in Google Sheets to a table in a Coda doc ([gist](https://gist.github.com/albertc44/5fd208938870390fae6a92856e2309f3))

## Setup for Coda to Google Sheets script
Starting in [line 9](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js#L9) to [line 14](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js#L14) of the [coda_to_sheet.js](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js#L10) script, you'll need to enter in some of your own data to make the script work. Step-by-step:
1. Go to [script.google.com](script.google.com) and create a new project and give your project a name.
2. Go to **Libraries** then **Resources** and paste the following string of text/numbers into the library field: `15IQuWOk8MqT50FDWomh57UqWGH23gjsWVWYFms3ton6L-UHmefYHS9Vl`.
3. Click **Add** and then select version 9 of library to use (as of August 2020, version 9 - Coda API v1.0.0 is the latest)
4. Copy and paste the [entire script](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_sheets.js) into your Google Apps Script project and click **File** then **Save**.
5. Go to your Coda [account settings](https://coda.io/account), scroll down until you see "API SETTINGS" and click **Generate API Token**. Copy and paste that API token into the value for `YOUR_API_KEY` in the script. *Note: do not delete the single apostrophes around* `YOUR_API_KEY`.
6. Get the the doc ID from your Coda doc by copying and pasting all the characters after the `_d` in the URL of your Coda doc (should be about 10 characters). You can also use the *Doc ID Extractor* tool in the [Coda API docs](https://coda.io/developers/apis/v1beta1#section/Using-the-API/Resource-IDs-and-Links). Copy and paste your doc ID into `YOUR_SOURCE_DOC_ID`.
7. Go back to your [account settings](https://coda.io/account) and scroll down to the very bottom until you see "Labs." Toggle "Enable Developer Mode" to **ON**.
8. Hover over the table name in your Coda doc and click on the 3 dots that show up next to your table name. Click on "Copy table ID" and paste this value into `YOUR_SOURCE_TABLE_ID`.
9. To get your Google Sheets ID, get all the characters after `/d/` in your Google Sheets file up until the slash and paste this into `YOUR_GOOGLE_SHEETS_ID`. See [this link](https://stackoverflow.com/a/36062068/1110697) for more info.
10. Write in the name of the worksheet from your Google Sheets file where data will be sycned into in the `YOUR_GOOGLE_SHEETS_WORKSHEET_NAME` value. 
11. In Google Sheets, create a new column name at the end of your column headers called something like `Coda Source Row URL` and make sure there is no data in that column below the header. Write that column name in `YOUR_SOURCE_ROW_URL_COLUMN_NAME`. 
12. Go back to Google Apps Script, click on the **Select function** dropdown in the toolbar, and select `runSync`. Then click the play ▶️ button to the left of the bug 🐞 button. This should copy over all the data from your Coda doc to Google Sheets.
13. To get the script to run every minute, hour, or day, click on the clock 🕒 button to the left of the ▶️ button to create a [time-driven trigger](https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers).
14. Click **Add Trigger**, make sure `runSync` is set as the function to run, "Select event source" should be `Time-driven`, and play around with the type of time based trigger that fits your needs. I like to set the "Failure notification settings" to `Notify me immediately` so I know when my script fails to run.

## Setup for Google Sheets to Coda script
Most of the steps above apply to the [sheets_to_coda.js](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js) script as well but there are few extra features.
1. You can follow steps 1-10 above to fill out [line 12](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js#L12) to [line 18](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js#L18) in the script (except [line 14](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js#L14) mentioned in the next step). The main difference is that "SOURCE" and "TARGET" are flipped around since you are now syncing from a *source* Google Sheet to a *target* Coda doc.
2. Your Coda table *cannot* have a column named `Coda Row ID`. If you need to use a column with this name, replace the `TARGET_ROW_ID_COLUMN` variable with another value.
3. If you have *edit access* to the Google Sheet, follow step 11 above and write in the column name in `YOUR_SOURCE_ROW_URL_COLUMN_NAME`.
4. If you want the ability to add rows to your Coda table and NOT have these rows deleted every time the sync runs, create a column in your Coda table and name it `Do not delete`. This column should be a checkbox column format and you will check the box for every row you manually add to your Coda table that you want to keep in that table. Otherwise, the script will delete that row and always keep the Coda table a direct copy of what's in your Google Sheets file. If you change the name of this `Do not delete` column, you must edit the value of the `DO_NOT_DELETE_COLUMN` variable in [line 22](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js#L22) of the script as well.
5. If you want the script to completely delete and re-write the rows in your Coda table each time the script runs, set the `REWRITE_CODA_TABLE` to `true` in [line 23](https://github.com/albertc44/coda-google-apps-script/blob/master/sheets_to_coda.js#L23). This may make the script run faster, but may not be faster for larger tables (few thousand rows). For Google Sheets files where you only have *view-only access*, this setting will automatically get set to `true`.
6. Follow steps 12-14 in the Coda to Google Sheets section above to set up your time-driven trigger.

## Setup for Coda to Coda script
You can follow steps 1-8 in the Coda to Google Sheets section above to get the values you need for the [coda_to_coda.js](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_coda.js) script. The only thing you need to add is a column in your *target* Coda table (the table where data is getting synced into) called `Source Row URL`. If you change the name of this column, you must change the value of the `TARGET_TABLE_SOURCE_ROW_COLUMN` variable in [line 5](https://github.com/albertc44/coda-google-apps-script/blob/master/coda_to_coda.js#L5) of the script.

## Other notes
* When possible, keep the column names between your Coda doc and Google Sheet the same. There are some exceptions to this which are mentioned in the [blog post](https://coda.io/@atc/how-to-sync-data-from-coda-to-google-sheets-and-vice-versa-with-google-apps-script-tutorial).
* You will have to set up a [time-driven installable trigger](https://developers.google.com/apps-script/guides/triggers/installable) in Google App Scripts to get the scripts to run every minute, hour, etc.
* For syncing data from a Google Sheet to Coda, you must have *edit* or *view* access to the Google Sheet
* You cannot sort data in the Google Sheet if syncing from Sheets to Coda (read [blog post](https://coda.io/@atc/how-to-sync-data-from-coda-to-google-sheets-and-vice-versa-with-google-apps-script-tutorial) for more detail)
* Formulas you write in Coda or Google Sheets will get lost when synced to the target

## Tutorials
Here are a few blog posts explaining how the scripts work. For syncing data Coda to Coda and Sheets to Sheets, read [this tutorial](https://coda.io/@atc/how-to-sync-data-between-coda-docs-and-google-sheets-using-googl). 
For sycning data between Coda and Google Sheets, read [this tutorial](https://coda.io/@atc/how-to-sync-data-from-coda-to-google-sheets-and-vice-versa-with-google-apps-script-tutorial). Here are a few YouTube tutorials on how to setup and use the scripts:

### Sync data between two Coda docs (and two Google Sheets)
[![sync data between coda docs](https://p-ZmF7dQ.b0.n0.cdn.getcloudapp.com/items/nOu8vQ8x/Tutorial_%20One-way%20data%20sync.png?v=9c40ea6fa52f90bd1070744c668abc65)](https://www.youtube.com/watch?v=PL_uSeKmYkc)

### Sync data from a Coda doc to a Google Sheet
[![sync data between coda docs](https://p-ZmF7dQ.b0.n0.cdn.getcloudapp.com/items/bLueqy0r/Tutorial_%20Coda%20to%20Sheets%20Sync.png?v=a7264086cdb7497c6574e4a7140896ba)](https://www.youtube.com/watch?v=mAdAe8GVCdA)

### Sync data from a Google Sheet to a Coda doc
[![sync data between coda docs](https://p-ZmF7dQ.b0.n0.cdn.getcloudapp.com/items/8Luj56nw/Tutorial_%20Sheets%20to%20Coda%20Sync.png?v=08559df04bfe82b143795afbc98f58e3)](https://www.youtube.com/watch?v=xVWu9jdBm_U)
