# Get-Dividends

A JavaScript project that can be added to your Google Sheets dividend portfolio spreadsheet that will automatically update the dividend amount, dividend frequency, ex dividend date, and dividend pay date.

For a full walkthrough on how to set the script up, check out my [YouTube video](https://youtu.be/p1ZSY8YkGeM) where I guide you through the process.  And check out [this video]() for instructions on the updated features.

Feel free to copy this spreadsheet to get you started as well: https://docs.google.com/spreadsheets/d/1z9cZji2wxNzrupr_zd-cOe7uSlaEdz5vzUu-SbO-5N0/edit#gid=0

## Steps:

1. Your dividend portfolio spreadsheet should be organized with the companies in individual rows with data about each company in columns.
2. In the top row of menus in your spreadsheet, click on Extensions → Apps Script.
3. In the Apps Script Editor click the “+” symbol to add a file.  Name it whatever you like and delete the contents of the file.
4. Open the Dividends.gs file in this github repository.  Copy all the contents and paste them into the new file you just created in the Apps Script Editor.
5. Edit the variables at the top of the file that have names with all capital letters. 
    1. SPREADSHEET_ID is the id of your spreadsheet which is a string of alphanumeric characters between the “d” and “edit” in the URL of your spreadsheet (make sure you take this from the URL of the spreadsheet, not the Apps Script Project).
    2. SUMMARY_SHEET_NAME is the name of the sheet that you want the script to enter the dividend info in.
    3. GET_TICKER_COL is the column that the ticker symbol of the stock is in.
    4. POST_DIV_COL is the column that you want the script to enter the annual dividend of the stock in.
    5. POST_FREQ_COL is the column that you want the script to enter the number of times per year that the company pays a dividend.
    6. POST_EX_DATE_COL is the column that you want the script to enter the ex dividend date in.
    7. POST_PAY_DATE_COL is the column that you want the script to enter the pay date in.
    8. POST_PAYOUT_COL is the column that you want the script to enter the payout ratio in.
    9. POST_YEARS_INCREASING_COL is the column that you want the script to enter the years of increasing dividends in.
    10. POST_CAGR_COL is the columnt that you want the script to enter the 5 year dividend cagr in.
    11. FIRST_TICKER_ROW is the first row that a ticker symbol appears in in your spreadsheet.
    12. Google Apps Scripts Functions have a maximum execution time of 6 minutes.  So if your portfolio has a lot of companies (more than 25-30 would be my guess), the script will end before all the tickers have been updated.  So I created the VarSheet to track what row we are on if the script ends early so that all of the tickers get updated.  So VAR_SHEET_NAME is the name of the sheet that you are going to track some variables in (for now its just the ticker row; see the example spreadsheet for what this sheet should look like).  If you just copy the VarSheet from my example you shouldn't have to change this variable or the next two.
    13. GET_TICKER_ROW is the row of the ticker row on VarSheet.
    14. GET_TICKER_COL is the column of the ticker row on VarSheet.
6. Note: If you don’t want one of the values to be entered anywhere in your spreadsheet, just set the corresponding POST column variable equal to 0.
7. Create one more file in the Apps Script Editor, delete the contents, and then past the contents of Keys.gs into it.  You will edit the POLYGON_KEY variable later.
8. Go to the [Polygon website](https://polygon.io/) (Polygon.io).  Then click “Get your free API key”.
9. If you don’t have an account, create one.
10. Once your account is created, go to the left side and click “API Keys”.  Then on the right side click “Copy” to copy the value of the key to your clipboard.
11. Paste the key into the POLYGON_KEY variable in the Keys.gs file.
12. Now let’s set up the trigger.  Go to the left hand side off the Apps Script Editor and click on the clock icon or the word “Triggers”.
13. In the bottom right, click “Add Trigger” then change the following items:
    1. Change “Choose which function to run” to “AddDividends”
    2. Change “Select event source” to “Time-driven”
    3. Change “Select type of time based trigger” to “Day timer”
    4. Change “Select time of day” to “10 pm to 11pm” (you can really choose whatever you want for this one but that’s what I prefer.
14. Then click “Save”


### If you are stuck on any of the steps or have any questions/issues feel free to check out my [YouTube video](https://youtu.be/p1ZSY8YkGeM) where I walk through each step or you can contact me directly on [Instagram](https://www.instagram.com/frankmularcik/), [Twitter](https://twitter.com/FrankMularcik) or through email (frank.mularcik.investing@gmail.com).

### If you want to support me make sure to follow me on the above social medias and subscribe to my [YouTube](https://www.youtube.com/c/FrankMularcik) channel.
