// edit these variables
var SPREADSHEET_ID = "your id here";
var DIV_COL = 0;
var FREQ_COL = 0;
var EX_DATE_COL = 0;
var PAY_DATE_COL = 0;
var SUMMARY_SHEET_NAME = "Summary";
//

function AddDividends()
{
  var summarySheet = GetSheet(SUMMARY_SHEET_NAME);
  var row = 2;
  var ticker = summarySheet.getRange(row, 1).getValue();
  do
  {
    var response = GetDividend(ticker);
    if (response != 0)
    {
      var date = new Date(response.ex_dividend_date);
      var div = "";
      var ex_date = "";
      var pay_date = "";
      var div_freq = 0;
      // ensures latest dividend data is current
      if (new Date() - date < YearsToMs(1))
      {
        div_freq = response.frequency;
        div = response.cash_amount * div_freq;
        ex_date = response.ex_dividend_date;
        pay_date = response.pay_date;
      }
      if (DIV_COL > 0)
      {
        summarySheet.getRange(row, DIV_COL).setValue(div);
      }
      if (EX_DATE_COL > 0)
      {
        summarySheet.getRange(row, EX_DATE_COL).setValue(ex_date);
      }
      if (PAY_DATE_COL > 0)
      {
        summarySheet.getRange(row, PAY_DATE_COL).setValue(pay_date);
      }
      if (FREQ_COL > 0)
      {
        summarySheet.getRange(row, FREQ_COL).setValue(div_freq);
      }
    }
    Wait(12100); // max of 5 polygon calls per minute so wait 12.1 seconds after each call
    row++;
    ticker = summarySheet.getRange(row, 1).getValue();
  } while (ticker != "");
}

function GetDividend(ticker) 
{
  var response = UrlFetchApp.fetch("https://api.polygon.io/v3/reference/dividends?ticker=" + 
  ticker + "&limit=1&apiKey=" + POLYGON_KEY);
  var data = JSON.parse(response);
  if (data.results.length == 0)
  {
    return 0;
  }
  return data.results[0];
}

function Wait(ms)
{
  var d1 = new Date();
  do
  {
    d2 = new Date();
  } while (d2 - d1 < ms);
}

function YearsToMs(years)
{
  return years*365*24*60*60*1000;
}

function GetSheet(name)
{
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

