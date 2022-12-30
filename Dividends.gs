// edit these variables
var SPREADSHEET_ID = "";
var SUMMARY_SHEET_NAME = "Summary";

var POST_DIV_COL = 10;
var POST_FREQ_COL = 17;
var POST_EX_DATE_COL = 22;
var POST_PAY_DATE_COL = 23;
var POST_PAYOUT_RATIO_COL = 28;
var POST_YEARS_INCREASING_COL = 20;
var POST_CAGR_COL = 26;
var FIRST_TICKER_ROW = 2;
var TICKER_COL = 2;

var VAR_SHEET_NAME = "VarSheet";
var GET_TICKER_COL = 1;
var GET_TICKER_ROW = 2;
//

function AddDividends()
{
  var start = new Date();
  var end = start.getTime() + MinutesToMs(5.5); // these scripts can only run a max of 6 minutes
  var summarySheet = GetSheet(SUMMARY_SHEET_NAME);
  var varSheet = GetSheet(VAR_SHEET_NAME);
  var row = varSheet.getRange(GET_TICKER_ROW, GET_TICKER_COL).getValue();
  var ticker = summarySheet.getRange(row, TICKER_COL).getValue();
  do
  {
    var data = GetDividend(ticker);
    var response = data[0];
    if (response != 0 && response != null)
    {
      var date = new Date(response.ex_dividend_date);
      var div = "";
      var ex_date = "";
      var pay_date = "";
      var div_freq = 0;
      // ensures latest dividend data is current
      if (new Date() - date < YearsToMs(1))
      {
        var date_index = 0;
        while (date > new Date())
        {
          response = data[date_index];
          date_index++;
          date = new Date(data[date_index].pay_date);
        }
        div_freq = response.frequency;
        var index = 1;
        while (div_freq == 0)
        {
          if (index >= data.length)
          {
            return;
          }
          else
          {
            response = data[index];
            div_freq = response.frequency;
            index++;
          }
        }
        div = response.cash_amount * div_freq;
        ex_date = response.ex_dividend_date;
        pay_date = response.pay_date;
        if (POST_DIV_COL > 0)
        {
          summarySheet.getRange(row, POST_DIV_COL).setValue(div);
        }
        if (POST_EX_DATE_COL > 0)
        {
          summarySheet.getRange(row, POST_EX_DATE_COL).setValue(ex_date);
        }
        if (POST_PAY_DATE_COL > 0)
        {
          summarySheet.getRange(row, POST_PAY_DATE_COL).setValue(pay_date);
        }
        if (POST_FREQ_COL > 0)
        {
          summarySheet.getRange(row, POST_FREQ_COL).setValue(div_freq);
        }
        if (POST_PAYOUT_RATIO_COL > 0 && div > 0)
        {
          summarySheet.getRange(row, POST_PAYOUT_RATIO_COL).setFormula("=" + div + "/GOOGLEFINANCE(\"" + ticker + "\", \"eps\")");
        } 
        if (POST_CAGR_COL > 0)
        {
          var cagr = Get5YearCAGR(data, div_freq);
          if (cagr != 0)
          {
            summarySheet.getRange(row, POST_CAGR_COL).setValue(cagr - 1);
          }
        }
        if (POST_YEARS_INCREASING_COL > 0)
        {
          summarySheet.getRange(row, POST_YEARS_INCREASING_COL).setValue(GetYearsIncreasingDiv(ticker));
        }
      }
    }
    
    var now = new Date();
    if (now.getTime() > end) // store next ticker row and end (out of time)
    {
      varSheet.getRange(GET_TICKER_ROW, GET_TICKER_COL).setValue(row + 1);
      return 0;
    }  
    Wait(12100); // max of 5 polygon calls per minute so wait 12.1 seconds after each call
    row++;
    ticker = summarySheet.getRange(row, TICKER_COL).getValue();
  } while (ticker != "");
  varSheet.getRange(GET_TICKER_ROW, GET_TICKER_COL).setValue(FIRST_TICKER_ROW);
}

function GetDividend(ticker) 
{
  var response = UrlFetchApp.fetch("https://api.polygon.io/v3/reference/dividends?ticker=" + 
  ticker + "&limit=24&apiKey=" + POLYGON_KEY);
  var data = JSON.parse(response);
  if (data.results.length == 0)
  {
    return 0;
  }
  return data.results;
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

function MinutesToMs(min)
{
  return min*60*1000;
}

function Get5YearCAGR(array, divs_per_year) {
  const divArray = [];
  const yearArray = [];
  if (array.length < divs_per_year * 6)
  {
    return 0;
  }
  for (var i = 0; i < 6; i++)
  {
    yearArray[i] = 0;
  }
  for (var i = 0; i < divs_per_year * 6; i++)
  {
    divArray[i] = array[i].cash_amount;
    var index = Math.trunc(i / 4);
    yearArray[index] += divArray[i];
  }

  var newDiv = yearArray[0];
  var oldDiv = yearArray[5];
  var ratio = newDiv / oldDiv;
  var cagr = Math.pow(ratio, (1/5));
  return cagr;
}

function GetYearsIncreasingDiv(ticker)
{
  var text = UrlFetchApp.fetch("https://www.finance.yahoo.com/quote/" + ticker + "?p=" + ticker).getContentText();
  var index = text.indexOf("Currency in");
  var searchString = text.substring(index - 100, index + 100);
  var nasdaqIndex = searchString.indexOf("NasdaqGS");
  var nyseIndex = searchString.indexOf("NYSE");
  var exchangeStr = "";
  if (nasdaqIndex > -1)
  {
    exchangeStr = "NASDAQ";
  }
  else if (nyseIndex > -1)
  {
    exchangeStr = "NYSE";
  }
  if (exchangeStr == "")
  {
    return 0;
  }
  text = UrlFetchApp.fetch("https://www.marketbeat.com/stocks/" + exchangeStr + "/" + ticker + "/dividend/").getContentText();
  index = text.indexOf("Dividend Increase ");
  searchString = text.substring(index, index + 300);
  index = searchString.indexOf("Year");
  var backIndex = 1;
  var increaseYears = 0;
  var lastIncreaseYears = 0;
  do
  {
    backIndex = backIndex + 1;
    lastIncreaseYears = increaseYears;
    increaseYears = parseInt(searchString.substring(index - backIndex, index - 1));
  }
  while(!isNaN(increaseYears) && increaseYears >= lastIncreaseYears);
  return lastIncreaseYears;
}

function GetSheet(name)
{
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

