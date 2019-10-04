# Stock analysis with VBA and Excel

**Macro** = ticker_100319.bas  
**Test data** = alphabetical_testing.xlsx  
**Original stock data** = Multiple_year_stock_data.xlsx  !!! Warning, this is a large file containing ~2 million rows !!!  
**Stock data with summary tables** = Multiple_year_stock_data_solved.xlsx

## Summary
The VBA macro contained in "ticker_100319.bas" takes a list of stocks with daily open/close prices and creates the following:
- Summary table for each sheet, 2 columns to the right of the data:
  - Ticker
  - Change in stock price for the year, highlighting greed for positive and red for negative
  - Percent change in stock price for the year
  - Total stock volume traded

![Summary table](https://github.com/L0per/Excel-VBA-Analysis-of-Stocks/blob/master/Images/summarytable.gif?raw=true)
  
- Limits table for each sheet that summarizes the summary table, 2 columns right of the summary table:
  - Greatest % increase in price
  - Greatest % decrease in price
  - Most traded
  
![Limits table](https://github.com/L0per/Excel-VBA-Analysis-of-Stocks/blob/master/Images/limitstable.gif?raw=true)

## Limitations
1. Original data needs to be formatted with headers in the example below:

![Headers](https://github.com/L0per/Excel-VBA-Analysis-of-Stocks/blob/master/Images/headers.GIF?raw=true)

2. Tickers need to be grouped together.
3. Each sheet should contain a single year of data.

## Future Improvements
1. Currently, the macro runs by performing on the first ticker symbol and then entering a loop for the remaining symbols for each sheet. Ideally this would be performed within a single loop.
2. The macro is not efficient (2-3 minutes for ~2 million rows). This may be caused by the excel match function.
