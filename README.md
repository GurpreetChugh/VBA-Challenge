This assignment consisted of analyzing a excel containing stock data for 3 years 2018-2020 using VBA coding. 

The raw data consisted of daily opening, low , high, close price and total volume traded for the stocks.

The first task was to create a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

Using for loop for each row we used the conditional statements to compare the value of ticker in the current row to the following row and if the ticker value is different (meaning the current row is the last day record for a ticker), we recorded the ticker and calculated the yeraly and % change of opening price on first day to the closing price on the last day and calculated the total volume traded in a year.

Based on the data, conditional formatting was applied that will highlight positive change in green and negative change in red in the calculated fields Yearly Change and % Change.

We then find the ticker with maximum % gain, maximum % loss(or min % gain) and maximum total traded volume in a year

We used the inbuilt WorksheetFunctions to calculate Max and Min Value. We created functions (GetMaxRow and GetMinRow) to get the row reference and hence row number to find the ticker corresponding to these max and min values and recorded them along with the respective values.

Finally we added the code to run this script on every worksheet i.e. for each year 2018, 2019 and 2020, so that by running the script once, we could get the analysis for the sheets in one shot.