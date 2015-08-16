# YahooFinanceForExcel
A collection of Excel UDF's that enables data extraction from Yahoo Finance directly into Excel

### Project Description
This project enables Excel users to import data from Yahoo Finance into Excel using User Defined Functions. The functions are array functions which means they can be applied to a range of your choice in Excel.

### Functions

#### YahooGetHistoricalPrices:
This function will get a list of historical prices from Yahoo Finance for you. It takes 4 parameters.  
Symbol: This is the symbol to search for, e.g., ^GSPC for the S&P 500, or AAPL for Apple.  
Granularity: Here you can specify if you want daily data with "d", weekly data with "w" or monthly data with "m".  
Start date: This is the date you want the data from.  
End date: This is the date you want the data to.  

The returned array is not in any way interpreted by the addin, this means that data might look at little bit funny on some localizations of Excel, e.g, the Danish one where you write 12.345.678,90 instead of 12,345,678.90. In this case Excel has another function called NumberValue, that can help you parse the dates.

#### YahooGetQuoteDetails:
Will attempt to fetch the latest quote details for a given symbol. This includes the name, symbol, last price, opening price, previous closing price, and the high and low price. The function only takes on parameter which is the given symbol.  
Symbol: The symbol to fetch the data for, e.g., AAPL for Apple.

#### YahooGetSectorOverview()
This function will return an overview of different sectors on Yahoo finance, containing different data such as, price change, P/E, Price to Book value, etc. This function takes no parameters.

### Running the add-in:
For one time use you can either run the 32 bit or 64 bit version of the .xll file depending on which version of Office you are running. After accepting the security warning you can use the UDF's described above.

If you plan to use the add-in across several times on the same machine you can.  
1) Open Excel  
2) Go to Settings  
3) Go to Add-in Settings  
4) Press execute at the bottom next to administer excel addins  
5) Select Browse  
6) Navigate to the folder where the addin is located and click ok.  

### Using the addin:
Since all of the functions are array functions they must be applied as such. You select a range in which you want the data. Then you enter the formula in the formula field and evaluate the formula using CTRL+SHIFT+ENTER.

### Frequently Asked Questions:

#### Q: I see a lot of #N/A values in my sheet when applying the UDFs. What is causing this?
#### A:
This is most likely because the range to which you are applying the formula is larger than the amount of values that is returned by the UDF function. You can reduce the area to which you apply the function or just ignore them if you expect the data set to increase in the future.

#### Q: I see alot of name fields in the my Excel spreadsheet. What is causing this?
#### A:
This is due to the addin not being loaded properly. Make sure that you have loaded the addin properly. Also make sure you have the right version of addin-loaded, if you run a 64 bit version of Office, you must use the 64 bit version of the addin. If you run a 32 bit version of Office, you must use the 32 bit version of the addin.
