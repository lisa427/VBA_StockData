# VBA_StockData

This VBA code is designed to work with an Excel file that includes the following stock information: ticker name, date, open price, high price, low price, close price, and stock volume.  The data for each year is on separate sheets, a sheet for each year.

The code will go through each sheet and provide the following summarized data for each ticker in the columns next to the original data: ticker name, yearly change, percent change, and total volume for that year.  It will also provide the name of the ticker with the largest percent increase, the smallest percent increase, and the largest volume, along with the values.

There is no need to run the code separately for each sheet in the Excel file, if the code is imported into a module it will run through each sheet on its own.
