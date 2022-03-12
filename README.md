# Stock Analysis: Excel- VBA

## Overview

### History of the Project

In the previous project, stock data of green energy for 2017 and 2018 were analyzed to determine if the investment was worthwhile. For that reason, the VBA programming language had used to analyze the small dataset programmatically. 
  
### Purpose of the Analysis

For this project, the original code was refactored to increase the efficiency of working on big datasets. Automating analyses with programming reduces the chance of errors and running time, especially with big datasets and repetitive analyses.  

## Results

### Original Code 

* The dataset that was worked on displays information on 12 different stocks for in years 2017 and 2018. [VBA_Challenge Excel File](https://github.com/duygusimsek/stock_analysis/blob/main/VBA_Challenge.xlsm)
* “All Stock Analysis” subroutine had been created.
* To measure the code performance of running time, startTime and endTime variables were set to Timer function.
* The code for InputBox, activation for worksheet, creating headers of the chart, arrays for 12 tickers were written. 
* “For loops” was created to find stock changes for the defining year in the dataset.
* A ticker array was created to go through rows in the dataset. [AllStockAnalysis_OriginalCode](https://github.com/duygusimsek/stock_analysis/blob/main/AllStockAnalysis_OriginalCode.bas)
* Lastly, for formating the worksheet, “font-style”, “line-style”,”number-format” and “columns-autofit” were written.
* To run the analysis easily, the button function (Run Analysis For All- Original Code) was used on the worksheet. 

### Refactoring Code
	
* “All Stock Analysis” subroutine was changed to “All Stocks Analysis Refactored”.
* InputBox, worksheet activation, timer function, headers row, and 12 tickers arrays were copied from the original code.
* tickerIndex variable was created and set zero. 
* tickerVolumes, tickerStartingPrices, and tickerEndingPrices output arrays were created with their data types. 
* tickerIndex variable was used to access the four arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices)
* To loop over all the rows in the worksheet, for loop was created.
* tickerVolumes was set to zero and using tickerIndex as the index, it was used to increase volume for the current tickerVolumes.
* If-Then statements were written to check if the current row is the first row and the last row.
* To get output for the three columns on the worksheet, for loop was created and loop over the four arrays. 
* For font-style, line-style,number-format, and columns-autofit coding parts were copied from original code. [AllStocksAnalysisRefactored_Code](https://github.com/duygusimsek/stock_analysis/blob/main/AllStocksAnalysisRefactored_Code.bas)
* To run the analysis easily, the button function (All Stock Analysis Refactored) was used on the worksheet. 
* After running the refactored code, the time performance of the code stock data analysis outputs for in years 2017 and 2018 are given below. 
 
 ![Image for 2017](https://github.com/duygusimsek/stock_analysis/blob/main/Resources/%20VBA_Challenge_2017.png)
 ![Image for 2018](https://github.com/duygusimsek/stock_analysis/blob/main/Resources/%20VBA_Challenge_2018.png)
 
 ## Summary
 
1. There are some advantages and disadvantages to refactoring a code.  Refactored code is more organized (clean code), therefore easy to read and understand, especially for other people than the creator. Once the code is ready to run, It helps to save time and money (because of usage of less memory footprint), and especially for big datasets running time of code is much faster. Although, the refactoring process is time-consuming and during the progress, mistakes can be made. 
2.  The advantage of refactoring our original VBA script is optimizing run time. The refactored code ran approximately 3 times faster than the original code. However, for this size of the dataset, the refactoring process took more time than it benefits.

