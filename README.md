# VBA Challenge 

## Overview of Project
A VBA macro, initially designed to compare returns from green energy stock options, refactored and expanded to analyze the performance of the entire stock market over the last several years.

### Purpose
A financal consultant provided us with data on several green energy stocks seeking our assistance with creating a VBA macro to compare the returns on these stock options in 2017 and 2018 to better advise their client on which companies and stocks to invest in long-term. The workbook was subsequently programmed to display the resulting analysis to its viewer with ease, and the financial consultant's satisfaction with this inital result prompted them to ask that we expand the scope of the macro to analyze data on the entire stock market for the past several years. 

## Results

### Green Stocks Analysis
To provide on option for running the analysis on different years' datasets, the first prompt initiated by the macro asks for the desired year.


    yearValue = InputBox("What year would you like to run the analysis on?")


The 3012 rows of data are then categorized according to the stock option's abbreviated title (or ticker); in this dataset, there are 12 different green stock options.


    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"


The number of rows in the dataset is accounted for by the following:


        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists


The dataset is looped through and the desired variables of 'total volume', 'starting price', and 'ending price' are collected utilizing:


       '4) Loop through tickers
        For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5) loop through rows in the data
        Worksheets(yearValue).Activate
         For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            End If
           
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            End If
           
           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            End If

            Next j


The collected variables are then transferred into a new spreadsheet and sorted according to 'Stock Ticker', 'Total Volume', and 'Return'


     '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
       
    Next i


After formatting to colorcode the returns for each stock option, the analysis of the 2017 stock data shows:

![All Stocks Analysis 2017]()

When the macro is run again for the 2018 stock data, it shows the returns for these green stock options is much lower than 2017:

![All Stocks Analysis 2018]()

### Refactored Code Analysis
When VBA's timer function is implemented, the run time of the original script is: 

![Original Script Execution Time]()

After the code has been refactored, the run time of the script is: 

![Refactored Script Execution Time]()

## Summary

### Code Refactoring
There is a detailed statement on the advantages and disadvantages of refactoring code in general.

### VBA Script Performance
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script. 
