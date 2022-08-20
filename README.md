# StockAnalysis With Excel Sheet
#### Click here to view Data File: [VBA_Challenge.xlsm](https://github.com/jbailey2705/StockAnalysis/blob/main/VBA_Challenge.xlsm)
# Project Overview
## Purpose
### The purpose of this project, to refactor a VBA Code to collect stock information for the 2017 & 2018 stock years & determine the valuation or devaluation of the stocks for the given years to determine if the stocks being considered were worth investing in.
## Data
### The data presented was based on 12 stock tickers that were being considered over the 2017 & 2018 years. Data used in the sheet looked at the value of each ticker, date issued, highest & lowest price points, opening & closing prices, & volume of the tickers.
# Results
## Analysis
### I began refoacoting the code, first by copying the code needed to build on, and add revised the code to be able to pull the relevant data including, input box, chart headers, ticker arrays. 
### I've inluded the instruction code that has been provided & written in the file.

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i

    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i

# Summary
## Refactoring the code Pros & Cons
### Refactoring the code help to clean up the data, especially if it's going to be reused. Save on space on the CPU, decrease run times for the user to quickley gather the information needed. Some disadvantages may inlude, enormouse files, or not having the right apps or programs to run the refactored code. 
## Advantage of refactoring the Stock Analysis code
### For it's purpose, consumers of the information can generate the data quicker & cleaner. Updating the tickers on the stocks you want to analyze will pull the same content without having to repurpose the code from scratch, although this could be done from the original code, small tweaks to existing code will save consumers of the code a ton of time.
