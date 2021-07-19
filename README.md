# Stock-analysis

## Overview of Project
### Purpose
The scope of this project was to refactor a VBA code in order to collect stock market data insights presented in data set for the year 2017 and 2018. In this exercise should be established if refactoring the proposed code has improved the execution time of the VBA script.

    '1a) Create a ticker Index
     tickerIndex = 0

    '1b) Create three output arrays
     Dim tickerVolumes(12) As Long
     Dim tickerStartingPrices(12) As Single
     Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
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

## Results
I opened the starter code file in Visual Studio Code, and I have edited it. Below are listed the steps with the description.
 
<img src="https://i.ibb.co/P1KmdL1/VBA-Challenge-2017.png" alt="VBA-Challenge-2017" border="0">

In 2017 there is a positive rate on return on all stocks except Ticker TERP which has a negative return of 7.2%. 

<img src="https://i.ibb.co/3mz1bRZ/VBA-Challenge-2018.png" alt="VBA-Challenge-2018" border="0">

The 2018 year appears to be less successful. There are only 2 Tickers with a positive return on stock: the ENPH with a rate of 81.9%, and the RUN with a rate of 84.0%. The rest of the Tickers experience a negative rate on return on stocks.
Another finding is that the execution time of the refactoring script has decreased in comparison with the execution time of the original script.

## Summary
The advantages of code refactoring are helping to find bugs, decreasing execution time, improving code readability and reducing complexity. 
However, there are disadvantages also. Refactoring a code is time consuming. You never know how much time it may take to execute the process. There is also a chance of mistake. 
