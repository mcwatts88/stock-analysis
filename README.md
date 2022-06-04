#VBA of Wall Street

##Overview of Project

Steve reached out with the goal of analyzing stock data to help build an investment portfolio for his parents. His parents want to invest in a green energy company named DAQO New Energy Corp. Steve wants to diversify for the portfolio to include other green energy companies. With a provided list of stock data for the years of 2017 and 2018 tabulated by ticker id, analysis was performed to determine which stocks Steve should include in the portfolio.

##Results

###Nested For Loop Analysis
The original analysis of the stocks utilized an array of ticker ids of specific companies to compile the yearly data for each company in the array and display the company's performance. This was done by iterating the array index with a for loop, then using a nested for loop to search the data for the applicable data. An example of the code is below.

    ```
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
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
    ```
In order to determine how optimal this script is, a timer was initiated in the script. This operation completed in approximately 0.7 seconds for each range of data. See the below screenshots.

![VBA_Challenge_2017_Original.png](https://github.com/mcwatts88/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)
![VBA_Challenge_2018_Original.png](https://github.com/mcwatts88/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

###Refactored For Loop Analysis

The above code was refactored to iterate a variable called tickerIndex in order to build output arrays. This allowed the loop to only have to excecute to completion a single time. An example of the refactored code is below

    ```
    '1a) Creates a ticker Index
    Dim tickerIndex As Integer
    
    tickerIndex = 0

    '1b) Creates three output arrays
    Dim tickerVolumes(11) As Long
    
    Dim tickerStartingPrices(11), tickerEndingPrices(11) As Single
    
    
    '2a) Creates a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
    
    Next i
        
    '2b) Loops over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increases volume for current ticker
        ticker = tickers(tickerIndex)
        
        If Cells(i, 1).Value = ticker Then
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        End If
        
        '3b) Checks if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = ticker And Cells(i - 1, 1) <> ticker Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
            
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
                
        If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
             
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    
            '3d Increase the tickerIndex.
                
            tickerIndex = tickerIndex + 1
            
        End If
        
        
    Next i
    
    '4) Loops through arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    ```

Not having to perform the loop multiple times allowed the script to complete much faster. The screenshots below show the time the script took to execute.

![VBA_Challenge_2017.png](https://github.com/mcwatts88/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](https://github.com/mcwatts88/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

##Summary

###Refactoring in General
    
Refactoring code can be advantageous in that it optimzes your code to run more efficiently. It can also be used to add features that were previously not included. The downside to this is that it takes time and money. The refactoring may only provide a marginal increase in optimization but may take quite a while to execute.

###Refactoring This Script

This script was refactored to lower the time it takes to run the analysis. This increase in efficiency gained us over a half second in processing time. This lowers the resources used by the script and and runs quicker which is very advantageous. The use of arrays simplifies the code so that it isnt necessary to keep track of nested loops. There are little downsides to this refactoring, as the only significant change is iterating the index inside of an existing for loop instead of using a separate for loop.

