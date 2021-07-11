# Stock Analysis

## Project Overview
The purpose of this project is to refactor All Stock Analysis VBA script and make it more efficient. 

## Results
Refactoring the VBA script by resulted in reducing the time to run the analysis for all stocks for years 2017 and 2018. The below table shows the difference between run time for the two scripts.
Analysis Year | Original VBA Script | Refactored VBA Script
--------------|----------------------|----------------------
2017 | 1.01 | 0.23 
2018 | 1.06 | 0.55

### What caused the differences?

Refactored VBA Script: _Looping **once** through the rows made the code run faster_
```
For i = 2 To RowCount
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
       If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
       
            tickerStartingPrices(tickerIndex) = tickerStartingPrices(tickerIndex) + Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
         
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerEndingPrices(tickerIndex) = tickerEndingPrices(tickerIndex) + Cells(i, 6).Value
            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
        End If
    Next i
 ```
      
 Original VBA Script: _Nested For Loop_
 
 ```
 '4) Loop through the tickers.

    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    '5) Loop through rows in the data.
         Worksheets(yearValue).Activate
            For j = 2 To RowCount
    '5a) Find the total volume for the current ticker.
    If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
    '5b) Find the starting price for the current ticker.
    
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        startingPrice = Cells(j, 6).Value
        
    End If
    
    
    '5c) Find the ending price for the current ticker.
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value
    End If
    
    Next j

    
'6)Output the data for the current ticker.
Worksheets("All Stocks Analysis").Activate

Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


Next i
```
 
### Results for 2018 All Stocks Analysis

![AllStocksAnylysis2018](/Resources/AllStockAnalysis2018.png)

The results show negative return for most of the stocks. Only **ENPH** and **RUN** stocks had positive return. 


### Results for 2017 All Stocks Analysis

![AllStocksAnalysis2017](/Resources/AllStockAnalysis2017.png)

The results show positive return for all stocks except for **TERP**.

Comparing stock performance over 2017 and 2018, **ENPH** and **RUN** stocks are the most profitable. Also **ENPH** had the largest daily volume in 2018.

## Summary
### Advantages and Disadvantages of the Refactored VBA Script:
The advantages of refactoring VBA code include enhancing code performance, making it easier to understand, and supplementing one's code developing skills. For this project, refactoring the original VBA script enhances the time to execute analysis by more than 50%. The code is easier to understand and follow as it become shorter. On a personal note, I will use arrays and single loop through the data for in upcoming VBA projects.
On the other hand, refactoring is time consuming. It took me considerable amount of time to refactor the code and run it without errors ❗

### Advantages and Disadvantages of the original VBA Script:
The original VBA code responds naturally to the first logic that comes in mind when you need to run analysis over multiple stocks. It is easier to run without complications.
However, it is not efficient in terms on performance as it runs through the data more than once. If we want to scale it up to a larger dataset it will take much longer time and may even cause a computer jam 🔴
