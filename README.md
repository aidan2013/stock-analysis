# stock-analysis

## Project Overview
The purpose of this project is to help Steve better determine which stock is the best option for his parents. By comparing the daily volume and yearly return for each stock option, this will aid Steve in making a more educated decision. In addition, we will build our script as well as refactor script that has been provided in order to compare the execution times.

## Results


### Refactored Script

In order to compare the daily volume and yearly return, we will loop through the data and add up the daily volume for each stock. It then will loop through the rows and bring in the stock starting price and ending price in order to calculate the yearly return for each stock. 

First, we will need to determine which dataset we will be looping through. In this case, we have a dataset for 2017 and 2018.
An input box will display for the year to be entered.

 ``` 
 yearValue = InputBox("What year would you like to run the analysis on?") 
 ```
 
In order to loop through each stock, we will set an array for each stock ticker.

```
    Dim tickers(11) As String
    
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
```

We will create an output array for the total volume, starting price and ending price in order to return a value for each ticker.

```
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
```

Since we want to get a sum of all the daily volumes for each ticker, we will need to start all the tickers at zero

```
For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
```

We then loop through all the rows in the indicated spreadsheet to sum the daily values of each ticker and locate the starting and ending prices of each ticker.

```
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           
           End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
          
            
        'End If
        End If
    Next i
```

These pieces of the script provide the data needed to compare the stock analysis outputs for all stocks in 2017 and 2018

![Stock_Analysis_Outputs_2018_](https://)
