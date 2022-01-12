# stock-analysis

## Project Overview

The purpose of this project is to help Steve better determine which stock is the best option for his parents. By comparing the daily volume and yearly return for each stock option, this will aid Steve in making a more educated decision. In addition, we will build our script as well as refactor script that has been provided in order to compare the execution times.

## Stock Performance Results

By formatting the data to highlight positive return values in green and negative return values in red, the significant differences between 2017 and 2018 are very distinct.
We can see that the return was much more successful for most Stocks in 2017 vs 2018. We can also see that both "ENPH" and "RUN" had a positive return in both years.

![Stock_Analysis_Outputs_2018_](https://github.com/aidan2013/stock-analysis/blob/main/Resources/Stock_Analysis_Outputs_2018_.png)

![Stock_Analysis_Outputs_2017_](https://github.com/aidan2013/stock-analysis/blob/main/Resources/Stock_Analysis_Outputs_2017_.png)

#### **Refactored Script**

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

These pieces of the script provide the data needed to compare the stock analysis outputs for all stocks in 2017 and 2018.

```
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
```

#### **Original Script**

The original scipt has a nested loop instead of two individual loops. 

```
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
```
#### **Execution Times**
A timer was built into the script in order to calculate the time it takes to execute the script. The start time is triggered when the year is entered into the input box.
```
Worksheets("All Stocks Analysis").Activate
    
    Dim startTime As Single
    Dim endTime As Single
       
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
```

The end timer is set to conclude when the script ends.

```
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
In comparing the execution times for the orginal and refactored scripts, we can see that the refactored script has a quicker execution.

**Refactored Script Execution Time (2017)**

![VBA_Challenge_2017](https://github.com/aidan2013/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

**Refactored Script Execution Time (2018)**

![VBA_Challenge_2018](https://github.com/aidan2013/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

**Original Script Execution Time (2017)**

![image](https://user-images.githubusercontent.com/91445591/149075265-073b784e-587e-4177-9046-a3ecc79ac7e1.png)

**Original Script Execution Time (2018)**

![image](https://user-images.githubusercontent.com/91445591/149075144-aa06991b-8db0-4117-885e-803700f15025.png)

## Summary

In general, there are advantages and disadvantages to refactoring code. By refactoring code, you can remove redundancies and make the code faster and more efficient. This also helps to make the code easier to read and understand. Although these are great benefits, refactoring code can also lead to breaking the code which may be very time consuming to fix. Depending on how large the script may be, the benefits may outweight the time spent on refactoring the script.

In refactoring the VBA script in this project, I ran into a challenge where the script was no longer bringing in the correct data. In an effort to solve for this, I ended up making it worse and having to dig deeper into solving the error. In order to solve ther error, I reviewed each piece of the code in order to make sure it all ties together, there were no spelling errors or redundancies. With a second pair of eyes from a TA, we were able to locate the errors and get the script to run successfully. This took up alot of time however in the end, the script execution was much faster than the original script.

