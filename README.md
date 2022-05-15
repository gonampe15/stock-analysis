# All Stock Analysis

## Project Overview

In this project, we are helping Steve to analyze and collect the desired data from the stocks dataset with the end purpose of find out what stocks are worth investing in. 
In order to do that, we'll be using VBA to write code that can make the data retrieval process pretty simple and easy. 

## Results

```
'2) Initialize array of all tickers
   
   
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
   
   
'3a) Initialize variables for starting price and ending price
   
   Dim startingPrice As Single
   
   Dim endingPrice As Single
   
'3b) Activate data worksheet
   
Worksheets(yearValue).Activate

'3c) Get the number of rows to loop over
   
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

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

## Summary

### Refactoring Code Advantages/Disadvantages

The major advantage of refactoring code is making the code more efficiently by taking less steps, reducing the amount of memory used and improving the logic of the code. The major disadvantage of refactoring code is that you could potentially break code that already works. 

### VBA Script Refactoring

I believe one of the major advantages of refactoring code in VBA Script is that you can keep and use as much of the original code as you want. The major disadvantage of refactoring code in VBA Script is having a good understanding of the code syntax. I feel like the syntax matters so much more when you are trying to make the code more efficient.
