# All Stock Analysis

## Project Overview

In this project, we are helping Steve to analyze and collect the desired data from the stocks dataset with the end purpose of find out what stocks are worth investing in. 
In order to do that, we'll be using VBA to write code that can make the data retrieval process pretty simple and easy. 

## Results

In order to make the refactored code more efficient, we created the variable tickerIndex which helped us assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each one of the tickers before interating through the dataset. Here is the comparison of both the original and refactored code, as well as the run time per year for the original and refactored code.

### Old Code

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
### Refactored Code
```
    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    
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
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

        '3d Increase the tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
            
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
       
            Cells(4 + i, 2).Value = tickerVolumes(i)
            
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
            
    Next i
```
### Run time Original Code

![2017](/Resources/VBA_Oldcode_2017.png)
![2018](/Resources/VBA_Oldcode_2018.png)

### Run time Refactored Code

![2017](/Resources/VBA_Challenge_2017.png)
![2018](/Resources/VBA_Challenge_2018.png)


## Summary

### Refactoring Code Advantages/Disadvantages

The major advantage of refactoring code is making the code more efficiently by taking less steps, reducing the amount of memory used and improving the logic of the code. The major disadvantage of refactoring code is that you could potentially break code that already works. 

### VBA Script Refactoring

I believe one of the major advantages of refactoring code in VBA Script is that you can keep and use as much of the original code as you want. The major disadvantage of refactoring code in VBA Script is having a good understanding of the code syntax. I feel like the syntax matters so much more when you are trying to make the code more efficient.
