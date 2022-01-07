# Stock Analysis with Excel VBA

## Overview of Project

### Purpose
There was a request by Steven to prepare a workbook for him analyzing the stock data provided. we used VBA to automate this process for Steven so that it was push button simple for him. We then were challenged to refactor the code to see if we could make it more efficient compared to our original code. 


## Results

### Analysis
Originally when we created the VBA code for this project, we used 2 nested For loops. 

     'Loop through the tickers.

    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
    'Loop through rows in the data.

        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
        
    'Find the total volume for the current ticker.

            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
    'Find the starting price for the current ticker.

            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                '  set starting price
                startingPrice = Cells(j, 6).Value
            End If
            
    'Find the ending price for the current ticker.

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1) <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j

    'Output the data for the current ticker.
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
    

This proved effective for the data set we were given.  

![2017](/resources/VBA_Challenge_2017_Before_Refactoring.png) 
![2018](/resources/VBA_Challenge_2018_Before_Refactoring.png)

This is a small sample set and we can do better. We refectored the code to use more arrays and only one For loop. 
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
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
       
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            '  set starting price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            'set Ending Price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
        'End If
        End If
        
    Next i

Doing this cause a much better result

![2017](/resources/VBA_Challenge_2017.png) 
![2018](/resources/VBA_Challenge_2018.png)

This lower the time taken by over 80% for both years. The percentage could increase a small bit if you added the formatting into the original (pre refactoring) code. 


## Summary

### Advantages and Disadvantages of Refactoring

One advantage of refactoring the code is that it allows you to make the code more efficient. It also gives you the chance to make the code more readable to others. It can also give you the chance to make the code more scalable to handle bigger datasets. 

A disadvantage is that you are spending time on refactoring something that is already working, leads to the old saying "If it ain't broke don't fix it" refactoring could also lead to a whole new list of bugs and other potential issues. 

### Pros and Cons of Refactoring the original VBA script

A pro of refactoring the original VBA script was the decrease in time that it takes to run the script by 80%. It is also easier to follow as there is not a nested for loop. 

A con of the refactoring is that if you aren't familiar with how arrays work, it can be more difficult to read. Another con is that we didn't add a sort function to the code. A sort of the ticker and date columns should've been added in to make sure the starting and ending price variables price out properly. 