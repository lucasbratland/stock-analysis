# Stock Analysis with Excel VBA

## Overview of Project

### Purpose
There was a request by Steven to prepare a workbook for him analyzing the stock data provided. we used VBA to automate this process for Steven so that it was push button simple for him. We then were challenged to refactor the code to see if we could make it more efficeint compared to our original code. 


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








When isolating the theater crowdfunding category, projects that launched in May are the most successful with 66.9% reaching their goal. This is followed by June and July with 65.4% and 63.0% respectively. ![Chart](/Theater_Outcomes_vs_Launch_percent.png)

If you look at the subcategory of plays, things change a little bit. June is the most successful month with 70.3% followed by May and March with 68.9% and 68.7% each. Overall for the plays subcategory all months are above 61% success except for December which is only at 48.2%. 
![Chart](/Plays_Outcomes_vs_Launch_percent.png)

### Analysis of Outcomes Based on Goals
When looking at goal dollar amount of the plays subcategory the greatest success was had when the goals were less than $5,000. Goals under $1000 were 76% successful, while projects with goals in the $1000 to $5000 range were 73% successful. The next closest goal range was $35,000-$45,000 obtaining a 67% success rate. ![Chart](/Outcomes_vs_Goals.png) 

### Challenges and Difficulties Encountered

the main challenge that I encounter is that the launch date charts that were asked for, I felt didn't tell the whole story. To overcome this, I created a couple additional charts that I felt told that story more effectively. These charts are discussed below. 
## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?

The main conclusion is that the best time to launch a Theater category or specifically a Theater/Play category project is in May or June. These months are #1 and #2 for both the theater category and plays subcategory. On the other end of the spectrum the worst time to launch a crowdfunding campaign is in December. December is the worst month by at least 10% when compared to any other month when looking at just play campaigns. 
- What can you conclude about the Outcomes based on Goals?

The ideal goal range for projects in the plays subcategory is less than $5,000. Combined these 2 goal groups will be successful 73.4% of the time. Once you step up to the next goal bracket of 5-10k, you drop to a 55% success rate. 

- What are some limitations of this dataset?

It's hard to account for the amount of work that each project owner put into their crowdfunding campaigns. This effort would contribute to the success of a campaign and we can't account for it. 

- What are some other possible tables and/or graphs that we could create?

I did create a couple extra graphs that I found more useful in my analysis. I duplicated the theater vs launch chart, but I changed it to percentage instead. Comparing percentages is makes more sense to me than just comparing straight counts. It tells a more complete story, because you might have a higher success count in a month but if more campaigns fail then is it really the best month? I also created a plays vs launch date chart using percentages. It also might be beneficial to look at average number of backers by launch and by goal range. 