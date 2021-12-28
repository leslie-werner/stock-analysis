# VBA Challenge
# VBA Stock Analysis

## Overview of Project

### Purpose
  In this analysis we wanted to help Steve look at datasets that coved the entire stock market for 2017 and 2018. Using this information, he wanted to help his parents be able to pick stocks that could be beneifical in the long run. Using code we had developed throughout the module, we refractored code to be able to analyze the dataset in a shorter amount of time. 
  
## Analysis
'Using images and examples of your code, compare the stock performance between 2017 and 2018. As well as the execution times of the original script and the refactored script.
### Results
When looking at stock performaces between 2017 and 2018, we can see that 2017 was a more successful year for stocks compared to 2018. 
In both 2017 and 2018, we can see that ENPH and RUN remained in the positive percent returns while TERP, was unsuccessful in both years. With the reuslts obtained, it is recommened that Steves' parents invest in stocks with positive returns for both years if a less risky approach is wanted. 
### Refactoring The Code
In this challenge we wanted to make the code we developed into a more efficent code that would reduce both memory and time to run.

First we had to make a ticker Index that we could use in different arrays.
             
    '1a) Create a ticker Index
    tickerIndex =0
    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
Then we had to create a for loop that initialized the Volume of tickers to zero so we could then use tickerIndex to increase the current stock ticker


      ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For j = 0 To 11
        tickerVolumes(j) = 0
        Next j
       ''2b) Loop over all the rows in the spreadsheet.
       For j = 2 To RowCount
              
           ' If the next row’s ticker doesn’t match, increase the tickerIndex.
           
              '3a) Increase volume for current ticker
              tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value


Next we had to use the tickerIndex to make sure that we were using the correct ticker when using the if-then condition

           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

               tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
               
          'End If
           End If

        '3c) check if the current row is the last row with the selected ticker
        'If  Then
           If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        '3d) add tickerIndex
                tickerIndex = tickerIndex + 1
          'End If
           End If
           
       Next j
  
Finally, we had to loop through the array to have the outputs on the correct worksheet.   
  
  
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For j = 0 To 11
           Worksheets("All Stocks Analysis").Activate
           tickerIndex = j
           Cells(4 + j, 1).Value = tickers(tickerIndex)
           Cells(4 + j, 2).Value = tickerVolumes(tickerIndex)
           Cells(4 + j, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

#### 2017 Code Run Time
<img width="266" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/65638310/147577498-b76f4114-6c47-4111-8552-1956d5ab7d87.png">   

### 2018 Code Run Time
<img width="265" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/65638310/147577603-0d7761ad-ec99-45bc-b664-13a6eef0c861.png">

### Challenges and Difficulties

## 
