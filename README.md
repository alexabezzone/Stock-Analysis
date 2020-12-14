# Stock Analysis Challenge

## Analysis

### Overview of Project

Our client Steve needs assistance analyzing stock from the company DAQ0 New Energy Corp for his parents who recently invested in the company. To get a full understanding of the value of the stock and to potentially diversify the funds for his parents, Steve researched other greenhouse company stock data to compare to DAQ0. Steve has asked us to analyze the stock data through VBA. After completing the first round of VBA analyzation, Steve was happy with the results of the analysis. At a single click, he was able to fully understand the data of the entire stock data spreadsheet. Now, Steve has asked us to help him find a solution to analyze the entire greenhouse stock market over the last couple years. To do this, we must refactor our current code to make it fast enough to review the stock data from over the past couple of years. 

### Results

In order to help Steve analyze the stock data from the past couple of years at a faster pace then our original VBA code, we will refactor the code to loop through the data individually and attain the information. The following below are steps on how to refactor the code:

The first step of refactoring our new code was to set the `tickerIndex` to zero and then iterating over the rows of data. 
    `For i = 0 To 11
    
    tickerIndex = tickers(i)
    
    Next i`
    
   The second step is to define the 3 output arrays and assign them a variable type. 
   
    `Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single`
    
 The third step is to create a for loop that sets tickerVolumes to zero and a for loop to loop through the rows in the dataset. 
    
    ` Worksheets(yearValue).Activate
      tickerVolumes = 0
    
      For j = 2 To RowCount`
      
The fourth steps are to increase the volume for the current ticker and add the tickerVolume to the current ticker with an If Then         statement. 
      
      If Cells(j, 1).Value = tickerIndex Then
        
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
            
         
   The fifth step is to write a conditional If Then statement to determine the starting price of the ticker index. 
   
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
    
    The sixth step is to write a conditional If Then statement to determine the ending price of the ticker index. 
    
       If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
    
    Finally to output the Ticker, Total Daily Volume, and Return, loop through the arrays to output the data. 
    
    Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickerIndex
       Cells(4 + i, 2).Value = tickerVolumes
       
  As a result, the running time of the solution for year 2017 was .45 seconds and the running time for 2018 was .11 seconds. 
    
  ### Summary
There are many advantages and disadvantages to refactoring code. An advantage is that it allows the code to perform faster and can be more universal for future Data Analysts to incorporae the data into their own projects. Another advantage is that it can make the code more efficient and solve the problem with less steps than the original code. A disadvantage of refactoring code is that it could be time consuming and the goal of the solution may get lost behind the code.

The original VBA script that was solved for our client Steve allowed him to analyze a few stocks at a time in an efficient manner, however, it did not allow him to look at a plethora of greenhouse stocks from the past couple of years. The refactored code allowed Steve to analyze multiple stocks from previous years at a more efficient and faster pace than the original code allowed. Other coders can also adapt the refactored code to their own VBA stock analysis. 
  
  

