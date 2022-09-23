# Stock-Analysis

## Overview of project
  -The purpose of this project was to code in VBA to determine what stocks are worthwhile investing in, from a set of stock data.
  -On top of returning positive and negative numbers, further visualisation was added with colors for even easier decision making when the results return.

### Results
  -This is the code that I used to refactor the previous VBS formulas. Along with screenshots of the final text box after successfully running the code.
  
  '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 6).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
        
![This is an image](C:\Users\Marco\Desktop\Resources)
![This is an image](C:\Users\Marco\Desktop\Resources)

## Summary

### Refactoring Pros
  -Refactoring allows for clearer and optomized code to run. 
  -It is a leaner version that runs faster and is easier to troubleshoot for other analysts to see and improve upon
  
### Refactoring Cons
  -I can only see that large data sets would be difficult to refactor since we do not have all the information at our disposal.
  -Potentially dynamic daata sets that are constantly changing cannot be refactored and as such will have unoptomized code. 
  
### Refactoring original code
  -After refactoring the original code the first and obvious benefit was the increased speed that the code was executed. 
  -The original code would run at about 1.5 seconds initially while the new refactored code ran much faster around under 0.25 seconds for both years. 
  -I cannot see any cons while refactoring this data set as everything went as expected and the data was large enough to contorl the outcome. 
