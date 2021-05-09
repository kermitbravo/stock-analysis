# Module 2 Challenge: Stock Analysis
*Performing financial analysis on specific **stocks** to help Steve and his family make informed investment and portfolio diversification decisions. *

## Table of Contents

|Contents                   |
|---------------------------|
|1. Overview of the Project |
|2. Results                 |
|3. Summary                 |

## Overview of the Project

### Purpose 

For this project we analyzed stocks for 12 Companies by using data for 2017 and 2018 in Excel leveraging VBA. 

As part of this analysis, we calculated the Total Daily Volume and Returns to help Steve and his family make investment decisions.

After delivering a workbook enabled with VBA macros that works for these 12 companies. Steve decided he wants to expand the scope of the analysis to all stocks in the stock market. To ensure that our code will scale and suffice the new requirements, we will have to refactor our code to optimize it to run for a larger dataset. 

## Results

After analyzing our code, we identified that even when it was working as intended, there were better ways to perform the analysis. Refactoring our code, would allow us to expand the scope of the analysis in the future to all the stocks in the stock market. 

Refactoring the code would allow us to speed up the calculation process, make our code more dynamic and ensure that our compute resources would be better utilized. 

### Observations

After taking a deeper look into our code, we decided to introduce an Index variable that would allow us to remove the nested loop in our VBA code and to leverage arrays to store the Daily Volumes and calculate the Returns. 

Although for/loops are very convenient and powerful, they usually impact the performance of the code. Hence why it's important to try to avoid nesting loops if it can be avoided.

Specifically, we are referring to this section of the code where we could strip one of the for loops out to improve the performance of our calculations:

```vba
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        For j = rowstart To RowEnd
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
        
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
        
            End If
        
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                endingPrice = Cells(j, 6).Value
        
            End If
            
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i
```
By leveraging the tickerIndex, we could achieve the following:

```vba
    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPirces(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = tickerIndex To tickerVolumes(12)
        tickerVolumes(i) = 0
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker does not match, increase the tickerIndex.
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPirces(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
```

We can clearly observe that there are some dramatic improvements by removing the nested for loop from our VBA code.

### 2017

![2017 Before Refactoring Analysis.](/Resources/VBA_Challenge_2017_Before.png " Timing of processing the 2017 dataset Before refactoring.")

![2017 After Refactoring Analysis.](/Resources/VBA_Challenge_2017.png " Timing of processing the 2017 dataset after refactoring.")

Looking at our performance improvements for 2017 we can observe a ~5.7x performance improvement after refactoring our code

### 2018

![2018 Before Refactoring Analysis.](/Resources/VBA_Challenge_2018_Before.png " Timing of processing the 2018 dataset Before refactoring.")

![2018 After Refactoring Analysis.](/Resources/VBA_Challenge_2018.png " Timing of processing the 2018 dataset after refactoring .")

Looking at our performance improvements for 2018 we can observe a ~4.5x performance improvement after refactoring our code

## Summary

### What are the advantages or disadvantages of refactoring code?

- Advantages
  - Performance can improve if the right decisions are made during refactoring code
  - Unnecessary code can be removed allowing the code to run using fewer computing resources
  - Taking a new look after the main objective of the code has been achieved can result in a simpler solution hence easier to maintain

- Disadvantages
  - Additional time is required to refactor code which might not be available due to other priorities

### How do these pros and cons apply to refactoring the original VBA script?

In our analysis, we could improve the performance of our code several times which will allow us to scale the solution to perform a market analysis. 
We can clearly see and measure how by refactoring our code we were able to reuse some of the code we had but still achieve the objective in a simpler more efficient way. 

## Sources

### The data used for the Stocks analysis was provided by UT for the Data Analysis Bootcamp

### Markdown References
[Markdown reference file] (https://markdownlivepreview.com/)

:sunglasses: :space_invader: :robot:	
:see_no_evil: :hear_no_evil: :speak_no_evil:
