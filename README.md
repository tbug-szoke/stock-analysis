# VBA of Wall Street
## Overview of Project

### Purpose
In the VBA of Wall Street project, we sought to provide Steve with an easy method to analyze stock data and determine Total Daily Volume and annual Returns. To achieve this goal we wrote VBA code to enable Steve to complete the desired analysis at the click of a button.  We then optimized the performance of the VBA code by refactoring it. The refactored code should enable Steve to expand his analysis from his small handful of stocks to a much larger group of stocks with confidence in the performance of the program.
### Background
When we initially started working on this project, the first goal was to analyze the performance of a single stock, DQ, in a single year, 2018. Code to achieve that goal was written and executed.  As a follow up, it was desired to analyze data for any additional 11 stocks for the same year.  The code for analyzing DQ was utilized as a starting point to create a new subroutine analyze all 12 stocks for which Steve had collected data. As a final follow up, it was desired to modify the code so that it could be run for either year, 2017 or 2018, for which Steve had collected data. The code was revised again.  At this point, our written code had grown organically in response to requests for additional functionality; it was not designed or optimized for the final requested functionality and was likely not scalable to run for many more stocks.

## Results
### Stock Performance
Steve wanted to analyze two measures of stock performance: Total Daily Volume and Return.  

To determine the Total Daily Volume, we simply needed to accumulate the individual daily volumes for each row of stock data for the same ticker symbol.  This was done by:
1. looping through all rows of data (`For j = rowStart To rowEnd`), 
2. using an `if` statement to determine if the ticker symbol for the current row of data is the same as the ticker symbol on the previous row of data and if so, accumulating the current row's daily volume with the previously accumulated daily volume total:
   ```
   If Cells(j, 1).Value = tickers(i) Then
          
     totalVolume = totalVolume + Cells(j, 8).Value
   ```
To determine the Return, it was necessary to first capture the starting price for the given ticker symbol, then capture the ending price for the same ticker symbol, then finally calculate the return and write it to the output report.  This was completed by:
1. Identifying the starting price with an `if` statement to determine if this was the first row of data for the ticker symbol, and if so store the starting price:
   ```
   If Cells(j, 1).Value = tickers(i) And Cells(j - 1, 1).Value <> tickers(i) Then
           
     startingPrice = Cells(j, 6).Value
   ```
2. Identifying the ending price with an `if` statement to determine if this was teh last row of data for the ticker symbol, and if so store the ending price:
   ```
   If Cells(j, 1).Value = tickers(i) And Cells(j + 1, 1).Value <> tickers(i) Then
                
    endingPrice = Cells(j, 6).Value
   ```
3. Writing the calculated return to the output page:
   ```
   Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
   ```
   
After running the stock analysis for both 2018 and 2017, it was visually apparent that 2017 had been a better year for the selected stocks than 2018 had.  In 2017, the majority of stocks had a positive return, as clearly evidenced by the conditional formatting showing those returns highlighted in green:

**INSERT IMAGE**

While in 2018, the majority of stocks had a negative return, as indicated by all but 2 stocks being highlighted in red:

**INSERT IMAGE**

Total Daily Volumes were similar across both years, even though the returns were quite different.

### Code Performance
The code snippets above were taken from the original code in the Subroute 'AllStocksAnalysis' that was developed to answer Steve's initial and follow up questions. In this code, we utilized a nested loop.  In the first, outer loop, we looped through each of the 12 ticker symbols:
```
'Loop through each ticker
For i = 0 To 11
```
In the second inner loop, we looped through each row of data to analyze the ticker selected in the outer loop:
```
  'Loop through all rows to find ticker(i) data
  For j = rowStart To rowEnd
  ```
By designing the code this way, for each stock we want to analyze we have to again loop through all rows of data.  

When executing this code for 2017, after a noticable wait, results were displayed along with the code performance:

**INSERT IMAGE**

When executing for 2018, after a similar noticeable wait, results were updated along with the 2018 code performance:

**INSERT IMAGE**

In the refactored code, we have eliminated the nested loop and have instead utilized new arrays to store the 12 total daily volumes, starting prices and ending prices as follows:
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
By utilizing these arrays, we only need to loop through the rows of data one time, and we can calculate and store the results for all 12 stocks.  We increase the index of the array as we pass through the rows on data when we note that the current row of data is for a new ticker symbol:
```
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
  tickerIndex = tickerIndex + 1
```

After refactoring, the code runs much faster.  For 2017:
/stock-analysis/resources/VBA_Challenge_2017.png

Similar performance was achieved for 2018:
/stock-analysis/resources/VBA_Challenge_2018.png

## Summary

### Advantages and Disadvantages of Refactoring Code
Refactoring code, editing it not for the purpose of adding functionality but rather for the purpose of performance/efficiency, clarity or security, can have both advantages and disadvantages.
The potential advantages, including the following, are numerous. 
* Improved speed
* Reduced memory consumption
* Improved supportability
* Improved security

Given all of the potential advantages, it seems like one should refactor code any time that one of the above outcomes seems feasible to achieve.  However, there are also disadvantages to refactoring code, including most notably the potential to introduce new bugs. If the code in question is part of a larger code package, then the benefits of refactoring the code may not be worth the potential harm of introducing a bug that could have downstream unintended consequences.

### Advantages and Disadvantages of the Refactored VBA script
In this case, the original VBA script was already showing performance issues, and there were no other downstream impacts of refactoring the code, outside of the time to develop it and to address any bugs. As mentioned in the introduction, the original code was developed in an ad hoc manner - first written to answer one question, then modified to answer a similar question, then modified again. By standing back to reconsider 'what is the code doing' and 'how can that be done more efficiently' we can change our mindset and approach to providing a solution, and can, in fact, find one that is much better than the original.  By writing the new solution in a new module, I was able to keep the original code for reference.  This proved helpful as my first attempt at refactoring did contain an error that I needed to debug. In fairly short order, I had the new code working and was able to see a 99.8% improvement in the time to complete the analysis!
