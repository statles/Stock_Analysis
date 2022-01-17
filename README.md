# Stock_Analysis
Analysis of 12 stock tickers to find total volume and daily return

## Overview of Project
The purpose of this project was to analyze twelve stock indexes to see which one(s) would be the best investment. A VBA code was written because the excel data was too large to analyze manually. This code would need to analyze all the stocks in the dataset and the user should be able to choose which year they wanted analyzed. In the end, this code was refractored to make the code run faster. The analysis will center on how effective refactoring the code was at speeding up the code, as well as any advantages or disadvantages that can come from refactoring the code.

## Results
The results from refactoring the code by using arrays rather than nested for loops showed that the code ran about 1 second faster than the non-refactored code. This may be based on factors such has the fact that the non-refactored code had to activate the stock worksheet and then activate the analysis worksheet for each ticker index. This most likely caused the code to slow down. The new, refactored code only activates each worksheet once, and stores the values in several arrays for later use.
For the stocks themselves, we can see that 2017 was a better year for our twelve stocks than 2018. In 2017, the majority of stocks saw a percent increase, except the stock index "TERP", which saw a decrease in daily return. For 2018, only "RUN" and "ENPH" saw a percent increase in daily return.


## Summary
Refactoring the code had the advantage of allowing the code to run faster by about a second. The refactoring also had the added advantage of making the code more efficient by not having to reactivate the worksheet for each instance of the for loop. The for loops were not nested with the refactored code, as the variables were stored in arrays.
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

