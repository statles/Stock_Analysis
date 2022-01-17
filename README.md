# Stock_Analysis
Analysis of 12 stock tickers to find total volume and daily return

## Overview of Project
The purpose of this project was to analyze twelve stock indexes to see which one(s) would be the best investment. A VBA code was written because the excel data was too large to analyze manually. This code would need to analyze all the stocks in the dataset and the user should be able to choose which year they wanted analyzed. In the end, this code was refractored to make the code run faster. The analysis will center on how effective refactoring the code was at speeding up the code, as well as any advantages or disadvantages that can come from refactoring the code.

## Results
The results from refactoring the code by using arrays rather than nested for loops showed that the code ran about 1 second faster than the non-refactored code. This may be based on factors such has the fact that the non-refactored code had to activate the stock worksheet and then activate the analysis worksheet for each ticker index. This most likely caused the code to slow down. The new, refactored code only activates each worksheet once, and stores the values in several arrays for later use.
![worksheetactivation](https://user-images.githubusercontent.com/95397823/149689976-629b0738-acb0-434b-b7c9-295999c6ee32.png)

For the stocks themselves, we can see that 2017 was a better year for our twelve stocks than 2018. In 2017, the majority of stocks saw a percent increase, except the stock index "TERP", which saw a decrease in daily return. For 2018, only "RUN" and "ENPH" saw a percent increase in daily return.


## Summary
In general, refactoring a VBA code can have the advantage of making the code run faster and more efficiently. Instead of writing the same portion of code many times, you can simplify the code. Moreover, you can address points in the code that cause a lot of stress on the machine. For example, if a code is activating a worksheet over and over again, it may be worth it to simplify the code.
The disadvantages of refactoring a code is that it may make it more cumbersome for the end user. It may also be more likely to cause syntax errors. Sometimes, it is more worth it to write a simpler code that is more user friendly and easier to read.

Refactoring the code had the advantage of allowing the code to run faster by about a second. The refactoring also had the added advantage of making the code more efficient by not having to reactivate the worksheet for each instance of the for loop. The for loops were not nested with the refactored code, as the variables were stored in arrays.
The disadvantage of refactoring the code was that it was much harder to write. I had trouble conceptualizing the arrays and indexing the stock ticker indexes. In addition, sometimes refactoring the code leads to a lot more work from the user, but the amount of time saved by running the code is not that significant.


