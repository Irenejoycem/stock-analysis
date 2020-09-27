# stock-analysis
## Overview of Project: 
> The purpose of this analysis was to provide a tool for a client that wanted to compare the Daily Volume of stocks traded and calculate the Return for multiple stocks. I compared a list of 12 stocks from 2017 and 2018. 

## Results: 

Based on the results given, it seems that ENPH and RUN would be two stocks that I would suggest to invest in as the Annual Return is positive from both 2017 and 2018. I was able to analyze this from the total data by creating macros that were able to search through data from each respective worksheet. I created for loops and applied conditionals to extract information from the data given to each respective year of the 12 stocks. For the original VBA script code used for this, the code was able to calculate the Returns and Daily Volume but the run time was significantly longer than the refactored code. 

 ### Here were the results I analyzed in 2017:

> ![2017 Stocks Results](2017Analysis.PNG)

 ### Here were the results I analyzed in 2018:

> ![2018 Stocks Results](2018Analysis.PNG)

### Below are the run times of the refactored code which is significantly faster than the run time from the original code created for the analysis.

 ![2017 Run Time of Refactored Code](VBA_Challenge_2017.png)
 ![2018 Run Time of Refactored Code](VBA_Challenge_2018.png)

## Summary: 

> Some advantages of refactoring code is to be able to analyze large amounts of data. However, the disadvantage that I noticed when refactoring the code is that sometimes you would have to debug any compile errors that kept arising as I was trying to run the code. Debugging code is difficult when you know the original code was working and running well. The pros and cons refactoring the VBA script for this analysis is that when I was trying to assign variables, to the refactored code, I would often run into compile errors or even at times missing closing loops or conditionals. I would like to learn how to avoid making any of these mistakes, but it seems that this is something that occurs often when refactoring code. 
