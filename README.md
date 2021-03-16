# Stock-Analysis and VBA_Challenge

## Project Overview

The purpose of this project is quickly construct a series of macros using VBA to analyze stock performance and provide insight to the client to assist his parents in assessing the best stocks to purchase and examine his parents preferred stock DQ.  VBA codes and macros were used to evaluate multiple stocks and look at yearly returns for 2017 and 2018, from start to finish.  An output sheet with formatting and quick buttons to run and clear the macros were provided.

Code was then refactored using VBA script from the module to make it run more efficiently.  This will allow it to save time and perform faster than the simple initial VBA code. Performance improvement were captured with a timer.  Additionally output tables were compared from the original module and the refactored code to make sure the value outputs matched.

### Results


Fig 1 ![2017 Stock Performance](Resources/AllStocks_2017.png)

Fig 2 ![2018 Stock_Performance](Resources/AllStocks_2018.png)

Fig 3 ![Original 2017 Performance](Resources/Module_2_AllStockAnalysis_2017.png)

Fig 4 ![Refactored 2017 Performance](Resources/Vba_Challenge_2017.png)


Fig 5 ![Original 2018 Performance](Resources/Module_2_AllStockAnalysis_2018.png)

Fig 6 ![Refactored 2018 Performance](Resources/Vba_Challenge_2018.png)


### Summary

#### Question 1

Advantages to refactoring code are that it runs more efficiently and saves computing resources and outputs are delivered faster.  Although the improvements in this project are slight as seen in the results, if the datasets grow larger a code that runs 4 -5times faster becomes much more efficient and useful.
 
Disadvantages are two-fold.  The time to generate the refactored code might not ultimately be worth the effort.  If it take two hours to refactor the code and only saves half a second, the number of times you plan to run the needs to factor in to whether the refactoring time saves enough overall time.  Also refactoring could lead to errors if not don't correctly.  Your output may change if you are not careful.

#### Question 2
In this example, the time saved was negligible in absolute terms but the magnitude of the reudction was significant. See table below:

|Year|Original Code Time|Refactored Code Time|
|----|-----|-----|
|2017|0.976 Seconds|0.219 Seconds|
|2018|0.960 Seconds|0.188 Seconds|

The refactoring time in this case likely was not necessarily worth the time.  The original question was to review the performance of the 15 stocks and look at the other stocks compared to the DQ stock in question.  The simple original code ran that in under a second.  The refactoring time took longer than the three quarters of a second savings.  However if we were to expand the years and number of stocks in the dataset it may make up for the time spent.
