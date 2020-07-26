# Stock Analysis with VBA

## Overview of Project

### Purpose and Background
The purpose of this project was to analyze stock performance through refactoring previously constructed code, which would analyze the total daily volume and return rate of 12 different stock tickers. Theses tickers were organized in alphabetical order in contiguous rows, and there was no excess junk data that required cleaning prior to the analysis. The starter code is able to successfully complete the task, however, it may prove inefficient for larger volumes of data due to its nested for loops. Ultimately, refactoring this code allows us to perform the task significantly faster, and handle larger volumes of data.


## Results
### Analysis of stock performance
Based on the analysis, we can observe a general positive trend in 2017 where all but 1 stock ticker had net positive returns. In contrast, the results from 2018 show all tickers having negative returns with the exception of 2 stocks.
![Screenshots](/Resources/2018_vs_2017_performance.PNG)

Furthermore, the visualizations reveal that there is a correlation between higher total volume contributing to net positive returns. In general, stock tickers that had high net positive returns and  tend to have higher daily volumes traded. However, there were some tickers including DQ, HASI, SEDG, and VSLR that had a negative return rate despite increasing their total daily volume from 2017 to 2018. This indicates that there were additional factors that affect the return rate outcome. 
![Screenshots](/Resources/2017_2018_visualizations.PNG)

In conclusion, stock tickers ENPH and RUN seem to be doing quite well, having a positive return rate despite all other companies having negative returns. It would be interesting to evaluate and see if there is a common theme resulting in these negative rates.


### Comparing time taken before and after refactoring
Initially, the complete analysis took about 0.65 seconds to complete.
![Screenshots](/Resources/unrefactored_time_results.PNG)

Upon refactoring the code, the running time was significantly reduced, requiring <0.01 second for each year. The key changes that increased its efficiency were utilizing arrays to store the ticker volumes, as well as starting and ending prices, allowing us to remove the nested for loops in the unrefactored code. 
![Screenshots](/Resources/refactored_time_results.PNG)

Unrefactored pseudocode with nested for loop:
'''
4) Loop through tickers
   For i = 0 to 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       For j = 2 to RowCount
           '5a) Get total volume for current ticker
           '5b) get starting price for current ticker
           '5c) get ending price for current ticker
       Next j
       '6) Output data for current ticker
   Next i
'''
Refactored code utilizing arrays and separating the nested loop:
'''
    For i = 2 To RowCount
        '7a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
        '7b) Check if the current row is the first row with the selected tickerIndex.
        '7c) check if the current row is the last row with the selected ticker
            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        ...
    Next i
'''
