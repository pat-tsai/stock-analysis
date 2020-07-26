# Stock Analysis with VBA

## Overview of Project

### Purpose and Background
The purpose of this project was to analyze stock performance through refactoring previously constructed code, which would analyze the total daily volume and return rate of 12 different stock tickers. Theses tickers were organized in alphabetical order in contiguous rows, and there was no excess junk data that required cleaning prior to the analysis. The starter code is able to successfully complete the task, however, it may prove inefficient for larger volumes of data due to its nested for loops. Ultimately, refactoring this code allows us to perform the task significantly faster, and handle larger volumes of data.


## Results
### Analysis of stock performance


### Comparing time taken before and after refactoring
Initially, the complete analysis took about 0.65 seconds to complete.
![Screenshots](/Resources/unrefactored_time_results.PNG)

Upon refactoring the code, the process took about 0.5 seconds less than before.
![Screenshots](/Resources/refactored_time_results.PNG)
