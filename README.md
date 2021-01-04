# Stock-Analysis

## Overview of Project

### Background.
Steve's parents wants to invest in a stock. They have chosen to invest in Daqo (ticker: DQ). But before investing they want to know how DQ's stocks performed in 2018. They have asked their son, Steve for help. 

Steve reached out for help with the analysis and soon discovered that DQ's stocks is not a good fit for his parents. So he decided to research on a broader stocks in the market. 

So assuming the vba code Steve used to run his analysis works works well for a dozen stocks. There is no guarantee it will work well for thousands of stocks. And if it does, it may take longer time to execute.

### Purpose
Rewrite the code Steve used for his stocks analysis to run faster, a process called refactoring in coding process, and give a written report of the analysis made.

## Results
The tickerIndex is set equal to zero before looping over the rows 
![1](https://user-images.githubusercontent.com/69058584/103587898-cc5eb680-4ead-11eb-9122-79bbb329875d.PNG)


Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
![2](https://user-images.githubusercontent.com/69058584/103588027-029c3600-4eae-11eb-88ca-fbd1d9f4a203.PNG) 

The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays 
![3](https://user-images.githubusercontent.com/69058584/103588070-1778c980-4eae-11eb-960d-3db9473f3ed1.PNG)

The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and       tickerEndingPrices
![4](https://user-images.githubusercontent.com/69058584/103588105-26f81280-4eae-11eb-9353-a621de7001af.PNG)



Code for formatting the cells in the spreadsheet is working
![5](https://user-images.githubusercontent.com/69058584/103588108-2790a900-4eae-11eb-94e0-4e0eb8b66980.PNG)


There are comments to explain the purpose of the code 
![6](https://user-images.githubusercontent.com/69058584/103588111-2790a900-4eae-11eb-9452-5401a66a1896.PNG)



The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module 
![7](https://user-images.githubusercontent.com/69058584/103588114-28293f80-4eae-11eb-87c5-b4718e1dfd05.PNG)
![8](https://user-images.githubusercontent.com/69058584/103588115-28293f80-4eae-11eb-82cb-59be06569ff3.PNG)




The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/69058584/103588119-28c1d600-4eae-11eb-8ee6-ac95634b6d62.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/69058584/103588121-28c1d600-4eae-11eb-9cc7-3fb2d5eff2f7.PNG)



## Summary
1. What are the advantages or disadvantages of refactoring code?
  
  ### Avantages of refactoring code are
  * Code runs faster and efficiently
  * commonly use on the job
  
  ### Disadvantages of refactoring include
  * Writting complex syntax or scripting in other to accomplish a wider scope of result
  * can be time demanding 
  
2. How do these pros and cons apply to refactoring the original VBA script? the pros of refactoring the original vba makes it efficient by requiring less memory for example. On the other hand, the cons to refactoring can be writing of a complex syntax to say. for example, instead of using a variable that hold a single value, array variables are utilized. And array requires indexing. 
  


