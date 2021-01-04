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
![test1] ("C:\Users\Sa2di\Desktop\stock-analysis\Resource2\1.PNG")


Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
![imaage2]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\2.PNG")  

The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays 
![image3]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\3.PNG")

The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and       tickerEndingPrices
![image4]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\4.PNG")



Code for formatting the cells in the spreadsheet is working
![image5]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\5.PNG")


There are comments to explain the purpose of the code 
![image6]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\6.PNG")



The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module 
![image7]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\7.PNG")
![image8]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\8.PNG")




The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png 
![vba_challenge_2017]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\VBA_Challenge_2017.PNG")
![vba_challenge_2018]("C:\Users\Sa2di\Desktop\Class Modules Folder\Resources\VBA_Challenge_2018.PNG")



## Summary
1. What are the advantages or disadvantages of refactoring code?
  
  ### Avantages of refactoring code are
  * Code runs faster and efficiently
  * commonly use on the job
  
  ### Disadvantages of refactoring include
  * Writting complex syntax or scripting in other to accomplish a wider scope of result
  * can be time demanding 
  
2. How do these pros and cons apply to refactoring the original VBA script? the pros of refactoring the original vba makes it efficient by requiring less memory for example. On the other hand, the cons to refactoring can be writing of a complex syntax to say. for example, instead of using a variable that hold a single value, array variables are utilized. And array requires indexing. 
  


