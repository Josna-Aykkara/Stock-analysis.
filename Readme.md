# Overview of Project

## Purpose
The main objective of the project is to refactor the VBA code to make it more efficient to retrieve information from the dataset based on stocks for the two years 2018 and 2017. The purpose was to help in the decision-making process whether to invest in these stocks based on the total volume and return for years 2018 and 2017. The code also enables to retrieve reports on larger set of data if required.

## Background
The background of the project is to run the report using the 12 different tickers and various information’s such as date of trading, opening & closing prices, volume of trade, etc. for the multiple stocks traded in 2018 and 2017. The data set is used to summarise the information more accurately such as total daily volume and return on each stock to facilitate the decision process at a glance.


## Results

### Analysis

For the analysis of this project, we used the stock trading data from 2017 and 2018.  I designed the code to analysis the information automatically in a systematic order. Initially defined the required variables, input box to filter the required information based on year, created tickers, used a Nested For loop to get the daily volume and return values for the stocks and formatted the results derived from the code for easy reference. The code was refactored to reduce the time taken to run the report. Here is the final code used for the project to reduce the processing time.
![This is an image](https://github.com/Josna-Aykkara/Stock-analysis./blob/main/Resources/Code%20used.JPG)

 

•Based on the 2017 data, we discovered that most of the stocks traded had a positive return and were considered reliable investments especially the stock for DQ as it had a return of 199.4%. But as the 2018 finding, the stocks traded mostly had a negative return except for ENPH and RUN stocks. The DQ stocks also had a negative return of -62.6% which alerts us that the stocks are volatile and can change quickly based on the various factors in 2018. You can view the differences for the stocks in 2017 and 2018 from the below summary of stock performance.

![This is an image](https://github.com/Josna-Aykkara/Stock-analysis./blob/main/Resources/Performance%20of%20stock%20in%202017.JPG)

![This is an image](https://github.com/Josna-Aykkara/Stock-analysis./blob/main/VBA_hallenge_2018(latest).jpg)

 
•The original code was refactored to make it more efficient and quicker. The main transformation was due to the creation of the arrays- tickerIndex and tickerVolumes which allowed to loop across the rows based on the If condition embedded in the loop. It enabled to increase the tickerVolumes based on the ticker that is stocks for the year. As a result of the refactored code, the run time has reduced and would enable to run efficiently with more dataset (more tickers). The processing time reduced from 0.2898437 seconds to 0.1445313 seconds. 
 
Advantages or disadvantages of refactoring code
Like all codes, refactored codes have its advantages and disadvantages depending on the nature and desired outcome of the code. The main advantage of refactored code is that facilitate to get a faster code run time. It allows the codes to be organized, comments can be added for codes to clearly specify the purpose of the code and easier debugging. It enables personnel’s other than the coder to view and understand the code easily.

One of the disadvantages is that we cannot refactor all the codes created. It can be due to the nature of the implementation, depending if it is on a larger scale the refactoring would be difficult. It can also happen when the coder doesn’t get the required tickers to test or arrays to get the required information from.

Pros and cons apply to refactoring the original VBA script
One of the main advantages of refactoring the original VBA script is the reduction in the run time for the code. The original run time for the script for 2018 was from 0.2898437 seconds which was decreased to 0.1445313 seconds. As a result, we saved nearly 50% of the script run time after factoring. The additional benefit is that the refactored code can be used for larger set of data with larger tickers and would run efficiently. Now, I don’t see any disadvantage for the refactored code.
The screen shots of the refactored run time are shown below for year 2017 and 2018.
 
