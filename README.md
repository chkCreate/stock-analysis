# Mod 2 - An Analysis of Green Stocks with VBA

## Overview of Project

### Purpose
  
The client would like to analyze an entire dataset of financial stock data with VBA code that takes a shorter time to execute. In this challenge, we refactored the existing VBA script to run faster using the array function. The refactored VBA script ran around nine (9) times faster than the previous script.
  
## Results
  
### Analysis of Stock Performance Between 2017 and 2018
  
I reviewed the Total Daily Volume and Return data for years 2017 and 2018 and identified stocks that were both higher than average Total Daily Volume and had positive returns both years. The stock "Run" met both of the criteria and would be recommended per these standards only. However, comprehensive stock analysis entails more steps and analyses including valuation of companies may need to be performed before any fincial recommendations can be given.
  
### Execution Times Between Original and Refactored Codes
  
The original VBA script ran roughly around 0.5 seconds for both 2017 and 2018 stock performance data. See the image right below for an example of the execution time for the year 2017 with the original code.
  
![alt text](2017 Original)
  
Rather than storing values in cells, the values are stored in arrays '''totalVolumes, tickerStartingPrices, tickerEndingPrices'''. Arrays are faster to store, retrieve, and modify. 
  
After the code was refactored, the execution time for year 2017 analysis was reduced by around nine-fold. See the image below for an example of the execution time for the years 2017 and 2018 with the refactored code.
  
![alt text](2017 Refactored)
  
![alt text](2018 Refactored)
  
## Summary
  
### General Advantages & Disadvantages of Refactoring Code
  
General advantages of refactoring code is to make the existing code more efficient: shorten, take less memory or improve readability. On the other hand, refactoring code may not always be favored or performed because of time and resources. If the original code already serves its purpose, it may be more of a want than a need with added costs.
  
### Pros & Cons of Refactoring the Original Code
  
The main benefit of refactoring the original code for this dataset is the reduced execution time especially for larger datasets. A relative disadvantage of refactoring original code is that the coder may need to spend more time to learn how to make the code more efficient and apply. 
  
## References
  
I worked on this module challenge #2 with a teammate @swlim314.
  
