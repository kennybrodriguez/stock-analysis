# stock-analysis

## Overview of Project
This was a project designed to help with a basic stock analysis on green energy stocks. Clients had previously invested in DAQO (DQ) and this was put together to help provide data on alternative green energy stocks to gauge performance in 2017-2018. The main performance indicators chosed were 'Yearly Volume' and 'Yearly Return'. Refactored previously run code to help it run faster. With a VBA script we were able to loop through the stock data provided to help calculate the yearly volume by summing up the daily volume traded. Also, took down the prices of stock at the beginning and end of year to calculate yearly return for the year. VBA script takes input (year) to display stock information for that year. 

## Results
### DQ analysis
Overall, the 2 years of data help to show that the clients investment in DQ was overall successful in terms of investment growth but there were better companies who were able to sustain a positive yearly return. DQ had a strong year in 2017 with a yearly return of 199.4% and lost some of those gains in 2018 with a yearly return of -62.6%. Overall return for the 2 years was 136.8%. 
![2017](/Resources/VBA_Challenge_2017.PNG)

The best option over the years of 2017 and 2018 was 'ENPH'. In 2017 they were able to obtain a return of 129.5% and followed that up in 2018 with a return of 81.9%. Overall the return for those 2 years was 211.4%. Second best option was SEDG with a overall return over the 2 years of 176.7%
![2017](/Resources/VBA_Challenge_2018.PNG)

In the original code, a list of tickers was established and then iterated over. This iteration went through all the rows in the data to check if the ticker in that row matches the ticker in the list. It ended up going over the same data everytime it checked for a ticker. This slowed down the script substantially. 

A better solution was to iterate over the rows once. This was done by taking advantage of that fact that the data was alphabetized by stock ticker and also sorted by date. The refactored code checked to see if the current row was the current ticker. If it was then it added the total daily volume. If it was the first row of that ticker then it also took down the price for the starting price. Lastly, if it was the last row it saved the price for the yearly return and increased the (tickerindex) to start the next stock ticker. Having to only go through the data once helped this run much faster.

## Summary

Overall, the advantages of the refactored code are that it is a lot more efficient and concise. It loops through the data once instead of once per ticker. Also, this code was designed in a way so that it can be run with other data, if it is presented in the same format. Also the use of color does help visualize the data a lot better.

Disadvantage to this would have to be that the data is usually not this clean. Script only runs if data is alphabetized and sorted by date. Lastly, since this is excel this would not work well with a ton of data. 
