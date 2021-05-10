# stocks-analysis
VBA of Wallstreet Project Folder
## Overview of Project
In this project, I analyze the performance of 12 indivdual stocks in 2017 and 2018 for my client, Steve, to provide him with potentially lucrative investment options. I used VBA to analyze a dataset for this basket of stocks, which consisted of daily values for the stock's trading volume as well as price information (high, low, open, close, and adjusted close). To help Steve identify the best stocks from this basket, I calculated the total annual trading volume and annual return for each stock. Additionally, I refactored the code to improve performance, which involved reducing the use of a double for-loop to a single one. In the following sections, I will outline my methodology for constructing the refactored code, discuss the results of my analysis, and provide a brief summary of the advantages and disadvantages of refactoring code. 
## Methodology
As previously mentioned, the dataset consisted of 12 unique stock tickers with daily data on the stock's trading volume, as well as high, low, open, close and adjusted close price. Below is an image of the raw data. 

![Stock Market Raw Data](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/Raw_Data.png)

My VBA code starts by having the user (i.e., Steve) click a Run button and then prompting him to input a year of data to analyze, either 2017 or 2018. It then starts keeping track of the time it takes to process the code. Generally, the code works by creating 4 separate arrays to keep track of the unique stock tickers, the trading volume for each stock, and the price of the stock at the start and end of the year. Below is some example code where I create the arrays and populate the "tickers" one with all the stock tickers from the dataset.

![Stock Market Arrays Creation](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_Array_Creation.png)

It then loops through each row of the dataset and uses a set of conditional statements to determine the start and end of data for each ticker. For each stock ticker, it will assign the first row as the starting price and the last row as the ending price, as well as keep track of a running total of the trading volume. The code then outputs to a worksheet the total trading volume for each stock and the annual return, which it calculates by finding the percentage change between the ending price and starting price. Finally, it applies some conditional formatting that highlights in green all of the stocks with a positive retun and red the ones with a negative return and prints out a messagebox that tells the user how long the code took to run. 
### Refactoring
Originally, the code operated using a double for-loop -- one loop to go through each unique stock ticker and one loop to go through each row in the dataset. However, I refactored the code to increase it's efficiency and allow for larger datasets to be run in a reasonable amount of time. To accomplish this task, I reduced the use of a double for-loop to a single for-loop. This was accomplished by using a 
