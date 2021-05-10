# stocks-analysis
VBA of Wallstreet Project Folder
## Overview of Project
In this project, I analyze the performance of 12 individual stocks in 2017 and 2018 for my client, Steve, to provide him with potentially lucrative investment options. I used VBA to analyze a dataset for this basket of stocks, which consisted of daily values for the stock's trading volume as well as price information (high, low, open, close, and adjusted close). To help Steve identify the best stocks from this basket, I calculated the total annual trading volume and annual return for each stock. Additionally, I refactored the code to improve performance, which involved reducing the use of a double for-loop to a single one. In the following sections, I will outline my methodology for constructing the refactored code, discuss the results of my analysis, and provide a brief summary of the advantages and disadvantages of refactoring code. 
## Methodology
As previously mentioned, the dataset consisted of 12 unique stock tickers with daily data on the stock's trading volume, as well as high, low, open, close and adjusted close price. Below is an image of the raw data. 

![Stock Market Raw Data](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/Raw_Data.png)

My VBA code starts by having the user (i.e., Steve) click a Run button and then prompting him to input a year of data to analyze, either 2017 or 2018. It then starts keeping track of the time it takes to process the code. Generally, the code works by creating 4 separate arrays to keep track of the unique stock tickers, the trading volume for each stock, and the price of the stock at the start and end of the year. Below is some example code where I create the arrays and populate the "tickers" one with all the stock tickers from the dataset.

![Stock Market Arrays Creation](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_Array_Creation.png)

It then loops through each row of the dataset and uses a set of conditional statements to determine the start and end of data for each ticker. For each stock ticker, it will assign the first row as the starting price and the last row as the ending price, as well as keep track of a running total of the trading volume. The code then outputs to a worksheet the total trading volume for each stock and the annual return, which it calculates by finding the percentage change between the ending price and starting price. Finally, it applies some conditional formatting that highlights in green all the stocks with a positive return and red the ones with a negative return and prints out a message box that tells the user how long the code took to run. Below is example code which shows how the conditionals check for the first row and last row of a particular stock market ticker and how I populate the arrays based on these conditions. 

![Stock Market For Loop](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_For_Loop.png)

### Refactoring
Originally, the code operated using a double for-loop -- one loop to go through each unique stock ticker and one loop to go through each row in the dataset. However, I refactored the code to increase it's efficiency and allow for larger datasets to be run in a reasonable amount of time. To accomplish this task, I reduced the use of a double for-loop to a single for-loop, which you can see in the last image from the prior paragraph. This was accomplished by using a integer variable called "tickerIndex," which was created in the second image from the previous paragraph. I gave this variable a starting value of zero and then increased it's value by one every time the code found the last row of a particular stock ticker. Thus, this variable allowed me to reference the particular stock ticker the code was currently working through and populate all of the arrays with the correct values for that particular stock. You can see the code populating the arrays based on the current "tickerIndex" value in the last image in the previous paragraph.

## Results 
### Stock Performance Results
Below are the results of my stock analysis for both 2017 and 2018. 

![Stock Market 2017 Results](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_2017_Results.png)

![Stock Market 2018 Results](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_2018_Results.png)

Based on the above results, in appears that ENPH may be the most attractive stock ticker for Steve based on past returns and trading volume. Between 2017 and 2018, it's trading volume increased by 174%, suggesting that it has become a very liquid and popular trading option. Furthermore, it's average annual returns for 2017 and 2019 was 105.7% and was one of two stocks that actually had a positive return in 2018. Although past returns don't guarantee future results and the stock price may have become overvalued, my analysis suggests ENPH is the highest performing stock of the dataset, at least in the two years I analyzed, and deserves further consideration. While 11 out of the 12 stocks had a positive return in 2017, only 2 out of 12 had a positive return in 2018, suggesting that ENPH may be a better option during a bad year for the market as well. The other option would be ticker "RUN", which had modest gains in 2017 but better gains in 2018. 

### Code refactoring Results
Below I have included two images: the first shows how long the refactored code took for the 2017 data and the second shows how long it took for the 2018 data.

![Stock Market 2017 Code Duration](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![Stock Market 2018 Code Duration](https://github.com/SethBoswell/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)

You can see from the above images that the 2017 data only took approximately 0.13 seconds to run and the 2018 data only took 0.14 seconds. The original code took 0.85 seconds for 2017 and 0.85 seconds for 2018, so refactoring the code reduced the run time by about 83.5%. In my summary section, I will discuss the pros and cons of this refactoring.

## Summary 
Generally, I think that refactoring code's main advantage is it reduces the runtime of the code, allowing it to run more efficiently. Reducing the double for-loop to a single one significantly improved the runtime, which means Steve could run this data on larger datasets with more stock tickers without worrying about the code timing out. The main disadvantage I found is that the refactoring the code requires some creativity in order to think of alternative and more efficient ways to structure your code. Using double for-loops may seem like the easiest and most logical way to structure your code, but you can save a significant amount of runtime by reducing this to a single one, even if it requires more time to think about how you will structure the code.

I believe the above analysis holds true for this particular project. The refactored code runs much faster, implying that VBA will be handle a much larger dataset in the refactored code versus the original. However, it took me more time to deconstruct the double for-loop and think about using the variable "tickerIndex" to keep track of which stock ticker my code was currently analyzing in the single for-loop. Despite being faster, it may also be a little tricker for another coder to edit my code with the refactored script just because it is a little more difficult to understand compared to the double for-loop.


