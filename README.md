# DQ Stock Analysis

##  Purpose:
Steve’s parents have just invested into green energy as Steve’s first clients. He has asked for assistance with analyzing the excel file he has created by automating some of the functions through VBA. The excel file contains a bundle of green energy stocks as well as the stock his parents invested into, “DQ”.

Steve’s parents wish to know how active was “DQ” traded as well as its performance within the years. Knowing how active the stock was traded they believe that it will accurately determine the value of the stock. To determine how often “DQ” was traded we have decided to find the sum of the daily volume for the year. To find the performance of “DQ” we calculated the yearly return by finding the percentage increase and decrease throughout the year. As we will be finding the daily volume and analyzing the yearly performance for “DQ” we also decided to include the 12 other stocks within the worksheet to see how each stock matched up against each other.

After determining the results of the yearly performance and how active the stocks were traded Steve was pleased with the code and the outcome results we were able to provide. As this code was for only a dozen stocks and two years of information, we are now looking to write a more complex code that can faster determine the same results but with a larger dataset.

## Tools:
Microsoft Excel, Microsoft VBA

## Results:

Prior to refactoring the final code, we first started with the “All Stocks Analysis” macro to read Steve’s worksheet that included the years of 2017 and 2018. our code by writing a basic outline of how we wanted the program to flow. We inputted comments to stay organized all prior to coding so we can get a better understanding of how our flow will be formulated. Once all comments we inputted we executed the following code.

Please see below for our outline of how stay organized by inputting comments with in the code as well as the following code.

* Formatted the Output sheet on a “All Stocks Analysis” worksheet in order to activate the worksheet.
* Initialized an array of twelve tickers with an assignment
* Prepared the analysis of the ticker’s by initializing variables for the starting price an ending price of each ticker
* Set a value to the total volume to zero so when we loop through each ticker the loop resets.
* Create an inner loop to find the total volume, starting price, and ending price for each ticker.
* Lastly, we output the data onto the new worksheet that provides the following results.

![AllStocksAnalysis1](https://user-images.githubusercontent.com/81484054/116826075-320f3680-ab60-11eb-8e23-7df57b5c7baa.png)

![AllStocksAnalysis2](https://user-images.githubusercontent.com/81484054/116826076-320f3680-ab60-11eb-866f-1c421c5b284e.png)

![Green_stocks_2017](https://user-images.githubusercontent.com/81484054/116826077-320f3680-ab60-11eb-96f4-112d447baa40.png)

![Green_stocks_2018](https://user-images.githubusercontent.com/81484054/116826078-32a7cd00-ab60-11eb-9545-8d9ccbb88fbc.png)

The following results from above shows our code script ran in a time of 5.86 seconds and .02 seconds when first ran to loop through all the worksheets to output the value of our total daily volume and return percentage. Based on our findings the “DQ” stock has outputted a total daily volume that is below the average sum as well as a negative yearly performance. With these results it is proven to say within 2017 and 2018 the “DQ” stock shrunk its yearly return investment by - 62.80%.	

Now that we have ran the code script and were presented with the following results from above as we wish to now rewrite our code in order to output a larger dataset in a quicker time, we have taken the following approach and steps as described below.

* Firstly, we set our tickerindex as zero
* Created three output arrayd for tickerVolumes, tickerEndingPrices, and tickerStartingPrices
* Created a loop for the tickervolumes and a loop that went over all the rows in the spreadsheet
* Within the loop we wrote a script that increase the tickervolumes.
* Then factored a “if then” statement for the tickerStartingPrices and tickerEndingPrices to assign variables
* Increased the tickerIndex
* Lastly after looping through, we outputted the results on the “All Stocks Analysis” worksheet

The following code below shows how the above steps were then executed.

![Refactored1](https://user-images.githubusercontent.com/81484054/116826114-6e429700-ab60-11eb-9a97-34446cddf1e1.png)

![Refactored2](https://user-images.githubusercontent.com/81484054/116826116-6edb2d80-ab60-11eb-9145-4c4ae23019ca.png)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/81484054/116826117-6edb2d80-ab60-11eb-888c-5fb939767df3.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/81484054/116826118-6edb2d80-ab60-11eb-893a-5f79d939c9ed.png)

Based on the following results we were able to determine that the refactor code was able to output the refactored code in a faster manner then the original script.

## Summary:

### Advantages and Disadvantages

Being able to refactoring code does allow Steve to output a larger dataset that extends further than a dataset that contains two years and only twelve stocks. His updated scripted code will also have a faster run time to show the output results when visually displaying within the spreadsheet. One disadvantage of refactoring the code could be misdirecting the code. When rewriting it is highly possible of making mistakes which then can change all results.

### Pros and Cons

Our refactored code coincides with the original code as we have now created a code that can adjust to any additional amount of data that Steve will use moving forward when determining the total daily volume and return percentage when he wishes to analyze additional stocks and additional time lines. By refactoring the code, the original codes objective will only maximize its output limited to the original spreadsheet we originally received from Steve.
