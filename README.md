# Stock Analysis with Excel VBA

## Overview

### Purpose

A recent finance college graduate named Steve wanted me, an excel power user, to help him with some financial analysis. I was able to create an Excel VBA script which automatically returned key trading stock metrics from a large data sets in excel spreadsheets. This information would eventually help inform Steve's parents to choose what environmentally friendly "green stocks" they should invest in. The purpose of this analysis was to learn VBA within Excel and to apply the principles of excel VBA macros. We used general coding logic such as for loops and conditional statments to create arrays. We then used iteration variables associated with arrays to store values pertaining to stock information such as stock tickers, annual trading volume, and annual return, in order to generate an automated report.



## Results


To refactor the code, an iteration variable called tickerIndex was created and was set to zero. This iteration variable was used to access the stock ticker index for the tickers array (all stock ticker names), the tickerValues array, the tickerStartingPrices array, and the tickerEndingPrices array.

`tickerIndex = 0`

This iteration variable was not set as the parameters for a for loop and it was unique because its value was manually adjusted by the script each time a new ticker was detected. Later, we will discuss how this variable was used to store and access data from our arrays. After receiving neccessary user input and declaring all required variables and arrays, we initialized all the values in the tickerVolumes array to zero.
 
`For j = 0 To 11`

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`tickerVolumes(j) = 0`

`Next j`

Then, we used with a for loop to read every row of the spreadsheet. In each row of the spreadsheet, the script would use to code to ensure it reached the first row containing the index position zero string for "ticker".  Then, when it detected the current ticker, the value in the daily trading volume column would be added to the tickerVolumes array for the corresponding ticker.

`For j = 2 To RowCount`

`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value`

The process would be repeated for the same ticker, looking into new rows each time, until the last row of the ticker was detected.

`If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then`
    
&nbsp;&nbsp;&nbsp;&nbsp;`tickerEndingPrices(tickerIndex) = Cells(j, 6).Value`

By this time, the script would have already added the final value of daily trading volume to the current index ticker. Then, the script would start looking in the next row for the next index position ticker. It was at this point in the script when the tickerIndex value was increased by 1, right before carrying out the loop again, in order to continue adding new values to the next index positions of our arrays.

&nbsp;&nbsp;&nbsp;&nbsp; `tickerIndex = tickerIndex + 1`

`End If`
 
 `Next j`

This process was repeated for all of the tickers. We were able to add individual daily trading volumes to each stock’s tickerVolume value because we created the tickerVolumes array. The values in the array could be accessed and destination cells could have their values updated using this line of code
  
  `Cells(4 + j, 2).Value = tickerVolumes(j)`
  
We could update each ticker’s starting prices and end prices as well with our tickerStartingPrices(j) and tickerEndingPrices(j) arrays. Each array had twelve different index positions, one corresponding to each index position of the tickers array (0 to 11). We could access each value in the array using the tickerIndex variable. In the non refactored version of the code, we looked through every row of the spreadsheet and added trading volume values to the ‘totalVolume’ value for each ticker. This totalVolume value was transient because it was exported to its destination sheet during every iteration for each ticker and then the same variable (totalVolume) was set to zero for the next ticker iteration. In the refactored version of our code, we created an array which stored all twelve of the tickerVolumes values. We could then access and export each tickerVolumes value by performing an independent for loop. This would iterate through the arrays we created to set the destination cells values to the values stored in each array’s corresponding index position. 



### A More Efficient VBA Script

After running a built in timer in our refactored Excel VBA script, we found the initial script run time dropped from 0.305 seconds to 0.0625 seconds.

![VBA_Challenge_2017.png](https://github.com/willmino/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/willmino/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary

### Refactoring Code


In general, refactoring code is an advantageous process. An unoptimized script could take a very long time to run if it is not optimized. The first way a code is written is generally the one of the easiest solutions to a coding problem. The relief a programmer may experience from solving a problem may cause them to overlook how the code could be optimized and run better. Often, a code may take unnecessary steps to repeatedly scan through all lines of data. This can lead to an efficient use of time for a business. The implied method of refactoring a code is to cut out unnecessary steps without compromising the precision and efficiency of of the code. 

A general disadvantage to refactoring code is that sometimes not everyone working with code will know exactly how the code works, even though it is a useful tool for a company. If you begin to refactor code without knowing how it works, it could be a huge waste of time for a company.

### Advantage and Disadvantages of this Excel VBA Script

- The first script for running the stock-analysis code used two iteration variables for loops. Variable ‘i’ in tickers(i) was used in a for loop to scan the data for the appearance of each index position ticker string at a time (i = 0 To 11). Then, within every iteration of the ticker loop, the variable ‘j’ was used in a nested for loop to iterate through every row of the data and perform a series of conditional statements (j = 2 to RowCount). This meant that the code was going to check every single row of the excel spreadsheet for the appearance of each ticker. This was inefficient because we did not need to check every row of the spreadsheet for every single ticker and it took too much time. In fact, it would be better to write a code that could check for the presence of each ticker one row at a time and then force the script to look for the next index position ticker once the final row of the current ticker was detected. In this way, the script would not loop through the entire excel spreadsheet for an unnecessary number of times. It would actually loop through all the rows only one time. Thankfully, we were able to accomplish this because each ticker was listed alphabetically and we designed our ticker array to be alphabetical with corresponding index positions.

- The inefficiency of the first version of our script gave rise to the need for a higher efficiency script through refactoring. Our refactored Excel VBA script was more efficient because it ran faster. We were able to cut out the unnecessary and repetitive steps from the first version of the script. This format for refactoring scripts would serve as a major advantage for tech companies who want to analyze large data sets in a highly efficient manner. 
