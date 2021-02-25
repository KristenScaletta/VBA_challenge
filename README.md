# VBA_challenge
This code runs through a list of stocks and calculates the following per stock:
-Yearly change
-Percent change
-Total volume

The stocks are organized by year based on worksheet tab, and the code can be run on any worksheet to print the results in a new summary sheet.

In terms of order, in order to get the summary sheet, the sub routine newsummarysheet must be run first and only once. Then, to print results in the sheet, the sub routine stockdata should be run second. Finally, to visually indicae whether yearly change is positive or negative, the subroutine cell color should be run last. Stockdata and cellcolor should be run each time you're looking for summary data for a new year. The summary sheet only holds one year's summary so will overwrite previous years if run on a new one.
