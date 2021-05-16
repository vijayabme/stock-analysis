# stock-analysis
## Overview of The Project
### Purpose
Steve's parents are interested in investing their savings in green energy stocks. Without doing much research they decided to invest all their money in DAQO New Energy Corp which makes silicon wafers for solar panels and asked Steve for help. Steve would like to help his parents diversisy their investment. The idea is to help Steve analyze some of the green energy stocks and help his parents make a well informed decision. In this I am using VBA script to automate the data analysis in excel. 

## Results
### Analaysis outcome based on 2017 and 2018 stock market
Based on the DAQO analysis, the stock dropped 63% in 2018 and from 2017 it dropped by 262%. Steve would like to provide better stock options for his parents. So we analyze all the green stocks that Steve listed for both 2017 and 2018. In 2017 lot of stocks had positive returns. But in 2018 only ENPH and RUN had positive returns with RUN actually improving its return by 78% from 2017. TERP had negative return in both years but it increased by 2% from 2017 to 2018 and the total daily volume also increased from 2017 to 2018. So investing in ENPH, RUN and TERP might be a better option for Steve's parents instead of DAQO.
![image]()
#Add figure link
### Refactoring the code

###### 1a) Create a ticker Index
We are using an index in this code which increments every time the tickers change. So it starts at 0 and increments to 11 which is the total number of tickers in our excel. 

Dim tickerIndex As Integer
    
    tickerIndex = 0

###### 1b) Create three output arrays

Created three output arrays with the size of 12. the array size is initialized to 12 as there will be 12 outputs for the 12 tickers. 

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
###### 2a) Create a for loop to initialize the tickerVolumes to zero.
The code belowe assigns zero to each of the 12 elements in the tickerVolumes array
   For j = 0 To 11
       
        tickerVolumes(j) = 0
        
    Next j
###### 3a) Increase volume for current ticker

The tickervolumes will be incremented by adding the volumnes in the corresponding rows for that tickerIndex 

 tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
 
 ###### 3b) Check if the current row is the first row with the selected tickerIndex.
 
 Check to see if the previous row ticker is not equal to the current ticker and the current row is equal to the current ticker then assign that closing price as the tickerstartingPrices for that tickerIndex.
 
 If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
 End If
 
 ###### 3c) check if the current row is the last row with the selected ticker
 check if the current row is the last row for that tickerIndex and assign the corresponding closing cost to the tickerEndingPrices and increment the tickerIndex.
 
 If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      tickerEndingPrices(tickerIndex) = Cells(i, 6).Value 
      '3d Increase the tickerIndex.
       tickerIndex = tickerIndex + 1
  End If
  
  ###### 4)Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
  
  After we get the outputs for all tickers we print it out to the "All Stocks Analysis Refactored" sheet 
  
  For i = 0 To 11
            
        Worksheets("All Stocks Analysis Refactored").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i

## Summary
### Advantages or disadvantages of refactoring code

Refactoring the code helps in cleaning up the code and making it more readable. 
It increases the processing speed of the code by removing the unnecessary loops and conditions.
Its easy to maintain and debug a refactored code and prevents any coding discrepancies.

### How do these Pros and cons apply to refactoring the original VBA script?

1) By using a tickerIndex and arrays we are getting rid of the outer for loop for the length of the tickers. By doing this we are saving time that it takes to loop through each   ticker.
2) We are using the for loop to only initialize the elements in the tickerVolumes array to zero. We got rid of the if and else condition for the tickerVolumes by using index as it automatically updates the tickerVolumes for each index.
3) we are printing the outputs to excel sheet in a for loop outside, instead of doing it after each ticker output calculation which in turn will help in increasing the processing speed.
4) when we compare the processing speeds before refactoring and after refactoring the processing speed increased.

![image]()
![image]()
#Add images




