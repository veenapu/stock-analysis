# stock-analysis
# 	**VBA CHALLENGE - Written Analsyis and Results**

##	**Overview of Project** 
	The overview of the project is to expand the research and analyze the entire dataset to include the entire stock of 12 tickers 		instead on one stock.

##	**Purpose of this analysis**
	The purpose of this analysis is to refactor the code to loop through all the data one time in order to gather the same 	information and determine whether the refactoring the code made the VBA script run faster or slower in 2017 vs 2018. 

##	**Results and Analysis of the Challenge**
###	Results: 
	Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of 	the original script and the refactored script.
	
	Looking at the comparison of the stock performance, ENPH and RUN seem to have done well in both 2017 and 2018 compared to the 		rest of them. Especially, RUN which outperformed the rest in 2018 and ENPH seems to hold.  Other than these two, rest of the 		stocks took a dive, especially DQ which took the biggest hit. So, it was wise to expand the analysis before deciding where to 	invest.
	
	The biggest benefit of refactoring is the run time for running the code.  It was less for 2018 then 2017 after refactoring the 	code. See attached screenshots of the comparison of the stock performance and the run time. 

###	Analysis:
	
#### 1a) Created a ticker Index and initialized it to zero
    tickerIndex = 0
#### 1b) Created three output arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices with the appropriate data types
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

#### 2a) Created a for loop to initialize the tickerVolumes, tickerStartingPrices and tickerEndingPrices to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
#### 2b) Looped over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
#### 3a) Increased the volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
#### 3b) Checked if the current row is the first row with the selected tickerIndex by using an If Then statement
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
#### 3c) Check if the current row is the last row with the selected ticker and if the next row’s ticker did not match, increased          the tickerIndex by using an IF Then statement.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
#### 3d Increased the tickerIndex and also closed the loop.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
             tickerIndex = tickerIndex + 1
        End If
    Next i
                
#### 4) Looped through the arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
              	Worksheets("All Stocks Analysis").Activate
        	        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Next i

![2017](https://github.com/veenapu/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![2018](https://github.com/veenapu/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)     

## Summary: 
### In a summary statement, address the following questions.

### What are the advantages or disadvantages of refactoring code?
#### Advantages:
-	Improves the design of the software
-	Makes software easier to understand
-	Helps finding bugs
-	Hels program run faster
#### Disadvantages:
-	Can be risky if the code/application is too big
-	Can be risky if the existing code does not have a proper test case
-	Can be risky if the developers do not understand what it is all about

### How do these pros and cons apply to refactoring the original VBA script?
	Refactoring allows the code to be made much simpler, much more efficient and easier to read which means that it will be easier to 	debug and run faster.  This in turn can save time and money for the company.
	However, if the code is too big and too complex, refactoring might not be an option or if you do not have a way to test it.  
