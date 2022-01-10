# Refactoring All Stocks Analysis

## Overview of Project
Analyze a collection of stocks and determine which tickers performed better in 2017 vs. 2018.
To achieve this I had to use VBA to run the analysis, measure the performance of the code, and provide an analysis of the results.
Afterwards, I refactored the VBA code so that it runs faster. 
 

## Results (analysis is well described with screenshots and code)

### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
-Based on returns, it would appear that 2017 was a much better year for these tickers/companies versus 2018. 
-Only two tickers/companies gave positive returns in 2018 (ENPH and RUN).
-Performance-wise, they're both similar. The 2017 script was slightly slower.

**2017:**
(https://github.com/zobisurf/stocks-analysis/blob/main/All_Stocks_Analysis_VBA_Original_2017.png)

**2018:**
(https://github.com/zobisurf/stocks-analysis/blob/main/All_Stocks_Analysis_VBA_Original_2018.png)

### Performance Time for Original Script vs. Refactored Script
-After refactoring the original script, I was able to signifnicaltly improve the performance time. 
Serrveral reasons for this:
	-the origal script used a nested `for` loop that was constantly checking information agaisnt the spreadsheet. 
	   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2017").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   -Additionally, none of the variables were set up as arrays that could be indexed. In the below example, you can see the difference. 
   
    '3a) Initialize variables for starting price and ending price
	Dim startingPrice As Single
    Dim endingPrice As Single

   'Output arrays
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickerVolumes(12) As Long
	
	-By setting up arrays, we can index the information once instead of referencing the sheet over and over again. Ultimately, this lead to me using single `for` loops to loop through the information.
	
	   'Output arrays
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickerVolumes(12) As Long
   
    'Create a for loop to initialize the tickerVolumes to zero
     For i = 0 To 11
        
        ticker = tickers(i)
        tickerVolumes(i) = 0
        
     Next i
    
    'Loop over all the rows in the spreadsheet
     For j = 2 To RowCount
     
      'Increase volume for current ticker
       If Cells(j, 1).Value = tickers(tickerIndex) Then
           
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        End If
        
        'Check if the current row is the first row with the selected tickerIndex
         If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
        End If
        
        'Check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex
         If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            tickerIndex = tickerIndex + 1
            
        End If

      Next j

    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return
     For k = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumes(k)
        Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
    
    Next k
	
	-This is not the full code of course, but once the script is run you get the following run times:
	
	<image challenge 2018>
	<image challenge 2017>

## Summary
- What are the advantages and disadvantages of refactoring code in general?
	- Advantages of refactoring code in general allows the person to optimize inefficient code, find code smells/reduce tech debt, and make the code overall easier to understand. 
	-It's also interesting to see different solutions regardless if they are optimal or not. 
	-Disadvantages of refactoring code in general is that it's time consuming. Having to learn how someone else coded something especially if comments are there or items aren't clearly defined it could cause confusing and potentially introduce more bugs.
	-Using personal work experience, refactoring is almost always done at the start of the sprint or well ahead of any hard deadlines because of the risks asscistedaited. 

- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script
	- Advantages of the original script is that it's easier to write out and it's less time consuming to create the overall solution even though it's not optimal for performance. Less items to maintain and keep track of.
	-Disadvantages of the original script is that it's slower and it's technically not using best practices even though the solution works. 
	-Advantages of the rafactored script includes better performance time, but it is harder to mess up as there are more items to keep track of. It's slightly more time consuming. 
	
	
