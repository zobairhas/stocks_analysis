# Refactoring All Stocks Analysis

## Overview of Project
Analyze a collection of stocks and determine which tickers performed better in 2017 vs. 2018.
To achieve this I had to use VBA to run the analysis, measure the performance of the code, and provide an analysis of the results.
Afterwards, I refactored the VBA code so that it runs faster. 
 

## Results (analysis is well described with screenshots and code)

- Based on returns, it would appear that 2017 was a much better year for these tickers/companies versus 2018.
- Only two tickers/companies gave positive returns in 2018 (ENPH and RUN).
- Performance-wise, they're both similar. The 2017 script was slightly slower.

<br>

**2017 All Stocks Analysis:**

![2017 All Stocks Analysis](https://user-images.githubusercontent.com/44425379/148717889-efef3765-570f-4eb1-9f28-64df9f458928.png)

<br>

**2018 All Stocks Analysis:**

![2018 All Stocks Analysis](https://user-images.githubusercontent.com/44425379/148719013-849cac19-a6b4-4958-a4cd-e1297e257cf1.png)

## Performance Time for Original Script vs. Refactored Script
- After refactoring the original script, I was able to significantly improve the performance time. 
- Several reasons for this:

 **1)** The original script used a nested `for` loop that was constantly checking information against the spreadsheet. 

```
Old script for All Stocks Analysis

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

```

**2)** None of the variables were set up as arrays that could be indexed. In the below example, you can see the difference. 
```  
Old script for All Stocks Analysis
   '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Single
    Dim endingPrice As Single
```

```
Refactored script 
   'Output arrays
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickerVolumes(12) As Long
```	
**3)** By setting up arrays, we can index the information once instead of referencing the sheet over and over again. Ultimately, this leads to me using single `for` loops.

```   
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

```	
**4)** The above is not the full code, but once the script is run you get the following:

<br>

**2017 Refactored Script:**

![2017 Refactored Script](https://user-images.githubusercontent.com/44425379/148718186-b62ea47e-cc47-4d42-85ee-85d6904fb7a9.png)

<br>

**2018 Refactored Script:**

![2018 Refactored Script](https://user-images.githubusercontent.com/44425379/148718195-111b0e5f-ce16-4fcc-9282-b5a6b55bd782.png)



## Summary
- What are the advantages and disadvantages of refactoring code in general?
	- Advantages of refactoring code allow the person to optimize inefficient code, find code smells/reduce tech debt, and make the code overall easier to understand. 
	- It's also interesting to see different solutions regardless of whether they are optimal or not. 
	- Disadvantage of refactoring code is that it's time-consuming. Learning how someone else coded something, especially if variables don't intuitively make sense, could cause confusion. Leading to time being wasted or it could potentially introduce more bugs in the future. 

- What are the advantages and disadvantages of the original script and refactored VBA script?
	- Advantage of the original script is that it's easier to write out, and it's less time-consuming to create the overall solution even though it's not optimal for performance. Fewer items to maintain and keep track of.
	- Disadvantage of the original script is that it's slower, and it's not using best practices even though the solution works. 
	- Advantage of the refactored script includes better performance time.
	- Disadvantage is that refactoring code is slightly more time-consuming. Also, it is easier to mess up the code as there are more items to keep track of. 
	
	
	
