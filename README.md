# Stock-Analysis with Excel VBA
## Project Overview
### Background
Steve has recently graduated with his Finance degree and has asked for our help in analyzing the Green Energy stocks for the years 2017 and 2018 for his parents so they can make a decision on what stocks to invest in. We used Excel VBA to first write code and analyze the stocks that Steve's parents had decided to invest in (DAQO New Energy Corp) and then we analyzed all the stocks for 2017 and 2018 pertaining green energy. 
### Purpose
The purpose of this project is to make the solution code for all stock analyses more efficient by refactoring that code. In other words, we make the code cleaner that makes it run faster in the system. We also compare the execution times of the stock performance with original script and then the refactored script and see if the time elapsed.
## Results
### Analysis
In order to make this refactoring work, I created four different arrays, called 'tickers', 'tickerVolumes' , 'tickerStartingPrices', and 'tickerEndingPrices' and used the variable 'tickerIndex' to access the correct index across all these arrays. 'tickerIndex' was set to zero before I used it to iterate over all the rows and a for loop was initiated to loop through stock data which read and stored all the values from tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices rows. This script has the same functionality as the original one, the only difference is that it is more organized and efficient.  Below is the original code compared to the refactored one. 

#### Original Code
    
    'Initialize array of all tickers
    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'initialize variables for starting and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Worksheets(yearValue).Activate
    'find # of rows to loop over (2 to 3013)
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'loop through the tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        'loop through rows through data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            End If
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            End If
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            End If
            
        Next j
        
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    

#### Refactored code:

    '1a) Create a ticker Index
       Dim tickerIndex As Single
       tickerIndex = 0

    '1b) Create three output arrays
       Dim tickerVolumes(12) As Long
       Dim tickerStartingPrices(12) As Single
       Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
    ''2b) Loop over all the rows in the spreadsheet.
       For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

            
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
           tickerIndex = tickerIndex + 1
            
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    
#### Stock Performance 
    
Below are the images for the stock performance results generated using our codes: 
    
![2017-Results](https://user-images.githubusercontent.com/107225715/174879475-6344d601-3195-4150-b156-10669c94a22e.png)
![2018_Results](https://user-images.githubusercontent.com/107225715/174879483-25572e56-8cdb-401a-a1b1-25a992342cff.png)

#### Run-Time Comparisons

Here are the run times for the year 2017 and 2018 with the original code:

![VBA_Challenge_2017_Original](https://user-images.githubusercontent.com/107225715/174879754-30809b28-974b-45d0-905a-13535ed8bf18.png)
![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/107225715/174879808-1d35fd8e-b6b2-42aa-862c-f170c6430ba7.png)

Here are the run times for the same years using our refactored code: 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107225715/174879940-12e99356-1c85-4f17-92ed-1dc7a5f63a8f.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107225715/174879978-0507fec2-78a3-4c4a-ab61-8f505393bc9c.png)

## Summary

#### 1. What are the advantages or disadvantages of refactoring code?
Some advantages of refactoring codes are that it makes the code more efficient by using fewer steps, it's cleaner and easier to read, and it also leads to a            better quality code that runs faster which might not have happened in the first try while trying to accomplish the task. The disadvantage of refactoring is             that it can take a longer time than expected and also that it could be risky if the original code hasn't been saved if anything goes wrong with refactoring. 

#### 2. How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original VBA script definitely resulted in a decreased macro run time as shown in the 'Run-Time Comparisons' section. It also made our code more organized and readable for future readers. The problem I ran into while refactoring the original VBA script is that it took way longer than I expected considering I already had accomplished the same task with the original code. I had the original code backed up in several places so even though I was running into a lot of syntax errors and messed up my Excel workbooks a few time, I was able to salvage the code and use it again from my back ups. 

  


    
