# Stocks Analysis

## Overview of Project

### Purpose
Steve is researching investments in alternative energy production. There are many forms of green energy to invest in which include hydroelectricity, wind energy, geothermal energy, and bioenergy. Steve is seeking assistance so that he can help diversify funds and analyze stocks. We examined and interpreted 12 different stock data to see if there are specific factors that could be beneficial to Steve in his research. In our analysis, we designed an interactive workbook using Visual Basic for Application (VBA) within Excel to provide the Total Daily Volume and Return on Investment (ROI) of each stock. 

Now Steve wants to expand the dataset to include the entire stock market over the last few years, but it may increase the run time of the VBA script. Although our code works well for a dozen stocks, it may not execute as well for thousands of stocks. We will refactor the previous code to determine whether refactoring our code successfully will make the VBA script run more efficiently and faster. In this analysis, we will be comparing the new execution time with the original VBA code. We will use these insights to help best support Steve in his research investments in alternative energy production.

## Results

### Stock Performance 2017 - 2018

There is a major difference in the alternative energy production stocks between the years 2017 and 2018. In 2017, almost all stocks had positive returns except for TerraForm Power Inc (TERP) while in 2018, almost all stocks had negative returns except for Enphase Energy Inc (ENPH) and Sunrun Inc (RUN). Additionally, much of the Total Daily Volumes of the stocks declined between 2017 and 2018. 

![2017 Return](https://user-images.githubusercontent.com/29410712/180332648-fb5a02c4-8973-4927-a720-6487380f595b.png)

![2018 Return](https://user-images.githubusercontent.com/29410712/180332662-e49a5a0a-0c8a-428e-98c0-223e68755c5c.png)

In our analysis, we only compared two different variables. There are many other factors that influence the stock market that was not examined. Economic impacts, government interventions, and other unforeseen events can impact the stock market. We recommend researching other variables before making informed decisions.

### Refactoring Code

In our code, we created three output arrays: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. We also created a `tickerIndex` variable to access the index across the different arrays. Additionally, we created `for` loops to loop over all the rows in the spreadsheet and conditional statements to calculate the `tickerStartingPrices` and `tickerEndingPrices` to complete the analysis.

#### Refactored VBA Code

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
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
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
               tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = 
  
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

#### Original VBA Code

```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
    
   '1) Format the output sheet on All Stocks Analysis worksheet
   
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
   '2) Initialize array of all tickers
    Dim tickers(11) As String

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
    
    '3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

  '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
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
   
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

        endTime = Timer
        
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
```

### Execution Times

In our analysis, we tested the execution times of the original VBA script and compared it with the execution times of the refactored VBA script. 

#### Original VBA Script

![Screenshot (287)](https://user-images.githubusercontent.com/29410712/180338624-77803c3e-6787-4cd3-aece-1c771aa591df.png)

![Screenshot (289)](https://user-images.githubusercontent.com/29410712/180338898-a6549c3e-0324-4aaa-b985-37de3ccae6b1.png)

#### Refactored VBA Script

![Screenshot (283)](https://user-images.githubusercontent.com/29410712/180339288-0976d460-3c98-4ec8-876a-3c6d91b202a5.png)

![Screenshot (284)](https://user-images.githubusercontent.com/29410712/180339280-c8b9d209-6087-4b1d-9a4b-75ca9bcf7e3e.png)

As we can see, the refactored code runs more quickly and efficiently than the original code. The run times of the original code for 2017 and 2018 are 1.382 seconds and 1.492 seconds respectively. Furthermore, the run times of the refactored code for 2017 and 2018 are 0.121 seconds and 0.117 seconds respectively. In our analysis, we can conclude that the refactored code is about 91.7% faster than the original code. It is now more efficient and can help run an analysis for thousands of stocks to help best support Steve in his research investments of alternative energy production.

## Summary

### Advantages & Disadvantages of Refactoring Code
As a result, there are many advantages of refactoring code. Some advantages include removing duplicate code which improves the effectiveness of the code, increases the speed of running the code, and better readability for future uses. This can be seen in our analysis with the code running about 91.7% faster. Although our research shows there are advantages, there are some disadvantages of refactoring code which include an increased chance of errors and can be unnecessary. It could potentially contain many bugs after an attempt to refactor and can stop the original code from running.

