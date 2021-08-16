# STOCK ANALYSIS WITH VBA EXCEL

## OVERVIEW: VBA Stock Analysis Project

### Purpose
In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

### Analysis and Challenges
Here's a quick look at the Kickstarting Analysis and Challenges of this Project, including the following tasks:
- Creating a ticker Index
- Creating output arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
- The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays
- The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and - tickerEndingPrices
- Making sure the Code for formatting the cells in the spreadsheet is working
- Comments to explain the purpose of the code
- The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module
- The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

> Using the knowledge of VBA and the starter code provided in this Project to refactor the VBA Script dataset so we loop through the data one time and collect all of the information.

#### Our Challenge Data Background
> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

> In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

### The Data
The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

## RESULTS: Refactor VBA Code and Measure Performance
In order to make my code more efficient, I needed to switch the nesting order of my for loops. To do this, I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays with the tickers array by using a variable called the tickerIndex.The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.

### Refactored Code
    Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStockAnalysis").Activate
    
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
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
  
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStockAnalysis").Activate
            tickerIndex = i
            Cells(4 + i, 1).Value = tickers(tickerIndex)
            Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
            Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
    'Formatting
    Worksheets("AllStockAnalysis").Activate
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

### Original Code
    Sub AllStockAnalysis()

        Dim startTime As Single
        Dim endTime  As Single
    
        yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet
        Worksheets("AllStockAnalysis").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

    '2) Initialize array of all tickers
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
    
    '3a) Initialize variables for starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
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
            Worksheets("AllStockAnalysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
   
    'Formatting
        Worksheets("AllStockAnalysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("A3:C3").Borders(xlEdgeBottom).Weight = xlThick
        Range("A3:C3").Font.Size = 12
        Range("A4:A15").Font.Bold = True
        Range("A4:A15").Borders(xlEdgeRight).LineStyle = xlContinous
        Range("A4:A15").Borders(xlEdgeBottom).LineStyle = xlContinous
        Range("A4:A15").Borders(xlEdgeRight).Weight = xlThick
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.00%"
        Columns("B").AutoFit
   
    Worksheets("AllStockAnalysis").Activate
        dataRowStart = 4
        dataRowEnd = 15
      For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then

            'Change cell color to red
            Cells(i, 3).Interior.Color = vbRed
            
        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

      Next i
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    
    End Sub



> Finally, we run the stock analysis for 2017 and 2018 using both **Original** and **Refactored** codes, to confirm that our stock analysis outputs are the same as dataset example provided (as shown in the images below, named **Dataset Examples Provided**). Below are the final Stock Analysis Results named, **Final VBA Analysis 2017 and 2018** and the pop-up messages showing elapsed run time for the refactored code as VBA_StockAnalysis_2017.png and VBA_StockAnalysis_2018.png.

#### Dataset Examples Provided

![Dataset examples provided](https://user-images.githubusercontent.com/85645485/129618121-73827285-978c-4384-b944-05a456c52ba4.PNG)

#### Final VBA Analysis 2017
<img width="490" alt="Calling_2017_dataset" src="https://user-images.githubusercontent.com/85645485/129618117-6a918aee-0fcf-440a-97a0-8ca9b30b6603.png">
<img width="391" alt="VBA_StockAnalysis_2017" src="https://user-images.githubusercontent.com/85645485/129618108-dcb520d4-c18e-4582-968e-0b6f82bcaba9.png">

#### Final VBA Analysis 2018
<img width="391" alt="VBA_StockAnalysis_2018" src="https://user-images.githubusercontent.com/85645485/129618113-1078c715-2619-4a81-b40d-e6640f2f9083.png">

### Run-time for Each Method and yearValue

> Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.

Run-time for 2017 Stock Analysis using Original Code
<img width="265" alt="OriginalCode_RunTime_2017" src="https://user-images.githubusercontent.com/85645485/129618122-8b518744-8960-4bec-b99d-c740cc90d759.png">
Run-time for 2017 Stock Analysis using Refactored Code
<img width="265" alt="RefactoredCode_RunTime_2017" src="https://user-images.githubusercontent.com/85645485/129618124-d41e493a-a3ba-47ae-88d2-80d2d240c9fc.png">
Run-time for 2018 Stock Analysis using Original Code
<img width="266" alt="OriginalCode_RunTime_2018" src="https://user-images.githubusercontent.com/85645485/129618123-e87a8053-ced8-4049-b620-f670f853503f.png">
Run-time for 2018 Stock Analysis using Refactored Code
<img width="266" alt="RefactoredCode_RunTime_2018" src="https://user-images.githubusercontent.com/85645485/129618125-786761e6-e4a5-4844-86db-399ef38ca8df.png">

