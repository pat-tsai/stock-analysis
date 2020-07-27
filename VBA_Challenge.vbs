Attribute VB_Name = "Module6"
Sub AllStocksAnalysisRefactored()

    'performance timer variables
    Dim startTime, EndTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'timer start
    startTime = Timer
    
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
       
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    '2) Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '3) Initialize array of all tickers
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '4b) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '5a) Create a ticker Index
    tickerIndex = 0

    '5b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12), tickerEndingPrices(12) As Single
    
    '6a) Initializing ticker volumes to zero
    For i = 0 To 11
        tickerIndex = i
        tickerVolumes(tickerIndex) = 0
    Next i
        
    tickerIndex = 0
    
    '6b) loop over all the rows
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i
    
    '9) Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    'output data sheet length
    dataRowStart = 4
    dataRowEnd = 15

    
    For i = dataRowStart To dataRowEnd
    'shading positive return rates as green 
        If Cells(i, 3) > 0 Then
           
            Cells(i, 3).Interior.Color = vbGreen
        
    'shading negative return rates as red    
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i

'timer end
EndTime = Timer

MsgBox "Analysis completed in " & (EndTime - startTime) & " seconds for the year of " & (yearValue)
End Sub
