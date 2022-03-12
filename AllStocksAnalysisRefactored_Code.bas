Attribute VB_Name = "Module8"
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

'Activate the worksheet
Worksheets(yearValue).Activate

'Number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Creating ticker Index(1a)
tickerIndex = 0

'Creating three output arrays(1b)

Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

'Create a for loop to initialize the tickerVolumes to zero(2a)
'If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
    tickerVolumes(i) = 0
Next i

'Loop over all the rows in the spreadsheet. (2b)

For j = 2 To RowCount

    'Increase volume for current ticker(3a)
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    
    'Check if the current row is the first row with the selected tickerIndex.(3b)
    
    If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
   
   End If
    
    'Check if the current row is the last row with the selected ticker(3c)

     If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
     
     End If

    'Increase the tickerIndex.(3d)
         
     If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        
        End If

    Next j
    
'4Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
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

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
