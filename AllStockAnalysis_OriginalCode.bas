Attribute VB_Name = "Module6"
 Sub AllStocksAnalysis()
  
  Dim startTime As Single
  Dim endTime  As Single
 
 yearValue = InputBox("What year would like to run to the analysis on?")
    
    startTime = Timer
    
'Format the output sheet on the "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'To create header rows
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

'Initialize an array of all tickers

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

'Initialize variables for the starting price and ending price

Dim startingPrice As Single
Dim endingPrice As Single

'Activate the data worksheet

Sheets(yearValue).Activate

'Find the number of rows to loop over

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Loop through the tickers


For i = 0 To 11

ticker = tickers(i)
totalVolume = 0

'Loop through rows in the data

Sheets(yearValue).Activate
For j = 2 To RowCount

    'Find the total volume for the current ticker

    If Cells(j, 1).Value = ticker Then

        totalVolume = totalVolume + Cells(j, 8).Value
    
    End If
    
    'Find the starting price for the current ticker

    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

        startingPrice = Cells(j, 6).Value

    End If

    'Find the ending price for the current ticker

    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
        endingPrice = Cells(j, 6).Value
    
    End If

Next j


'Output the data for the current ticker

Worksheets("All Stocks Analysis").Activate

Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i
 
 endTime = Timer
 
 MsgBox "This code ran in" & endTime - startTime & "second for the year" & (yearValue)

End Sub
