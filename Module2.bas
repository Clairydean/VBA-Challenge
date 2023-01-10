Attribute VB_Name = "Module1"



Sub WORKSHEET_LOOP()
For Each ws In Worksheets
'Declaring variables
Dim ticker As String
Dim yearOpen, yearClose, yearHigh, yearLow, yearlyChange, summ As Double
Dim lastRow, stockVolume, i As LongLong
    stockVolume = 0
Dim Stock_Summary_Row As Integer
    Stock_Summary_Row = 2
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Create Headers
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
'Loop the Worksheets
For i = 2 To lastRow
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        yearOpen = ws.Cells(i, 3).Value
        ticker = ws.Cells(i, 1).Value
        summ = summ + ws.Cells(i, 7).Value
        percentChange = ws.Cells(i, 12)
    End If
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        yearClose = ws.Cells(i, 6).Value
        ticker = ws.Cells(i, 1).Value
         stockVolume = stockVolume + ws.Cells(i, 7).Value
            ws.Range("J" & Stock_Summary_Row).Value = ticker
        summ = summ + ws.Cells(i, 7).Value
            ws.Range("M" & Stock_Summary_Row).Value = stockVolume
        yearlyChange = yearClose - yearOpen
            ws.Range("K" & Stock_Summary_Row).Value = yearlyChange
        percentChange = yearlyChange / yearOpen
            ws.Range("L" & Stock_Summary_Row).Value = percentChange
        ws.Cells(Stock_Summary_Row, 11).Value = yearlyChange
    If yearlyChange < 0 Then
        ws.Cells(Stock_Summary_Row, 11).Interior.ColorIndex = 3
    ElseIf yearlyChange > 0 Then
        ws.Cells(Stock_Summary_Row, 11).Interior.ColorIndex = 4
    End If
        Stock_Summary_Row = Stock_Summary_Row + 1
        stockVolume = 0
    Else
    stockVolume = stockVolume + ws.Cells(i, 7).Value
    End If
Next i
Next ws
End Sub
