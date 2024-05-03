    VBA challenge HW
    
    Sub stocks():

    Dim i As Long ' row number
    Dim cell_vol As Double ' contents of column G
    Dim vol_total As Double ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I

    Dim k As Long ' leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double

    ' asked the xpert
    Dim lastRow As Long
    
    Dim ws As Worksheet
    
    'loop across all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    lastRow = ws.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    vol_total = 0
    k = 2
    
    ' Write Leaderboard Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ' assign open for first ticker
    ticker_open = ws.Cells(2, 3).Value
    
    For i = 2 To lastRow:
        cell_vol = ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        
        ' LOOP rows 2 to lastRow
        ' check if next row ticker is DIFFERENT
        ' if the same, then we only need to add to the vol_total
        ' if DIFFERENT, then we need add last row, write out to the leaderboard
        ' reset the vol_total to 0
        
        If (ws.Cells(i + 1, 1).Value <> ticker) Then
        ' oh JOLLY we have a different ticker
            vol_total = vol_total + cell_vol
        
        ' get closing price of ticker
            ticker_close = ws.Cells(i, 6).Value
            price_change = ticker_close - ticker_open
        
        ' Check if open price is 0 first to prevent hiccups
            If (ticker_open > 0) Then
                percent_change = price_change / ticker_open
            Else
                percent_change = 0
            End If
            
            ws.Cells(k, 9).Value = ticker
            ws.Cells(k, 10).Value = price_change
            ws.Cells(k, 11).Value = percent_change
            ws.Cells(k, 12).Value = vol_total
        
        ' formatting
            If (price_change > 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
            ElseIf (price_change < 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
            Else
                ws.Cells(k, 10).Interior.ColorIndex = 8 ' Cyan
            End If
        
        ' reset
            vol_total = 0
            k = k + 1
            ticker_open = ws.Cells(i + 1, 3).Value ' look ahead to get next ticker open
        Else
            ' we just add to the total
            vol_total = vol_total + cell_vol
        End If
    Next i
    
    ' Style leaderboard
    ' asked the xpert
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Columns("I:L").AutoFit

    ' Write Leaderboard #2 Columns
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
     ' take the max and min of Leaderboard #1
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    
    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    
     ' final ticker symbol for  total, greatest % of increase and decrease, and average
    ws.Range("O2") = ws.Cells(increase_number + 1, 9)
    ws.Range("O3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("O4") = ws.Cells(volume_number + 1, 9)
    
    Next ws
      
End Sub