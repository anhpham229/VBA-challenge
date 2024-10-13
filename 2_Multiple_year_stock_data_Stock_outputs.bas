Attribute VB_Name = "Module1"
Sub AddText()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Quarterly change"
        ws.Range("K1") = "Percent change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    Next ws

End Sub

Sub stock_volume()

    Dim ws As Worksheet
    Dim up_row As Long
    Dim Qtr_change As Double
    Dim Per_change As Double
    Dim total_vol As Double
    Dim open_rate As Double
    
    For Each ws In ThisWorkbook.Worksheets
    up_row = 2
    total_vol = 0
    open_rate = ws.Cells(2, 3).Value
    
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If IsNumeric(ws.Cells(i, 7).Value) Then
                 total_vol = total_vol + ws.Cells(i, 7).Value
            End If
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Range("I" & up_row) = ws.Cells(i, 1).Value
                Qtr_change = ws.Cells(i, 6).Value - open_rate
                Per_change = Qtr_change / open_rate
                ws.Range("J" & up_row) = Qtr_change
                ws.Range("K" & up_row) = Per_change
                ws.Range("K" & up_row).NumberFormat = "0.00%"
                ws.Range("L" & up_row) = total_vol
                up_row = up_row + 1
                total_vol = 0
                open_rate = ws.Cells(i + 1, 3)
            End If
        Next i
    Next ws

End Sub
