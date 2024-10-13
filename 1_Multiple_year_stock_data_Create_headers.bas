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
