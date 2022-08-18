Sub stock_summary():

For Each ws In Worksheets

    Dim stock_name As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_close As Double
    Dim stock_open As Double
    Dim stock_vol As LongLong
    Dim summary_table_row As Integer
    Dim max As Double
    Dim min As Double
    Dim max_stock_vol As LongLong
    summary_table_row = 2

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    stock_vol = 0
    
    stock_open = Cells(2, 3).Value
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            stock_name = ws.Cells(i, 1).Value
            stock_close = ws.Cells(i, 6).Value
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            yearly_change = stock_close - stock_open
            percent_change = (stock_close / stock_open) - 1
            ws.Range("I" & summary_table_row).Value = stock_name
            ws.Range("J" & summary_table_row).Value = yearly_change
            ws.Range("J" & summary_table_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            ws.Range("J" & summary_table_row).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
            ws.Range("J" & summary_table_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            ws.Range("J" & summary_table_row).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
            ws.Range("K" & summary_table_row).Value = percent_change
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            ws.Range("L" & summary_table_row).Value = stock_vol
            summary_table_row = summary_table_row + 1
            stock_vol = 0
            stock_open = ws.Cells(i + 1, 3).Value
        
        Else
            stock_vol = stock_vol + ws.Cells(i, 7).Value
        End If
    Next i

max = Application.WorksheetFunction.max(ws.Columns("K"))
ws.Cells(2, 16).Value = max
ws.Cells(2, 16).NumberFormat = "0.00%"
min = Application.WorksheetFunction.min(ws.Columns("K"))
ws.Cells(3, 16).Value = min
ws.Cells(3, 16).NumberFormat = "0.00%"
max_stock_vol = Application.WorksheetFunction.max(ws.Columns("L"))
ws.Cells(4, 16).Value = max_stock_vol

ws.Range("O2") = Application.WorksheetFunction.XLookup(ws.Range("P2"), ws.Range("K2:K91"), ws.Range("I2:I91"))
ws.Range("O3") = Application.WorksheetFunction.XLookup(ws.Range("P3"), ws.Range("K2:K91"), ws.Range("I2:I91"))
ws.Range("O4") = Application.WorksheetFunction.XLookup(ws.Range("P4"), ws.Range("L2:L91"), ws.Range("I2:I91"))

ws.Columns("I:P").AutoFit

Next ws

End Sub
