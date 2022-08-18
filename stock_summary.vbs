Sub stock_summary():

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

    Cells(1, 9).Value = "Ticker"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    stock_vol = 0
    
    stock_open = Cells(2, 3).Value
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stock_name = Cells(i, 1).Value
            stock_close = Cells(i, 6).Value
            stock_vol = stock_vol + Cells(i, 7).Value
            yearly_change = stock_close - stock_open
            percent_change = (stock_close / stock_open) - 1
            Range("I" & summary_table_row).Value = stock_name
            Range("J" & summary_table_row).Value = yearly_change
            Range("J" & summary_table_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            Range("J" & summary_table_row).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
            Range("J" & summary_table_row).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            Range("J" & summary_table_row).FormatConditions(2).Interior.Color = RGB(0, 255, 0)
            Range("K" & summary_table_row).Value = percent_change
            Range("K" & summary_table_row).NumberFormat = "0.00%"
            Range("L" & summary_table_row).Value = stock_vol
            summary_table_row = summary_table_row + 1
            stock_vol = 0
            stock_open = Cells(i + 1, 3).Value
        
        Else
            stock_vol = stock_vol + Cells(i, 7).Value
        End If
    Next i

max = Application.WorksheetFunction.max(Columns("K"))
Cells(2, 16).Value = max
min = Application.WorksheetFunction.min(Columns("K"))
Cells(3, 16).Value = min
max_stock_vol = Application.WorksheetFunction.max(Columns("L"))
Cells(4, 16).Value = max_stock_vol

Range("O2") = Application.WorksheetFunction.XLookup(Range("P2"), Range("K2:K91"), Range("I2:I91"))
Range("O3") = Application.WorksheetFunction.XLookup(Range("P3"), Range("K2:K91"), Range("I2:I91"))
Range("O4") = Application.WorksheetFunction.XLookup(Range("P4"), Range("L2:L91"), Range("I2:I91"))

End Sub
