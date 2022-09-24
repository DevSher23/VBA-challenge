Attribute VB_Name = "Module12"
Sub Multipleyearstockdata()

Dim lrow As Long
Dim ticker_name As String
Dim total_volume As Double
Dim summary_table_row As Integer
Dim open_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double

For Each ws In Worksheets

ws.Activate

ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

total_volume = 0
summary_table_row = 2
open_price = Cells(2, 3).Value

'run to last row
lrow = Cells(Rows.Count, 1).End(xlUp).Row

'loop
For i = 2 To lrow

    total_volume = total_volume + Cells(i, 7).Value

    'ticker name
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_name = Cells(i, 1).Value
    

        'set the closing price
        closing_price = Cells(i, 6).Value

        'yearly_change
        yearly_change = closing_price - open_price

        'percent change from open to close at the end of year
        percent_change = (yearly_change / open_price) * 100

        'add to summary table
        ws.Range("i" & summary_table_row).Value = ticker_name
        ws.Range("l" & summary_table_row).Value = total_volume
        ws.Range("k" & summary_table_row).Value = percent_change
        ws.Range("j" & summary_table_row).Value = yearly_change
        
        'set new opening price
        open_price = Cells(i + 1, 3).Value
        
        'reset total volume
        total_volume = 0
 
        'make column k a percent
        ws.Range("k" & summary_table_row).NumberFormat = "0.00\%"

        'add color to yearly change
        If ws.Range("j" & summary_table_row).Value > 0 Then
            ws.Range("j" & summary_table_row).Interior.ColorIndex = 4
        Else
        'If ws.Range("j" & summary_table_row).Value < 0 Then
            ws.Range("j" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        summary_table_row = summary_table_row + 1
    End If
Next i

'make columns fit
ws.Columns("i:l").AutoFit

Next ws


End Sub
