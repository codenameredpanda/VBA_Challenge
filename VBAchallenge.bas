Attribute VB_Name = "Module1"
Sub vba_challenge()

' declare variables

Dim brand As String
Dim brand_total As Double
brand_total = 0
Dim summary_table_row As Integer
Dim openprice As Double
Dim closeprice As Double
Dim Quarter As Double
Dim percent As Double
Dim max_percent As Double
Dim min_percent As Double


' loop through worksheets
For Each ws In Worksheets

' define summary tables
summary_table_row = 2

ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Quartely Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

ws.Range("n2").Value = "Greatest % Increase:"
ws.Range("n3").Value = "Greatest % Decrease:"
ws.Range("n4").Value = "Greatest Total Volume:"
ws.Range("o1").Value = "Ticker"
ws.Range("p1").Value = "Value"

' determine opening price
openprice = ws.Cells(2, 3).Value

' find last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
worksheetname = ws.Name

'loop through conditionals to find similar brands
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

brand = ws.Cells(i, 1).Value

closeprice = ws.Cells(i, 6).Value

Quarter = closeprice - openprice

percent = (closeprice - openprice) / openprice

brand_total = ws.Cells(i, 7).Value + brand_total
ws.Range("i" & summary_table_row).Value = brand
ws.Range("j" & summary_table_row).Value = Quarter
ws.Range("k" & summary_table_row).Value = percent
ws.Range("l" & summary_table_row).Value = brand_total

'loop through rows to apply conditional formatting
If ws.Range("j" & summary_table_row).Value < 0 Then
ws.Range("j" & summary_table_row).Interior.ColorIndex = 3
ElseIf ws.Range("j" & summary_table_row).Value > 0 Then
ws.Range("j" & summary_table_row).Interior.ColorIndex = 4

End If

summary_table_row = summary_table_row + 1
brand_total = 0

openprice = ws.Cells(i + 1, 3)

Else

brand_total = brand_total + ws.Cells(i, 7).Value

End If

Next i

lastrow_summary_table_row = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow_summary_table_row

If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("k2:k" & lastrow_summary_table_row)) Then
ws.Range("o2").Value = ws.Cells(i, 9).Value
ws.Range("p2").Value = ws.Cells(i, 11).Value

ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("k2:k" & lastrow_summary_table_row)) Then
ws.Range("o3").Value = ws.Cells(i, 9).Value
ws.Range("p3").Value = ws.Cells(i, 11).Value

ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("l2:l" & lastrow_summary_table_row)) Then
ws.Range("o4").Value = ws.Cells(i, 9).Value
ws.Range("p4").Value = ws.Cells(i, 12).Value

End If

Next i

ws.Range("k:k").NumberFormat = "0.00%"
ws.Range("p2:p3").NumberFormat = "0.00%"
ws.Range("i1:p1").EntireColumn.AutoFit



Next ws

End Sub
