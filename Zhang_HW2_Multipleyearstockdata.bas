Attribute VB_Name = "RibbonX_Code"
Sub stock():

Dim last_row As Long
Dim year_change As Double
Dim ticker As String
Dim total_volume As Variant
Dim percent_change As Double
Dim summary_table_row As Integer
Dim start_point As Double

start_point = Cells(2, 3).Value
summary_table_row = 2
percent_change = 0
total_volume = 0

    With ActiveSheet
        last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greastest % Increase"
Cells(3, 15).Value = "Greastest % Decrease"
Cells(4, 15).Value = "Greastest Total Volume"

For i = 2 To last_row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

yearly_change = Cells(i, 6).Value - start_point
Range("J" & summary_table_row).Value = yearly_change
If yearly_change >= 0 Then
Range("J" & summary_table_row).Interior.ColorIndex = 4
Else
Range("J" & summary_table_row).Interior.ColorIndex = 3
End If

If start_point < 1E-30 And start_point > -1E-30 Then
Range("K" & summary_table_row).Value = "0.00%"
Else
percent_change = yearly_change / start_point
Range("K" & summary_table_row).Value = percent_change
Range("K" & summary_table_row).NumberFormat = "0.00%"
End If


ticker = Cells(i, 1).Value
Range("I" & summary_table_row).Value = ticker

total_volume = total_volume + Cells(i, 7).Value
Range("L" & summary_table_row).Value = total_volume

summary_table_row = summary_table_row + 1
total_volume = 0
start_point = Cells(i + 1, 3).Value

Else
total_volume = total_volume + Cells(i, 7).Value

End If
Next i

Dim greastest_percent As Double
Dim smalleset_percent As Double
Dim greatest_volume As LongLong

greatest_percent = Application.WorksheetFunction.Max(Columns("K"))
Cells(2, 17).Value = greatest_percent
Cells(2, 17).NumberFormat = "0.00%"

smallest_percent = Application.WorksheetFunction.Min(Columns("K"))
Cells(3, 17).Value = smallest_percent
Cells(3, 17).NumberFormat = "0.00%"

greatest_volume = Application.WorksheetFunction.Max(Columns("L"))
Cells(4, 17).Value = greatest_volume

For j = 2 To last_row
If Cells(j, 11).Value = greatest_percent Then
Cells(2, 16).Value = Cells(j, 9).Value
ElseIf Cells(j, 11).Value = smallest_percent Then
Cells(3, 16).Value = Cells(j, 9).Value
ElseIf Cells(j, 12).Value = greatest_volume Then
Cells(4, 16).Value = Cells(j, 9).Value
End If
Next j

End Sub

