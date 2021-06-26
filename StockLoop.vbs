Attribute VB_Name = "Module1"
Sub stockloop()

Dim lastrow As Long
Dim ticker_name As String
Dim Startdate As String
Dim enddate As Long
Dim rowcount As Integer
Dim startrow As Long
Dim endrow As Long
Dim percentagechange As Long

Dim TotalWorkSheets As Integer
Dim summary_table_row As Long
Dim j As Long
Dim TotalSheets As Long

TotalSheets = Application.Worksheets.Count

For j = 1 To TotalSheets

Sheets(j).Select
ActiveSheet.Range("A1").Value = "<ticker>"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


Range("I2").Value = WorksheetFunction.Max(Range("B2:B" & lastrow))

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"


summary_table_row = 2



For i = 2 To lastrow

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

'Ticker symbols in Column I

ticker_name = Cells(i, 1).Value

Range("I" & summary_table_row).Value = ticker_name

rowcount = Application.WorksheetFunction.CountIf(Columns(1), ticker_name)

'Stock Value

enddate = Application.WorksheetFunction.Max(Range("B" & i + 1 - rowcount & ":B" & i))

endrow = Range("B" & i + 1 - rowcount & ":B" & i).Find(enddate).Row

Startdate = Application.WorksheetFunction.Min(Range("B" & i + 1 - rowcount & ":B" & i))

startrow = Range("B" & i + 1 - rowcount & ":B" & i).Find(Startdate).Row

Range("J" & summary_table_row).Value = Cells(endrow, 6) - Cells(startrow, 3)


'percentage change
If Cells(startrow, 3).Value <> 0 Then

Range("K" & summary_table_row).Value = (Cells(summary_table_row, 10).Value) / Cells(startrow, 3).Value

End If

'Conditional Formatting

If Cells(summary_table_row, 11) > 0 Then

Cells(summary_table_row, 11).Interior.ColorIndex = 4

ElseIf Cells(summary_table_row, 11) < 0 Then

Cells(summary_table_row, 11).Interior.ColorIndex = 3

ElseIf Cells(summary_table_row, 11) = 0 Or Cells(summary_table_row, 11) = "" Then

Cells(summary_table_row, 11).Interior.ColorIndex = 0

End If

'sum of stock
Dim stocksum As String

stocksum = Application.WorksheetFunction.Sum(Range("G" & i + 1 - rowcount & ":G" & i))

Range("L" & summary_table_row).Value = stocksum

summary_table_row = summary_table_row + 1

End If

Next i


Dim PercentIncrease As Double
Dim PercentIncreaseRow As Integer
Dim PercentIncreaseticker As String
Dim PercentDecrease As Double
Dim PercentDecreaseRow As Integer
Dim PercentDecreaseticker As String
Dim TotalVolume As Double
Dim TotalVolumeTicker As String
Dim summarylastrow As String

summarylastrow = Cells(Rows.Count, 9).End(xlUp).Row

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

PercentIncrease = Application.WorksheetFunction.Max(Range("K2:K" & summarylastrow))
PercentIncreaseRow = Application.WorksheetFunction.Match(PercentIncrease, Range("K2:K" & summarylastrow), 0)
PercentIncreaseticker = Application.WorksheetFunction.Index(Range("I2:I" & summarylastrow), PercentIncreaseRow)

Range("P2").Value = PercentIncreaseticker
Range("Q2").Value = PercentIncrease

PercentDecrease = Application.WorksheetFunction.Min(Range("K2:K" & summarylastrow))
PercentDecreaseRow = Application.WorksheetFunction.Match(PercentDecrease, Range("K2:K" & summarylastrow), 0)
PercentDecreaseticker = Application.WorksheetFunction.Index(Range("I2:I" & summarylastrow), PercentDecreaseRow)

Range("P3").Value = PercentDecreaseticker
Range("Q3").Value = PercentDecrease

TotalVolume = Application.WorksheetFunction.Max(Range("L2:L" & summarylastrow))
TotalVolumeRow = Application.WorksheetFunction.Match(TotalVolume, Range("L2:L" & summarylastrow), 0)
TotalVolumeTicker = Application.WorksheetFunction.Index(Range("I2:I" & summarylastrow), TotalVolumeRow)

Range("P4").Value = TotalVolumeTicker
Range("Q4").Value = TotalVolume


Next j

End Sub
