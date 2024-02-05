# challenge-VBA
Create a script that loops through all the stocks for one year and outputs the following information:
he ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
Sub datastock()

Dim ws As Worksheet
Application.ScreenUpdating = False
For Each ws In Worksheets
ws.Select
Call dataloop
Next
Application.ScreenUpdating = True
End Sub

Sub dataloop()

Dim tickername As String
Dim tickervolume As Double
tickervolume = 0
Dim summary_ticker_row As Integer
summary_ticker_row = 2
Dim open_price As Double
open_price = Cells(2, 3)

Dim close_price As Double
Dim yeraly_change As Double
Dim percent_change As Double

Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "total stock volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To lastrow

If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
tickername = Cells(I, 1).Value
tickervolume = tickervolume + Cells(I, 7).Value
Range("I" & summary_ticker_row).Value = tickername
Range("L" & summary_ticker_row).Value = tickervolume
close_price = Cells(I, 6).Value
yearly_change = (close_price - open_price)
Range("J" & summary_ticker_row).Value = yearly_change

If open_price = 0 Then
percent_change = 0
Else
percent_change = yearly_change / open_price
End If

Range("K" & summary_ticker_row).Value = percent_change
Range("K" & summary_ticker_row).NumberFormat = "0.00%"

summary_ticker_row = summary_ticker_row + 1

tickervolume = 0
open_price = Cells(I + 1, 3)
Else
tickervolume = tickervolume + Cells(I, 7).Value
End If
Next I

lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

For I = 2 To lastrow_summary_table
If Cells(I, 10).Value > 0 Then
Cells(I, 10).Interior.ColorIndex = 10
Else
Cells(I, 10).Interior.ColorIndex = 3
End If
Next I

Cells(2, 15).Value = "greatest %increase"
Cells(3, 15).Value = "greatest % decrease"
Cells(4, 15).Value = "greatest total volume"
Cells(1, 16).Value = "ticker"
Cells(1, 17).Value = "value"

For I = 2 To lastrow_summary_table
If Cells(I, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
Cells(2, 16).Value = Cells(I, 9).Value
Cells(2, 17).Value = Cells(I, 11).Value
Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(I, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
Cells(3, 16).Value = Cells(I, 9).Value
Cells(3, 17).Value = Cells(I, 11).Value
Cells(3, 17).NumberFormat = "0.00%"

ElseIf Cells(I, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
Cells(4, 16).Value = Cells(I, 9).Value
Cells(4, 17).Value = Cells(I, 12).Value

End If
Next I



End Sub




