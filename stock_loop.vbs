Attribute VB_Name = "Module1"
Sub ticker_loop()

For Each ws In Worksheets

Dim stock_ticker As String

Dim yearly_change As Double
yearly_change = 0

Dim total_vol As Double
total_vol = 0

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim stock_open As Double
Dim stock_close As Double

stock_open = ws.Cells(2, 3).Value

ws.Range("I" & (Summary_Table_Row - 1)).Value = "Ticker"

ws.Range("J" & (Summary_Table_Row - 1)).Value = "Yearly Change"

ws.Range("K" & (Summary_Table_Row - 1)).Value = "Percent Change"

ws.Range("L" & (Summary_Table_Row - 1)).Value = "Total Volume"

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

stock_ticker = ws.Cells(i, 1).Value

total_vol = total_vol + ws.Cells(i, 7).Value

stock_close = ws.Cells(i, 6).Value

yearly_change = yearly_change + (stock_close - stock_open)

percent_change = (yearly_change / stock_open)

ws.Range("I" & Summary_Table_Row).Value = stock_ticker

ws.Range("J" & Summary_Table_Row).Value = yearly_change

If yearly_change <= 0 Then ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

If yearly_change > 0 Then ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change)

ws.Range("L" & Summary_Table_Row).Value = total_vol

Summary_Table_Row = Summary_Table_Row + 1

total_vol = 0

yearly_change = 0

stock_open = ws.Cells(i + 1, 3)

Else

total_vol = total_vol + ws.Cells(i, 7).Value

End If

Next i

'Bonus Section
 ' -----------

Dim ticker As String
Dim maxIncrease As Double
Dim minIncrease As Double
Dim maxVolume As Double

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

maxIncrease = WorksheetFunction.max(ws.Range("K:K"))
ws.Range("P2").Value = FormatPercent(maxIncrease)

minIncrease = WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("P3").Value = FormatPercent(minIncrease)

maxVolume = WorksheetFunction.max(ws.Range("L:L"))
ws.Range("P4").Value = maxVolume

For j = 2 To lastrow

    If ws.Cells(j, 11).Value = maxIncrease Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O2").Value = ticker
    End If
    If ws.Cells(j, 11).Value = minIncrease Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O3").Value = ticker
    End If
    If ws.Cells(j, 12).Value = maxVolume Then
        ticker = ws.Cells(j, 9).Value
        ws.Range("O4").Value = ticker
     End If

Next j

Next ws

End Sub
