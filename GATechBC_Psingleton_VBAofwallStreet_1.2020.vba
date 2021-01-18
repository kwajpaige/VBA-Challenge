Attribute VB_Name = "Module1"
Sub WallStreetVBA()

Dim ws As Worksheet

For Each ws In Worksheets

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As Double
Dim LastRow As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim ColumnI As Double
Dim StockValue As Double
Dim ColumnL As Double




ColumnI = 2
ColumnL = 2




ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

TotalStockVolume = 0

OpenPrice = ws.Cells(2, 3).Value



For i = 2 To LastRow



If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(ColumnI, 9).Value = ws.Cells(i, 1).Value
ClosePrice = ws.Cells(i, 6).Value

YearlyChange = ClosePrice - OpenPrice

ws.Cells(ColumnI, 12).Value = TotalStockVolume

If OpenPrice <> 0 Then
PercentChange = YearlyChange / OpenPrice * 100
If PercentChange > 0 Then
ws.Cells(ColumnI, 11).Interior.ColorIndex = 4
Else
ws.Cells(ColumnI, 11).Interior.ColorIndex = 3

End If

Else

End If

OpenPrice = ws.Cells(i + 1, 3).Value

ws.Cells(ColumnI, 10).Value = YearlyChange
ws.Cells(ColumnI, 11).Value = PercentChange

ColumnI = ColumnI + 1
TotalStockVolume = 0


Else




TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value


End If

YearlyChange = 0
ClosePrice = 0


Next i

Next ws

End Sub



