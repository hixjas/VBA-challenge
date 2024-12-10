VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AnalyzeStocks()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
'Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0

Dim TickerRow As Long: TickerRow = 1
Dim percentChange As Double

'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Ticker symbol output
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerRow = TickerRow + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerRow, "I").Value = Ticker

'Calculate change in Price
close_price = ws.Cells(i, 6).Value
open_price = ws.Cells(i, 3).Value
price_change_percent = close_price - open_price
ws.Cells(TickerRow, "J").Value = price_change_percent

percentChange = (close_price / open_price)
ws.Cells(TickerRow, "K").Value = percentChange

volume = ws.Cells(i, 7).Value
ws.Cells(TickerRow, "L").Value = volume

'Fixing the open price equal zero problem
ElseIf open_price <> 0 Then
price_change_percent = (price_change_percent / open_price) * 100

End If

Next i

Next ws

End Sub

Sub ChangeColor()
Dim rCell As Range
With Sheet1
For Each rCell In .Range("J2:J1501")
If rCell.Value <= 0 Then
rCell.Interior.Color = vbRed
ElseIf rCell.Value >= 0 Then
rCell.Interior.Color = vbGreen
Else: rCell.Interior.Color = vbWhite
End If
Next
End With
End Sub
