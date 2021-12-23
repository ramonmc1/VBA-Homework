#  VBA HomeWork_Ramon Martinez

Sub VBAStock()

Set WB = ThisWorkbook

For Each ws In WB.Worksheets

ws.Activate
Dim TickerName As String
Dim LastRow As Long
Dim TicketCount As Long
Dim BegPrice As Double
Dim FinPrice As Double
Dim PriceChange As Double
Dim PercentChange As Double
Dim TotStockVol As Double
Dim InitialStockVol As Long
Dim FinalStockVol As Long
Dim Lasti As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = "Ticket Symbol"
Cells(1, 10).Value = "Price Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Initial conditions before commencing the loop
TicketCount = 2
TicketName = Cells(2, 1)
BegPrice = Cells(2, 3)
InitialStockVol = Cells(2, 7)
Lasti = 2

Cells(2, 9).Value = TicketName

For i = 2 To LastRow
    If Cells(i, 1) <> Cells(i + 1, 1) Then
    TicketCount = TicketCount + 1
    FinPrice = Cells(i, 6)
    FinalStockVol = Cells(i, 7)
    TicketName = Cells(i + 1, 1)
    PriceChange = FinPrice - BegPrice
        
        If BegPrice <> 0 Then
        PercentChange = PriceChange / BegPrice
        Else
        PercentChange = 0
        End If
    
    TotStockVol = Application.WorksheetFunction.Sum(Range("G" & Lasti & ":G" & i))
    Cells(TicketCount, 9).Value = TicketName
    Cells(TicketCount - 1, 10).Value = PriceChange
      If PriceChange <= 0 Then
      Cells(TicketCount - 1, 10).Interior.ColorIndex = 3
      Else
      Cells(TicketCount - 1, 10).Interior.ColorIndex = 4
      End If
    Cells(TicketCount - 1, 11).Value = PercentChange
    Cells(TicketCount - 1, 11).NumberFormat = "#.##%"
    Cells(TicketCount - 1, 12).Value = TotStockVol
    BegPrice = Cells(i + 1, 3)
    InitialStockVol = Cells(i + 1, 7)
    Lasti = i + 1
    End If
Next i

Dim LastRow2 As Long
Dim GreInc As Double
Dim GreDec As Double
Dim GreTotVol As Double
Dim TickerNameHigh As String
Dim TickerNameLow As String
Dim TickerNameVol As String

Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(2, 14) = "Greatest Increase"
Cells(3, 14) = "Greatest Decrease"
Cells(4, 14) = "Greatest Total Volume"

LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row

GreInc = Cells(2, 10).Value
GreDec = Cells(2, 11).Value
GreTotVol = Cells(2, 12).Value

For j = 2 To LastRow2
If Cells(j + 1, 10) > GreInc Then
GreInc = Cells(j + 1, 10).Value
TickerNameHigh = Cells(j + 1, 9)
End If
Next j

For i = 2 To LastRow2
If Cells(i + 1, 10) < GreDec Then
GreDec = Cells(i + 1, 10).Value
TickerNameLow = Cells(i + 1, 9)
End If
Next i

For j = 2 To LastRow2
If Cells(j + 1, 12) > GreTotVol Then
GreTotVol = Cells(j + 1, 12)
TickerNameVol = Cells(j + 1, 9)
End If
Next j

Cells(2, 16).Value = GreInc
Cells(2, 15).Value = TickerNameHigh
Cells(3, 16).Value = GreDec
Cells(3, 15).Value = TickerNameLow
Cells(4, 16).Value = GreTotVol
Cells(4, 15).Value = TickerNameVol

Columns("J:P").AutoFit
Next ws

End Sub



