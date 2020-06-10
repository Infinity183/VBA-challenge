Sub StockGauger()

Dim StockNumber As Double
Dim LastRow As Double
Dim StockCumulative As Double
Dim FirstPrice As Double
Dim LastPrice As Double
Dim NetChange As Double
Dim PercentChange As Double
Dim MaxGain As Double
Dim MaxLoss As Double
Dim MaxVolume As Double
Dim Ticker1 As String
Dim Ticker2 As String
Dim Ticker3 As String

StockNumber = 1
StockCumulative = 0
NetChange = 0
PercentChange = 0
MaxGain = 0
MaxLoss = 0
MaxVolume = 0
Ticker1 = ""
Ticker2 = ""
Ticker3 = ""

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
FirstPrice = Cells(2, 3).Value
'By default, the very first price will always be the value in this cell.
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
'The new columns will now be titled.

For I = 2 To LastRow
    If Cells(I + 1, 1).Value = Cells(I, 1).Value Then
        StockCumulative = StockCumulative + Cells(I, 7).Value
    ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        LastPrice = Cells(I, 6).Value
        'Since we know this is the last cell of the given ticker,
        'we can now define the ending price.
        NetChange = LastPrice - FirstPrice
        If FirstPrice <> 0 Then
            PercentChange = (NetChange / FirstPrice)
        ElseIf FirstPrice = 0 Then
            PercentChange = 0
        End If
        Cells(StockNumber + 1, 9).Value = Cells(I, 1).Value
        Cells(StockNumber + 1, 10).Value = NetChange
        Cells(StockNumber + 1, 11).Value = PercentChange
        If PercentChange > MaxGain Then
            MaxGain = PercentChange
            Ticker1 = Cells(StockNumber + 1, 9).Value
        End If
        If PercentChange < MaxLoss Then
            MaxLoss = PercentChange
            Ticker2 = Cells(StockNumber + 1, 9).Value
        End If
        If StockCumulative > MaxVolume Then
            MaxVolume = StockCumulative
            Ticker3 = Cells(StockNumber + 1, 9).Value
        End If
        'This gives us the percentage.
        Cells(StockNumber + 1, 12).Value = StockCumulative
        'We will now review our values for the stock to see if it affects the Challenge results.

        'We need to reset the cumulative Stock Value for the next Ticker.
        StockCumulative = 0
        StockNumber = StockNumber + 1
        'Since we're moving on to a new Ticker, we can predict the FirstPrice
        'to be the <open> value of the following row.
        FirstPrice = Cells(I + 1, 3).Value
    End If
Next I

'Let's fill in the Challenge cells.
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("P2").Value = Ticker1
Range("P3").Value = Ticker2
Range("P4").Value = Ticker3
Range("Q2").Value = FormatPercent(MaxGain)
Range("Q3").Value = FormatPercent(MaxLoss)
Range("Q4").Value = MaxVolume


'Before moving on to the next sheet, we'll recolor the net change cells.
For j = 2 To StockNumber
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    ElseIf Cells(j, 10).Value = 0 Then
        Cells(j, 10).Interior.ColorIndex = 15
    End If
Next j

End Sub

