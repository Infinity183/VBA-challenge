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

'We need a for loop that goes through every sheet.
Dim Page As Worksheet
For Each Page In ActiveWorkbook.worksheets

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

LastRow = Page.Cells(Rows.Count, 1).End(xlUp).Row
FirstPrice = Page.Cells(2, 3).Value
'By default, the very first price will always be the value in this cell.
Page.Range("I1").Value = "Ticker"
Page.Range("J1").Value = "Yearly Change"
Page.Range("K1").Value = "Percentage Change"
Page.Range("L1").Value = "Total Stock Volume"
'The new columns will now be titled.

For I = 2 To LastRow
    If Page.Cells(I + 1, 1).Value = Page.Cells(I, 1).Value Then
        StockCumulative = StockCumulative + Cells(I, 7).Value
    ElseIf Page.Cells(I + 1, 1).Value <> Page.Cells(I, 1).Value Then
        StockCumulative = StockCumulative + Page.Cells(I, 7).Value
        'This is the final time we're adding to the Stock Value.
        LastPrice = Page.Cells(I, 6).Value
        'Since we know this is the last cell of the given ticker,
        'we can now define the ending price.
        NetChange = LastPrice - FirstPrice
        If FirstPrice <> 0 Then
            PercentChange = (NetChange / FirstPrice)
        ElseIf FirstPrice = 0 Then
            PercentChange = 0
        End If
        Page.Cells(StockNumber + 1, 9).Value = Page.Cells(I, 1).Value
        Page.Cells(StockNumber + 1, 10).Value = NetChange
        Page.Cells(StockNumber + 1, 11).Value = PercentChange
        If PercentChange > MaxGain Then
            MaxGain = PercentChange
            Ticker1 = Page.Cells(StockNumber + 1, 9).Value
        End If
        If PercentChange < MaxLoss Then
            MaxLoss = PercentChange
            Ticker2 = Page.Cells(StockNumber + 1, 9).Value
        End If
        If StockCumulative > MaxVolume Then
            MaxVolume = StockCumulative
            Ticker3 = Page.Cells(StockNumber + 1, 9).Value
        End If
        'This gives us the percentage.
        Page.Cells(StockNumber + 1, 12).Value = StockCumulative
        'We will now review our values for the stock to see if it affects the Challenge results.

        'We need to reset the cumulative Stock Value for the next Ticker.
        StockCumulative = 0
        StockNumber = StockNumber + 1
        'Since we're moving on to a new Ticker, we can predict the FirstPrice
        'to be the <open> value of the following row.
        FirstPrice = Page.Cells(I + 1, 3).Value
    End If
Next I

'Let's fill in the Challenge cells.
Page.Range("O2").Value = "Greatest % Increase"
Page.Range("O3").Value = "Greatest % Decrease"
Page.Range("O4").Value = "Greatest Total Volume"
Page.Range("P1").Value = "Ticker"
Page.Range("Q1").Value = "Value"
Page.Range("P2").Value = Ticker1
Page.Range("P3").Value = Ticker2
Page.Range("P4").Value = Ticker3
Page.Range("Q2").Value = FormatPercent(MaxGain)
Page.Range("Q3").Value = FormatPercent(MaxLoss)
Page.Range("Q4").Value = MaxVolume

'Before moving on to the next sheet, we'll recolor the net change cells.
For j = 2 To StockNumber
    If Page.Cells(j, 10).Value > 0 Then
        Page.Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Page.Cells(j, 10).Value < 0 Then
        Page.Cells(j, 10).Interior.ColorIndex = 3
    ElseIf Page.Cells(j, 10).Value = 0 Then
        Page.Cells(j, 10).Interior.ColorIndex = 15
    End If
Next j
Next Page

End Sub