Sub stock_analysis()
    ' Declare variables
    Dim total As Double
    Dim tickername As String
    Dim tickervolume As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim averageChange As Double
    ' Set column titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    ' Initialize variables
    j = 2
    total = 0
    change = 0
    start = 2
    ' Find the last row in column A
    rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop through the data
    For i = 2 To rowCount
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Calculate quarterly change
            change = Cells(i, 6).Value - Cells(i, 3).Value
            ' Calculate percentage change
            If Cells(i, 3).Value <> 0 Then
                percentChange = (change / Cells(i, 3).Value) * 100
            Else
                percentChange = 0
            End If
            ' Update total stock volume
            total = total + Cells(i, 7).Value
            ' Print results
            Range("I" & j).Value = Cells(i, 1).Value
            Range("J" & j).Value = change
            Range("K" & j).Value = percentChange
            Range("L" & j).Value = total
            ' Reset variables for the next ticker
            total = 0
            change = 0
            j = j + 1
        End If
    Next i
    ' Find the stock with the greatest % increase, % decrease, and total volume
    ' You can add this logic here
Range("O2").Value = maxTicker
Range("P2").Value = maxIncrease & "%"
Range("O3").Value = minTicker
Range("P3").Value = maxDecrease & "%"
Range("O4").Value = volumeTicker
Range("P4").Value = maxVolume

' Update Q3 and Q4 with the respective values
Range("Q3").Value = minTicker
Range("Q4").Value = volumeTicker
Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
Range("P4") = WorksheetFunction.Max(Range("L2:L" & rowCount))
End Sub
