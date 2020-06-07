Sub stockLooper():
    Dim symbol, greatUpSym, greatDownSym, greatVolSym As String
    symbol = Cells(2, 1).Value
    Dim startPrice, endPrice, rollingVolumeSum, greatUp, greatDown, greatVol As Double
    startPrice = Cells(2, 3).Value
    rollingVolumeSum = 0
    greatUp = 0
    greatDown = 0
    greatVol = 0
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Range("K:K").NumberFormat = "0.00%"
    Range("P2:P3").NumberFormat = "0.00%"
    
    Dim i As Long
    Dim j As Long
    j = 2
    'iterate over each row
    For i = 2 To Rows.Count
        If Cells(i, 7).Value > greatVol Then
            greatVol = Cells(i, 7).Value
            greatVolSym = Cells(i, 1).Value
        End If
        'if column A matches current symbol, add to rollingVolumeSum
        If symbol = Cells(i, 1) Then
            rollingVolumeSum = rollingVolumeSum + Cells(i, 7).Value
        'if not, populate a row in the summary table...
        Else
            'the Dec 31 entry for a stock corresponds with the row before the symbol change
            endPrice = Cells(i - 1, 6).Value
            Cells(j, 9).Value = symbol
            Cells(j, 10).Value = endPrice - startPrice
            If Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            Else
                Cells(j, 10).Interior.ColorIndex = 4
            End If
            If startPrice = 0 Then
                Cells(j, 11).Value = "N/A"
            Else
                Cells(j, 11).Value = (endPrice - startPrice) / startPrice
                If Cells(j, 11).Value > greatUp Then
                    greatUp = Cells(j, 11).Value
                    greatUpSym = Cells(j, 1).Value
                End If
                If Cells(j, 11).Value < greatDown Then
                    greatDown = Cells(j, 11).Value
                    greatDownSym = Cells(j, 1).Value
                End If
            End If
            Cells(j, 12).Value = rollingVolumeSum
            '...and re-initalize values
            symbol = Cells(i, 1).Value
            startPrice = Cells(i, 3).Value
            rollingVolumeSum = 0
            j = j + 1
        End If
    Next i
    'populate superlatives chart after the main loop has run
    Cells(2, 15).Value = greatUpSym
    Cells(2, 16).Value = greatUp
    Cells(3, 15).Value = greatDownSym
    Cells(3, 16).Value = greatDown
    Cells(4, 15).Value = greatVolSym
    Cells(4, 16).Value = greatVol
    Columns("I:P").AutoFit
End Sub