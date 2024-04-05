Sub stockStats()

    Dim symbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim pctChange As Double
    Dim summaryRow As Long
    Dim runningVolume As Variant
    Dim lastRow As Long
    
    For Each ws In Worksheets
    
        symbol = ""
        openPrice = 0
        closePrice = 0
        yearlyChange = 0
        pctChange = 0
        summaryRow = 2
        runningVolume = 0
    
        ws.Range("I:Q").ClearContents
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        lastRow = WorksheetFunction.CountA(ws.Range("A:A"))
        
        For Row = 2 To lastRow + 1
            rowSymbol = ws.Cells(Row, 1).Value
            rowOpen = ws.Cells(Row, 3).Value
            rowClose = ws.Cells(Row - 1, 6).Value
            rowVolume = ws.Cells(Row, 7).Value
            If rowSymbol <> symbol Then
                If symbol = "" Then
                    openPrice = rowOpen
                Else
                    closePrice = rowClose
                    yearlyChange = closePrice - openPrice
                    pctChange = yearlyChange / openPrice
                    openPrice = rowOpen
                    ws.Cells(summaryRow, 9).Value = symbol
                    ws.Cells(summaryRow, 10).Value = yearlyChange
                    If yearlyChange > 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = vbGreen
                    ElseIf yearlyChange < 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = vbRed
                    End If
                    ws.Cells(summaryRow, 11).Value = pctChange
                    ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                    ws.Cells(summaryRow, 12).Value = runningVolume
                    summaryRow = summaryRow + 1
                End If
                runningVolume = rowVolume
                symbol = rowSymbol
            Else
                runningVolume = runningVolume + rowVolume
            End If
        Next Row
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("Q2").Value = "=MAX(K:K)"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = "=MIN(K:K)"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = "=MAX(L:L)"
        
        ws.Range("P2").Value = "=INDEX(I:I,MATCH(Q2,K:K,0),1)"
        ws.Range("P3").Value = "=INDEX(I:I,MATCH(Q3,K:K,0),1)"
        ws.Range("P4").Value = "=INDEX(I:I,MATCH(Q4,L:L,0),1)"
    
        ws.Columns("I:Q").AutoFit
        
    Next ws
    
    MsgBox "Finished"
    
End Sub