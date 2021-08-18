Attribute VB_Name = "Module1"
Sub KBowman_HW2()

For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    Dim OpenPrice As Double
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    Dim TotalStockVolume As LongLong
    TotalStockVolume = 0
           
    For i = 2 To LastRow
        
        Dim TickerDate As Long
        TickerDate = ws.Cells(i, 2).Value
        
            If Right(TickerDate, 4) = "0101" Then
                OpenPrice = ws.Cells(i, 3)
            End If
        
        'assign OpenPrice value
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Ticker = ws.Cells(i, 1).Value
            YearEndClose = ws.Cells(i, 6).Value
            ws.Range("I" & SummaryTableRow).Value = Ticker
            'ticker letter added to summary table, year end value defined
        
        
            
        YearlyChange = YearEndClose - OpenPrice
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        'calculate yearly change with stop cell value
        
        If OpenPrice = 0 Then
            ws.Range("K" & SummaryTableRow).Value = "n/a"
            
            Else

        PercentChange = YearlyChange / OpenPrice
        ws.Range("K" & SummaryTableRow).Value = PercentChange
        ws.Range("K" & SummaryTableRow).Style = "Percent"
        
        End If
            
        SummaryTableRow = SummaryTableRow + 1
        TotalStockVolume = 0
    
    End If
    
    Next i
    
    'color coding below
        
    For i = 2 To LastRow
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
            
        
        Next i
           
    
        
Next ws


End Sub

