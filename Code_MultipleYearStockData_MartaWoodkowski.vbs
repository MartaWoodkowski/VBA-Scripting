Attribute VB_Name = "Module1"
Sub Stocks():
    Dim WriteRow As Integer
    Dim OpenFirstDay As Double
    Dim CloseLastDay As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim WS As Worksheet
    
    
    For Each WS In ActiveWorkbook.Worksheets
    
        'Headers
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
    
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        WriteRow = 2 'all results need to start writing from the second row
        TotalVolume = 0 'initialize Volume to 0
        OpenFirstDay = WS.Cells(2, 3).Value
    
        For i = 2 To LastRow
            If (WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value) Then
                'Ticker
                WS.Cells(WriteRow, 9) = WS.Cells(i, 1).Value
                'Yearly Change
                CloseLastDay = WS.Cells(i, 6).Value
                YearlyChange = CloseLastDay - OpenFirstDay
                WS.Cells(WriteRow, 10) = YearlyChange
            
                ' Color the interior of the cells
                If YearlyChange > 0 Then
                    WS.Cells(WriteRow, 10).Interior.ColorIndex = 4
                Else
                    WS.Cells(WriteRow, 10).Interior.ColorIndex = 3
                End If
                
                'Percentage Change
                If OpenFirstDay <> 0 Then
                    PercentageChange = (CloseLastDay - OpenFirstDay) / OpenFirstDay
                Else
                    PercentageChange = 0 'when there are zeros in data, I put the percentage change to 0 (from 0 to 0, means 0% change)
                End If
        
                WS.Cells(WriteRow, 11) = PercentageChange
                WS.Cells(WriteRow, 11).NumberFormat = "0.00%"
                TotalVolume = TotalVolume + WS.Cells(i, 7).Value
                WS.Cells(WriteRow, 12) = TotalVolume
                TotalVolume = 0 'reset the Total Volume
                
                'Increment row for next write
                WriteRow = WriteRow + 1
                OpenFirstDay = WS.Cells(i + 1, 3).Value
            Else
                'Total volume
                TotalVolume = TotalVolume + WS.Cells(i, 7).Value
            End If
        Next i
        
        
        'BONUS
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Total_Volume As Long
        
        'Headers
        WS.Cells(1, 15).Value = "Ticker"
        WS.Cells(1, 16).Value = "Value"
        
        WS.Cells(2, 14).Value = "Greatest % increase"
        WS.Cells(3, 14).Value = "Greatest % decrease"
        WS.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Calculating the Values
        WS.Cells(2, 16).Value = WorksheetFunction.Max(WS.Range("K2:K" & LastRow))
        WS.Cells(2, 16).NumberFormat = "0.00%" 'format the Value
        WS.Cells(3, 16).Value = WorksheetFunction.Min(WS.Range("K2:K" & LastRow))
        WS.Cells(3, 16).NumberFormat = "0.00%" 'format the Value
        WS.Cells(4, 16).Value = WorksheetFunction.Max(WS.Range("L2:L" & LastRow))
        WS.Cells(4, 16).NumberFormat = "0.0000E+00" 'format the Value
        
        'Identifying the Tickers
        For i = 2 To LastRow
            'Greatest % increase Ticker
            If (WS.Cells(2, 16).Value = WS.Cells(i, 11).Value) Then
            WS.Cells(2, 15).Value = WS.Cells(i, 9).Value
            
            'Greatest % decrease
            ElseIf (WS.Cells(3, 16).Value = WS.Cells(i, 11).Value) Then
            WS.Cells(3, 15).Value = WS.Cells(i, 9).Value
            
            'Greatest Total Volume
            ElseIf (WS.Cells(4, 16).Value = WS.Cells(i, 12).Value) Then
            WS.Cells(4, 15).Value = WS.Cells(i, 9).Value
            End If
            
        Next i
            
        WS.Columns("I:P").AutoFit

   Next WS
            
End Sub
