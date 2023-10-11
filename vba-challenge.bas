Attribute VB_Name = "Module1"
Sub Ticker()

    'Loop All Worksheets'
    For Each ws In Worksheets
        
     'Create labels'
      ws.Cells(1, 9) = "Ticker"
      ws.Cells(1, 10) = "Yearly Change"
      ws.Cells(1, 11) = "Percent Change"
      ws.Cells(1, 12) = "Total Stock Volume"
      ws.Cells(1, 15) = "Ticker"
      ws.Cells(1, 16) = "Value"
      ws.Cells(2, 14) = "Greatest % Increase"
      ws.Cells(3, 14) = "Greatest % Decrease"
      ws.Cells(4, 14) = "Greatest Total Volume"
      
      'Find Lastrow'
      
      Lastrow = ws.Cells(Rows.Count, 2).End(xlUp).Row
      
      
    'Set Ticker Row'
    Tickerrow = 2
    
    'Set Open Row'
    
    YearOpen = 2
    
    'Name worksheet'
    
    Worksheetname = ws.Name
    
    'Volume Count'
    
        For R = 2 To Lastrow
        
            If ws.Cells(R + 1, 1) <> ws.Cells(R, 1) Then
            
            ws.Cells(Tickerrow, 9) = ws.Cells(R, 1).Value
            ws.Cells(Tickerrow, 12) = Volumesum + ws.Cells(R, 7)
            
            
            'Create Year Close Price'
            YearClose = ws.Cells(R, 6)
            
            ws.Cells(Tickerrow, 10) = YearClose - ws.Cells(YearOpen, 3)
            ws.Cells(Tickerrow, 11) = (YearClose - ws.Cells(YearOpen, 3)) / ws.Cells(YearOpen, 3)
            
                'Color Index'
                
                If ws.Cells(Tickerrow, 10) < 0 Then
                
                ws.Cells(Tickerrow, 10).Interior.ColorIndex = 3
                
                ElseIf ws.Cells(Tickerrow, 10) > 0 Then
                
                ws.Cells(Tickerrow, 10).Interior.ColorIndex = 4
            
                End If
                
                If ws.Cells(Tickerrow, 11) < 0 Then
                
                ws.Cells(Tickerrow, 11).Interior.ColorIndex = 3
                
                ElseIf ws.Cells(Tickerrow, 11) > 0 Then
                
                ws.Cells(Tickerrow, 11).Interior.ColorIndex = 4
                
                End If
                
            'Reset Year Open Price Row'
            
            YearOpen = R + 1
            
            'Reset Volumesum'
            Volumesum = 0
            
            'reset for next ticker'
            Tickerrow = Tickerrow + 1
            
            Else
            
            Volumesum = Volumesum + ws.Cells(R, 7).Value
            
            End If
        
        Next R
        
            'Find Greatest Values'
            
            GreatestPercentIncrease = ws.Cells(2, 11).Value
            GreatestPercentDecrease = ws.Cells(2, 11).Value
            GreatestVolume = ws.Cells(2, 12)
            
            Lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            For R2 = 2 To Lastrow2
            
                If ws.Cells(R2 + 1, 11) > GreatestPercentIncrease Then
                
                GreatestPercentIncrease = ws.Cells(R2 + 1, 11).Value
                GreatestPercentTicker = ws.Cells(R2 + 1, 9).Value
                
                End If
                
                If ws.Cells(R2 + 1, 11) < GreatestPercentDecrease Then
                
                GreatestPercentDecrease = ws.Cells(R2 + 1, 11).Value
                GreatestPercentDecreaseTicker = ws.Cells(R2 + 1, 9).Value
            
                End If
                
                If ws.Cells(R2 + 1, 12) > GreatestVolume Then
                
                GreatestVolume = ws.Cells(R2 + 1, 12).Value
                GreatestVolumeTicker = ws.Cells(R2 + 1, 9).Value
            
                End If
                
                Next R2
        
                ws.Cells(2, 16) = GreatestPercentIncrease
                ws.Cells(2, 15) = GreatestPercentTicker
                ws.Cells(3, 16) = GreatestPercentDecrease
                ws.Cells(3, 15) = GreatestPercentDecreaseTicker
                ws.Cells(4, 16) = GreatestVolume
                ws.Cells(4, 15) = GreatestVolumeTicker
        
        'Apply Formatting'
        
        ws.Range("K2:K" & Lastrow).NumberFormat = "0.00%"
        ws.Range("P2, P3").NumberFormat = "0.00%"
        
    Next ws

End Sub
