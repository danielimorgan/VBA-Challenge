Attribute VB_Name = "Module2"
Sub great_stocks()

    'Loop the Worksheets
    For Each ws In Worksheets
    
        'Find the last row in the summary.
        SummaryLastRow = ws.Range("I" & ws.Rows.Count).End(xlUp).Row
       
        'Set initial value for Greatest % Increase and ticker.
        Dim Great_Pct_Inc As Double
        Dim Ticker_Pct_Inc As String
        
        Great_Pct_Inc = ws.Cells(2, 11).Value
        
        'Set initial value for Greatest % Decrease and ticker.
        Dim Great_Pct_Dec As Double
        Dim Ticker_Pct_Dec As String
        
        Great_Pct_Dec = ws.Cells(2, 11).Value
        
        'Set initial value for Greatest Total Volume and ticker.
        Dim Great_Volume As Double
        Dim Great_Ticker_Volume As String
        
        Great_Volume = ws.Cells(2, 12).Value
        
        'Find the Greatest %s and Greates Total Volume.
        For j = 2 To SummaryLastRow
        
            If ws.Cells(j, 12).Value > Great_Volume Then
            
                Great_Volume = ws.Cells(j, 12).Value
                
                Great_Ticker_Volume = ws.Cells(j, 9).Value
                
                ws.Range("P4").Value = Ticker_Volume
                
                ws.Range("Q4").Value = Great_Volume
            
            End If
        
            If ws.Cells(j, 11).Value > Great_Pct_Inc Then
            
                Great_Pct_Inc = ws.Cells(j, 11).Value
                
                Ticker_Pct_Inc = ws.Cells(j, 9).Value
                
                ws.Range("P2").Value = Ticker_Pct_Inc
                
                ws.Range("Q2").Value = Great_Pct_Inc
            
            End If
            
            If ws.Cells(j, 11).Value < Great_Pct_Dec Then
                
                Great_Pct_Dec = ws.Cells(j, 11).Value
                
                Ticker_Pct_Dec = ws.Cells(j, 9).Value
                
                ws.Range("P3").Value = Ticker_Pct_Dec
                
                ws.Range("Q3").Value = Great_Pct_Dec
                
            End If
        
        Next j
    
    Next ws

End Sub
