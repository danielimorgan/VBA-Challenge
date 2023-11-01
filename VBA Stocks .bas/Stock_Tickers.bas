Attribute VB_Name = "Module1"

Sub stock_tickers()

    'Loop the Worksheets
    For Each ws In Worksheets
    
        'Create a Variable to hold File Name and Last Row.
        Dim Worksheet_Name As String
        
        'Determine the Last Row.
        Dim LastRow As Long
        LastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
        
        'Grab the Worksheet name.
        Worksheet_Name = ws.Name
        
        'Set up Summary Table Areas.
        ws.Range("I1").Value = "Ticker"
        ws.Range("P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = Worksheet_Name
        'Set initial variable for holding the ticker name.
        Dim Ticker_Name As String
        
        'Set an initial variable for holding the total volume per ticker.
        Dim Ticker_Volume As Double
        
        'Set an initial value for year opening date.
        Dim Open_Date As Double
        
        Open_Date = ws.Cells(2, 2).Value
        
        'Set initial value for year closing date.
        Dim Closing_Date As Double
        
        'Set initial value for year opening price.
        Dim Open_Price As Double
        
        Open_Price = ws.Cells(2, 3).Value
        
        'Set initial value for year closing price.
        Dim Close_Price As Double
        
        'Keep track of the location of each Ticker in the Summary Table.
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Loop through all the tickers.
        For i = 2 To LastRow
            
            'Check to see if you're within the same ticker symbol. If not, then..
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker name.
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Add to the volume total.
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                'Print the ticker name.
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'Print the tcker total volume.
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
                
                'Print the year price change and change the color.
                ws.Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
                
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                Else
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                End If
        
                'Print the year price percentage.
                ws.Range("K" & Summary_Table_Row).Value = (Close_Price - Open_Price) / Open_Price * 100
                
                
                'Add one to the summary table row.
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset and rehash formulas.
                Ticker_Volume = 0
                
                Open_Date = ws.Cells(i + 1, 2).Value
                
                Open_Price = ws.Cells(i + 1, 3).Value
                
            'If the tickers are the same...
            Else
        
                'Set the year close date and the year close price.
                If ws.Cells(i, 2) >= Open_Date Then
                
                    Close_Date = ws.Cells(i, 2)
                    
                    Close_Price = ws.Cells(i, 6)
        
                End If
        
            'Add to the volume total.
            Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        
            End If
        
        Next i
        
    Next ws

End Sub
