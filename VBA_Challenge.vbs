    
    Sub VBAChallengeStockData():

        For Each ws In Worksheets
        
            Dim ActiveWorkSheetName As String
            Dim i As Long
            Dim j As Long
            Dim TickerRowCount As Long
            Dim LastRowA As Long
            Dim LastRowI As Long
            Dim PerChange As Double
            Dim GreatIncr As Double
            Dim GreatDecr As Double
            Dim GreatVol As Double
            
            ActiveWorkSheetName = ws.Name
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
            TickerRowCount = 2
            
            'Set start row to 2
            j = 2
            
            LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox ("Last row in column A is " & LastRowA) 
            
                'Loop through all rows
                For i = 2 To LastRowA
                
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    ws.Cells(TickerRowCount, 9).Value = ws.Cells(i, 1).Value
                    
                    ws.Cells(TickerRowCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
                        'Conditional formating
                        If ws.Cells(TickerRowCount, 10).Value < 0 Then
                    
                        'Set cell background color to red
                        ws.Cells(TickerRowCount, 10).Interior.ColorIndex = 3
                    
                        Else
                    
                        'Set cell background color to green
                        ws.Cells(TickerRowCount, 10).Interior.ColorIndex = 4
                    
                        End If
                        
                        'Calculate and write percent change in column K (#11)
                        If ws.Cells(j, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                        
                        'Percent formating
                        ws.Cells(TickerRowCount, 11).Value = Format(PerChange, "Percent")
                        
                        Else
                        
                        ws.Cells(TickerRowCount, 11).Value = Format(0, "Percent")
                        
                        End If
                        
                    ws.Cells(TickerRowCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                    
                    'Increase TickerRowCount by 1
                    TickerRowCount = TickerRowCount + 1
                    
                    'Set new start row of the ticker block
                    j = i + 1
                    
                    End If
                
                Next i
                
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            'MsgBox ("Last row in column I is " & LastRowI)
            
            GreatVol = ws.Cells(2, 12).Value
            GreatIncr = ws.Cells(2, 11).Value
            GreatDecr = ws.Cells(2, 11).Value
            
                'Loop for summary
                For i = 2 To LastRowI
                
                    If ws.Cells(i, 12).Value > GreatVol Then
                    GreatVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatVol = GreatVol
                    
                    End If
                    
                    If ws.Cells(i, 11).Value > GreatIncr Then
                    GreatIncr = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatIncr = GreatIncr
                    
                    End If
                    
                    If ws.Cells(i, 11).Value < GreatDecr Then
                    GreatDecr = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatDecr = GreatDecr
                    
                    End If
                    
                'Write summary results in ws.Cells
                ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
                ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
                ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
                
                Next i
                
            Worksheets(ActiveWorkSheetName).Columns("A:Z").AutoFit
                
        Next ws
            
    End Sub


