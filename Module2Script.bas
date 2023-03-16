Attribute VB_Name = "Module1"

Sub Yearlystockdata_Macro()
Attribute Yearlystockdata_Macro.VB_Description = "Module #2 Challenge"
Attribute Yearlystockdata_Macro.VB_ProcData.VB_Invoke_Func = " \n14"
    
    For Each ws In Worksheets
    
    Dim WorksheetName As String
    WorksheetName = ws.Name
    
    Dim i As Long
    Dim j As Long
    Dim TickCount As Long
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim PerChange As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    
    'Column Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Set Ticker Counter
    TickCount = 2
    
    'Start Row to 2
    j = 2
    
   LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'Time to loop all rows
    For i = 2 To LastRowA
    
    'create the if statement for the ticker name
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Ticker for Column 1
        ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
    
        'Calculating Yearly Change in J column
        ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
            'Adding in Conditional Formating
            If ws.Cells(TickCount, 10).Value < 0 Then
            
            'Set Color to Red
            ws.Cells(TickCount, 10).Interior.ColorIndex = 3
            
            Else
            
            'Setting Green Color
            ws.Cells(TickCount, 10).Interior.ColorIndex = 4
            
            End If
        'Calculate Percent Change
            If ws.Cells(j, 3).Value <> 0 Then
            PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            
        'Format for Percent
            ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
            
            Else
            
            ws.Cells(TickCount, 11).Value = Format(0, "Percent")
            
            End If
            
        
        'Calculate total Volume
        ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        TickCount = TickCount + 1
        'adding tickcount by 1
        'new row for ticker block
        j = i + 1
        
        End If
        
    Next i
    
   LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

        
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
        'Looping summary
            For i = 2 To LastRowI
        'Creating if for next value if larger. If yes, new value for all cells (Volume)
            If ws.Cells(i, 12).Value > GreatVol Then
            GreatVol = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
                Else
                GreatVol = GreatVol
            End If
        'If next value is greater/larger, new value for cells (Increase)
            If ws.Cells(i, 11).Value > GreatIncr Then
            GreatIncr = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                GreatIncr = GreatIncr
            End If
        
        'If next value is smaller, new value for cells (Decrease)
            If ws.Cells(i, 11).Value < GreatDecr Then
            GreatDecr = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
                Else
                GreatDecr = GreatDecr
            End If
        'Summary Conclusion
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
        'Format column
            Worksheets(WorksheetName).Columns("A:Z").AutoFit
        Next ws
        
    End Sub
