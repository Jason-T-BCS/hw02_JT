Attribute VB_Name = "Module1"
Sub Moderate_Hard():


        For Each ws In Worksheets
            Dim ticker As String
            Dim total_vol As Double
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            Dim summary_row As Integer
            Dim year_open As Double
            Dim year_close As Double
            Dim year_change As Double
            total_vol = 0
            summary_row = 2
            ws.Cells(1, 9).Value = ws.Cells(1, 1).Value
            ws.Cells(1, 16).Value = ws.Cells(1, 1).Value
            
            ws.Cells(1, 10).Value = ws.Cells(1, 7).Value
            ws.Cells(1, 17).Value = ws.Cells(1, 7).Value
            
            ws.Cells(1, 11).Value = "Yearly_Change"
            ws.Cells(1, 12).Value = "Percent Change"
            
                
                For i = 2 To lastrow
                    ' Take year open value
                    year_open = ws.Cells(i, 3).Value
                    
                    'calculate volume total for every ticker
                    total_vol = total_vol + ws.Cells(i, 7).Value
                    
                    '   Display the volume total
                    ws.Range("j" & summary_row).Value = total_vol
                    
                        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        ticker = ws.Cells(i, 1).Value
                        ws.Range("i" & summary_row).Value = ticker
                        
                        '   Take closing year value
                        year_close = ws.Cells(i, 6).Value
                        
                        '   Calculate diff
                        year_change = year_close - year_open
                        
                        '   Display yearly change values in a column
                        ws.Range("k" & summary_row).Value = year_change
                        
                        '   Calculate percent change
                        percent_change = (year_close - year_open) / year_open
                        
                        '   Display percent change values in a column
                        ws.Range("l" & summary_row).Value = percent_change
                        
                        'jump to next summary row
                        summary_row = summary_row + 1
                        
                        '   Reset total_vol
                        total_vol = 0
                        
                        End If
                    
                      '   Color cells
                        If ws.Range("k" & i).Value >= 0 Then
                        ws.Range("k" & i).Interior.ColorIndex = 4
                        Else: ws.Range("k" & i).Interior.ColorIndex = 3
                        End If
                        
                    Next i
                    
' --------------------------------------------------------
' Hard Solution

                '   Find Max
                   ' If ws.Cells(i, 12) > Max Then
                        'Max = ws.Cells(i, 12)
                        'Tick_Max = ws.Cells(i, 9)
                        'ws.Cells(2, 16) = Tick_Max
                        ''ws.Cells(2, 17) = Max
                    'End If
                '   Find Min
                    'If ws.Cells(i, 12) < Min Then
                        'Min = ws.Cells(i, 12)
                        'Tick_Min = ws.Cells(i, 9)
                        'ws.Cells(3, 16) = Tick_Min
                        'ws.Cells(3, 17) = Min
                    'End If
                '   Find Max Vol
                    'If ws.Cells(i, 10) > Max Then
                        'Max_Vol = ws.Cells(i, 17)
                        'Tick_Max_Vol = ws.Cells(4, 16)
                    'End If
' --------------------------------------------------------
        
                '   Autofit cells
                    ws.Range("i:l").HorizontalAlignment = xlCenter
                    ws.Range("i:l").EntireColumn.AutoFit
                    
                '   Format column number type
                    ws.Range("l" & i).NumberFormat = "0.00%"
                    
            Next ws
            
            MsgBox ("All Sheets Have Been Processed!")
            
End Sub





