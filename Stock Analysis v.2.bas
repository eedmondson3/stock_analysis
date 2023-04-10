Attribute VB_Name = "Module11"
Sub stock_analysis():

For Each ws In Worksheets
    ' Variables
        Dim tvolume As LongLong
        
        tvolume = 0
        summary_row = 2
        opener = Range("C2")
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'Column Creation & Format
        ws.Range("J1") = "Ticker Symbol"
        ws.Range("K1") = "Yearly Change"
        ws.Range("L1") = "Percent Change"
        ws.Range("M1") = "Total Stock Volume"
        ws.Columns("L").NumberFormat = "#.##%"
        ws.Columns("K").NumberFormat = "$#.##"

        'Ticker Return & Volume Loop
            For i = 2 To LastRow
            
                'Ticker Name Detector
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                        'Ticker Name
                        ticker = ws.Cells(i, 1).Value
                        
                        'Ticker Total Volume
                        tvolume = tvolume + ws.Cells(i, 7).Value
                        
                        'Print Total Volume
                         ws.Cells(summary_row, 13) = tvolume
                        
                        'Print Ticker Name
                         ws.Cells(summary_row, 10) = ticker
                        
                        'Reset Volume Counter
                         tvolume = 0
                        
                        'Calculate & Print Yearly Change as ydelta
                        ydelta = ws.Cells(i, 6) - opener
                        ws.Cells(summary_row, 11) = ydelta
                        
                        'Calculate & Print Opener to Closer % Change
                        ypdelta = (ydelta / opener)
                        ws.Cells(summary_row, 12) = ypdelta
                    
                        '+/- Color Change
                            If ws.Cells(summary_row, 11) >= 0.01 Then
                                ws.Cells(summary_row, 11).Interior.ColorIndex = 4
                                
                            ElseIf ws.Cells(summary_row, 11) = 0 Then
                                ws.Cells(summary_row, 11).Interior.ColorIndex = 6
                            
                            ElseIf ws.Cells(summary_row, 11) <= 0.99 Then
                                ws.Cells(summary_row, 11).Interior.ColorIndex = 3
                            
                            End If
                    
                    'Move summary row for each ticker
                        summary_row = summary_row + 1
                    
                    'Reset opener for each new ticker
                        opener = ws.Cells(i + 1, 3).Value
                    
                      
                    Else
                
                    'Continuation for same name ticker
                        tvolume = tvolume + ws.Cells(i, 7).Value
                    
                    End If
            
            Next i
    
'Greatest Increase, Decrease, & Total Volume
    
        'Variables
            LastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
            
        'Column and Row Formatting
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            ws.Range("Q2", "Q3").NumberFormat = "#.##%"
        
        
        'Greatest Increase, Decrease, & Total Volume Values Loop
            For j = 2 To LastRow
                Max = WorksheetFunction.Max(ws.Range("l2:l" & j))
                Min = WorksheetFunction.Min(ws.Range("l2:l" & j))
                MaxV = WorksheetFunction.Max(ws.Range("M2:l" & j))
                    
                'Value Print Functions
                    ws.Range("q2") = Max
                    ws.Range("q3") = Min
                    ws.Range("q4") = MaxV
      
            Next j
    
        'Ticker Match & Print Loop
            For k = 2 To LastRow
                If ws.Cells(k, 12) = Max Then
                ws.Range("p2") = ws.Cells(k, 10)
            
                ElseIf ws.Cells(k, 12) = Min Then
                ws.Range("p3") = ws.Cells(k, 10)
                    
                ElseIf ws.Cells(k, 13) = MaxV Then
                ws.Range("p4") = ws.Cells(k, 10)
            
                End If
       
            Next k
      
        'AutoFit Columns
        ws.Columns("J:Q").AutoFit
    
Next ws
    
End Sub

