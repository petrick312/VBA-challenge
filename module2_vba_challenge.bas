Attribute VB_Name = "Module1"
Sub yearlystockdata():

    For Each ws In Worksheets
        
        Dim wsname As String            'worksheet name
        Dim j As Long                   'start row of ticker block
        Dim tickercount As Long         'index counter for ticker row
        Dim lastrowa As Long            'last row of column a
        Dim lastrowi As Long            'last row of column i
        Dim percentchange As Double     'percent change calculation
        Dim greatestinc As Double       'greatest increase calculation
        Dim greatestdec As Double       'greatest decrease calculation
        Dim greatestvol As Double       'greatest total volume
        
        wsname = ws.Name                'grabs worksheet name
        
        'column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        tickercount = 2                 'ticker counter starts at 2nd row
        j = 2                           'starts at 2nd row
        
        'look for last cell containing a value in column a
        lastrowa = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'loops through all rows
            For i = 2 To lastrowa
            
                'check until ticker name changes
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'display ticker name in 9th column (column I)
                    ws.Cells(tickercount, 9).Value = ws.Cells(i, 1).Value
                
                    'calc & display yearly change in 10th column (column J)
                    ws.Cells(tickercount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'format 10th column based on positive or negative % - check if change is negative number
                    If ws.Cells(tickercount, 10).Value < 0 Then
                
                        'if cell is negative, make it red
                        ws.Cells(tickercount, 10).Interior.ColorIndex = 3
                
                    Else
                
                        'if not negative, we assume positive then make it green
                        ws.Cells(tickercount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'calc and display percent change in 11th column (column K)
                    If ws.Cells(j, 3).Value <> 0 Then
                
                        percentchange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                        'adjust formating to display as a percent
                        ws.Cells(tickercount, 11).Value = Format(percentchange, "Percent")
                    
                    Else
                    
                        ws.Cells(tickercount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                    'calc and display total volume 12th column (column L)
                    ws.Cells(tickercount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                    'tickercount increased by 1
                    tickercount = tickercount + 1
                
                    'set the next start row of the ticker block
                    j = i + 1
                
                End If
            
            Next i
            
        'look for last cell containing a value in column I
        lastrowi = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
        'set starting postion for summary
        greatestinc = ws.Cells(2, 11).Value
        greatestdec = ws.Cells(2, 11).Value
        greatestvol = ws.Cells(2, 12).Value
        
            'loop for the summary
            For i = 2 To lastrowi
            
                'find greatest total volume by looking if next value is larger
                If ws.Cells(i, 12).Value > greatestvol Then
                    
                    'if so keep the larger value
                    greatestvol = ws.Cells(i, 12).Value
                    
                    'copy value in display cell
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                    
                    'if not larger, value remains the same
                    greatestvol = greatestvol
                
                End If
                
                'find greatest increase by looking if next value is larger
                If ws.Cells(i, 11).Value > greatestinc Then
                
                    'if so keep the larger value
                    greatestinc = ws.Cells(i, 11).Value
                    
                    'copy value in display cell
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                    'if not larger, value remains the same
                    greatestinc = greatestinc
                
                End If
                
                'find greatest decrease by looking if next value is smaller
                If ws.Cells(i, 11).Value < greatestdec Then
                    
                    'if so keep the smaller value
                    greatestdec = ws.Cells(i, 11).Value
                
                    'copy value in display cell
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                    'if not smaller, value remains the same
                    greatestdec = greatestdec
                
                End If
                
            'format summary results in display cells
            ws.Cells(2, 17).Value = Format(greatestinc, "Percent")
            ws.Cells(3, 17).Value = Format(greatestdec, "Percent")
            ws.Cells(4, 17).Value = Format(greatestvol, "Scientific")
            
            Next i
            
        'change formatting of column width to autofit
        Worksheets(wsname).Columns("A:Q").AutoFit
            
    'run code on next worksheet
    Next ws
        
End Sub

