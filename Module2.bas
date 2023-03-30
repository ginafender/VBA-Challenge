Attribute VB_Name = "Module1"
Sub stocksfixed()

'Loops through sheets.
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    'add column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Year Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'declare variables
    Dim ticker As String
    Dim opener As Double
    Dim closer As Double
    Dim yearchange As Double
    Dim nextticker As String
    Dim previousticker As String
    
    'find the last row
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'create volume counter
    Dim volume As Double
    volume = 0
    
    'create output area variable to determine what row the output goes into
    Dim newrow As Integer
    newrow = 2
    
        'loop through all rows in the sheet
        For I = 2 To lastRow
            ticker = ws.Cells(I, 1).Value
            nextticker = ws.Cells(I + 1, 1).Value
            previousticker = ws.Cells(I - 1, 1).Value
            volume = volume + ws.Cells(I, 7).Value
            
                'capture opening price on first day of year
                If ticker <> previousticker Then
                    opener = ws.Cells(I, 3).Value
                End If
                'capture closing price on last day of year
                If ticker <> nextticker Then
                    closer = ws.Cells(I, 6).Value
                'calculate closing price
                    yearchange = closer - opener
                'record ticker into new column
                    ws.Cells(newrow, 9).Value = ticker
                'record year change into new row
                    ws.Cells(newrow, 10).Value = yearchange
                 'set color for year change cells
                        If yearchange > 0 Then
                            ws.Cells(newrow, 10).Interior.ColorIndex = 4
                        ElseIf yearchange < 0 Then
                            ws.Cells(newrow, 10).Interior.ColorIndex = 3
                        End If
                'calculate percentage change and record to new column
                    pchange = (yearchange / opener)
                    pchange = FormatPercent(pchange)
                    ws.Cells(newrow, 11).Value = pchange
                'output volume totals into new column
                    ws.Cells(newrow, 12).Value = volume
                'start new row for new ticker
                    newrow = newrow + 1
                'clear volume counter after ticker changes
                    volume = 0
                End If
        
    Next I
    
    'add functionality to find greatest increase, decrease, volume
    'declare variables
     Dim greatInc As Integer
     Dim greatDec As Integer
     Dim greatVol As Integer
     
            'find the biggest increase
            ws.Range("P2") = (Application.WorksheetFunction.Max(ws.Range("K2", ws.Range("K2").End(xlDown)).Rows))
            ws.Range("P2") = FormatPercent(ws.Range("P2"))
            'match the ticker value to the biggest increase
            greatInc = WorksheetFunction.Match(ws.Range("P2").Value, (ws.Range("K2", ws.Range("K2").End(xlDown))), 0) + 1
            ws.Range("O2").Value = ws.Cells(greatInc, 9)
            
            
            'find the biggest decrease
            ws.Range("P3") = Application.WorksheetFunction.Min(ws.Range("K2", ws.Range("K2").End(xlDown)).Rows)
            ws.Range("P3") = FormatPercent(ws.Range("P3"))
            'match the ticker value to the biggest decrease
            greatDec = WorksheetFunction.Match(ws.Range("P3").Value, (ws.Range("K2", ws.Range("K2").End(xlDown))), 0) + 1
            ws.Range("O3").Value = ws.Cells(greatDec, 9)
            
            'find the max volume
            ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L2", ws.Range("L2").End(xlDown)).Rows)
            'match the ticker value to the max volume
            greatDec = WorksheetFunction.Match(ws.Range("P4").Value, (ws.Range("L2", ws.Range("L2").End(xlDown))), 0) + 1
            ws.Range("O4").Value = ws.Cells(greatDec, 9)
        
        
    
Next ws

End Sub



