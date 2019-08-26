Sub DoStuff() 'Excel VBA to extract the unique items.



    Dim ws As Worksheet
    Dim VarI As Range
    Dim i, sumry, stockvol, mnday, Arow, run, start, fin, Irow, Lrow As Long
    Dim opn, cls As Double
    Dim stock As String


    
    For Each ws In Sheets
    
        'create range variables
        Arow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'write data titles
        ws.Range("I1,O1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("N2") = "Greatest Increase"
        ws.Range("N3") = "Greatest Decrease"
        ws.Range("N4") = "Greatest Volume"
        ws.Range("P1") = "Values"
        
        stockvol = 0 'total stock volume variable
        Count = 0 'helper column counter for count of each variable
        sumry = 2 'counter for new data location
        For i = 2 To Arow
            If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then
                stock = ws.Range("A" & i).Value
                ws.Range("I" & sumry).Value = stock
                ws.Range("L" & sumry).Value = stockvol + ws.Range("G" & i).Value
                ' create helper column to calculate first and last row for each ticker
                ws.Range("H" & sumry).Value = Count + 1
                stockvol = 0
                Count = 0
                sumry = sumry + 1
            Else:
                stockvol = ws.Range("G" & i).Value + stockvol
                Count = Count + 1
            End If
        Next i
        
        ColLength = sumry - 1
        run = 0
        For i = 2 To ColLength
            ' calculate first row of ticker
            start = 2 + run
            'calculate last row of ticker
            fin = start + ws.Range("H" & i).Value - 1
            
            opn = ws.Range("C" & start).Value 'Opening stock value
            cls = ws.Range("F" & fin).Value 'closeing stock value
            ws.Range("J" & i) = cls - opn
            
            'if opening value is 0 set %change to 0 to avoid 0 division error
            If opn = 0 Then
                ws.Range("K" & i) = 0
            Else
                ws.Range("K" & i) = (cls - opn) / opn
            End If
            
            run = run + ws.Range("H" & i).Value
            
            
            'conditional formating for positive or negative difference
            If ws.Range("J" & i) > 0 Then
                ws.Range("J" & i).Interior.Color = vbGreen
            ElseIf ws.Range("J" & i) < 0 Then
                ws.Range("J" & i).Interior.Color = vbRed
            End If
        Next i
        
        'get Max Volume
        ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L2", ws.Range("L" & Rows.Count).End(xlUp)))
        
        'get greatest increase
        ws.Range("p2") = Application.WorksheetFunction.Max(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))
        
        'get greatest decrease
        ws.Range("p3") = Application.WorksheetFunction.Min(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))
        
        For i = 2 To ColLength
            ' put ticker symbols next to summary data
            If ws.Range("L" & i) = ws.Range("P4") Then
                ws.Range("O4") = Range("I" & i)
            ElseIf ws.Range("K" & i) = ws.Range("P3") Then
                ws.Range("O3") = Range("I" & i)
            ElseIf ws.Range("K" & i) = ws.Range("P2") Then
                ws.Range("O2") = ws.Range("I" & i)
            End If
        Next i
        
        ' format columns/cells
        
        ws.Columns("J:P").NumberFormat = "#,#0.0#" 'set numbers to include comma if over 1,000 and at min show value as 0.0
        ws.Range("P2:P3,K2:K" & ColLength).NumberFormat = "0.00%" ' set select cells/columns to percent
        ws.Columns("i:p").AutoFit ' autofit columns
        ws.Range("H:H").Value = "" 'remove values from helper column
    Next ws

End Sub
