Sub DoStuff() 'Excel VBA to extract the unique items.



    Dim ws As Worksheet
    Dim rng, crit, sumcrit, datecrit, opnrng, clsrng, VarI As Range
    Dim i, j, sumry, stockvol, mxday, mnday, mxvol, mxchng, mnchng, Arow, lrow As Long
    Dim Irow As Integer
    Dim opn, cls As Double
    Dim stock As String


    
    For Each ws In Sheets
        
        Arow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'set ranges to variables
        Set crit = ws.Range("A2", ws.Range("A" & Rows.Count).End(xlUp))
        Set sumcrit = ws.Range("G2", ws.Range("G" & Rows.Count).End(xlUp))
        Set datecrit = ws.Range("B2", ws.Range("B" & Rows.Count).End(xlUp))
        Set opnrng = ws.Range("C2", ws.Range("C" & Rows.Count).End(xlUp))
        Set clsrng = ws.Range("F2", ws.Range("F" & Rows.Count).End(xlUp))
       
        
        'write data titles
        ws.Range("I1,O1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"

        
        sumry = 2
        stockvol = 0
        For i = 2 To Arow
            If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then
                stock = ws.Range("A" & i).Value
                ws.Range("I" & sumry).Value = stock
                ws.Range("L" & sumry).Value = stockvol + ws.Range("G" & i).Value
                sumry = sumry + 1
                stockvol = 0
            Else:
                stockvol = ws.Range("G" & i).Value + stockvol
            End If
        Next i
        
       Irow = ws.Cells(Rows.Count, 9).End(xlUp).Row
       For i = 1 To Irow
        For j = 2 To Arow
       Set VarI = ws.Range("I" & i + 1)
        ' find first and last day and put in variables
         mxday = Application.WorksheetFunction.MaxIfs(datecrit, crit, ws.Range("I" & i + 1))
         mnday = Application.WorksheetFunction.MinIfs(datecrit, crit, ws.Range("I" & i + 1))
        
        'put closing and opening prices into variables
         If VarI = ws.Range("A" & j) And mxday = ws.Range("B" & j) Then
             cls = ws.Range("C" & i + 1).Value
         End If
         If ws.Range("A" & j) = VarI And ws.Range("B" & j) = mnday Then
             opn = ws.Range("F" & i + 1).Value
         End If
        'calculate difference and % change for each and place in columns J and K
         ws.Range("J" & i + 1) = cls - opn
         ws.Range("K" & i + 1) = (cls - opn) / opn
         Next j
         ws.Range("J1") = "Yearly Change"
         ws.Range("K1") = "Percent Change"

        
         'conditional formating for positive or negative difference
         If ws.Range("J" & i + 1) > 0 Then
             ws.Range("J" & i + 1).Interior.Color = vbGreen
         ElseIf ws.Range("J" & i + 1) < 0 Then
             ws.Range("J" & i + 1).Interior.Color = vbRed
         End If

         ws.Range("N2") = "Greatest Increase"
         ws.Range("N3") = "Greatest Decrease"
         ws.Range("N4") = "Greatest Volume"
         ws.Range("P1") = "Values"
         'get Max Volume
         ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L2", ws.Range("L" & Rows.Count).End(xlUp)))

         'get greatest increase
         ws.Range("p2") = Application.WorksheetFunction.Max(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))

         'get greatest decrease
         ws.Range("p3") = Application.WorksheetFunction.Min(ws.Range("k2", ws.Range("k" & Rows.Count).End(xlUp)))

        ' put ticker symbols next to summary data
         If ws.Range("L" & i + 1) = ws.Range("P4") Then
            ws.Range("O4") = Range("I" & i)
         ElseIf ws.Range("K" & i + 1) = ws.Range("P3") Then
            ws.Range("O3") = Range("I" & i)
         ElseIf ws.Range("K" & i + 1) = ws.Range("P2") Then
            ws.Range("O2") = ws.Range("I" & i)
         End If
        Next i

        ' format columns/cells

        ws.Columns("J:P").NumberFormat = "#,#0.0#" 'set numbers to include comma if over 1,000 and at min show value as 0.0
        ws.Range("P2:P3,K2:K" & lrow).NumberFormat = "0.00%" ' set select cells/columns to percent
        ws.Columns("i:p").AutoFit ' autofit columns
        
    Next ws
    
  End Sub
