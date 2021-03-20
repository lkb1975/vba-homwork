Sub Financial_Analysis()
    For Each ws In Worksheets
    
    'Declaring the values
    Dim T As Double
    Dim OV As Double
    Dim CV As Double
    Dim TV As Double
    Dim end1 As Double
    
    ws.Range("J:j").NumberFormat = "0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    
    
    OV = ws.Cells(2, 3).Value
    TV = 0
     
    'Setting up the labels
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly_Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock"
    end1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    T = 2
    For i = 2 To end1

    'Calculating all values
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     ws.Cells(T, 9).Value = ws.Cells(i, 1).Value
     CV = ws.Cells(i, 6).Value
     ws.Cells(T, 10).Value = CV - OV
     If OV = 0 Then
     ws.Cells(T, 11).Value = 0
     Else
     ws.Cells(T, 11).Value = (CV - OV) / OV
     End If
     
     TV = TV + ws.Cells(i, 7).Value
     ws.Cells(T, 12).Value = TV
     
     If ws.Cells(T, 10).Value > 0 Then
     ws.Cells(T, 10).Interior.ColorIndex = 4
     Else
     ws.Cells(T, 10).Interior.ColorIndex = 3
     End If
     
     
     TV = 0
     
     OV = ws.Cells(i + 1, 3).Value
     
     T = T + 1
Else
     TV = TV + ws.Cells(i, 7).Value

End If
     Next i
        Next ws
        
        
 End Sub
