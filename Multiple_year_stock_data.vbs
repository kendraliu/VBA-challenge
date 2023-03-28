Sub stock()

For Each x In Worksheets
    Dim j As Integer
    j = 2
    volume = 0
    x.Cells(1, 9).Value = "Ticker"
    x.Cells(1, 10).Value = "Yearly change"
    x.Cells(1, 11).Value = "Percent change"
    x.Cells(1, 12).Value = "Total stock volume"
    x.Cells(2, 14).Value = "Greatest % increase"
    x.Cells(3, 14).Value = "Greatest % decrease"
    x.Cells(4, 14).Value = "Greatest total volume"
    x.Cells(1, 15).Value = "Ticker"
    x.Cells(1, 16).Value = "Value"
    
    lastrow = x.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
           If x.Cells(i, 1).Value <> x.Cells(i - 1, 1).Value Then
           
                x.Cells(j, 9).Value = x.Cells(i, 1).Value
                
                If x.Cells(i - 1, 6).Value <> "<close>" Then
                    closingprice = x.Cells(i - 1, 6).Value
                    x.Cells(j - 1, 10).Value = closingprice - openingprice
                    x.Cells(j - 1, 11).Value = x.Cells(j - 1, 10).Value / openingprice
                    x.Cells(j - 1, 12).Value = volume
                End If
                
                openingprice = x.Cells(i, 3).Value
                volume = 0
                j = j + 1
                
           End If
           
           volume = volume + x.Cells(i, 7).Value
             
    Next i
    
    x.Cells(2, 16).Value = x.Cells(2, 11).Value
    x.Cells(2, 15).Value = x.Cells(2, 9).Value
    x.Cells(3, 16).Value = x.Cells(2, 11).Value
    x.Cells(3, 15).Value = x.Cells(2, 9).Value
    For i = 2 To lastrow
        If x.Cells(i + 1, 11).Value > x.Cells(2, 16).Value Then
            x.Cells(2, 16).Value = x.Cells(i + 1, 11).Value
            x.Cells(2, 15).Value = x.Cells(i + 1, 9).Value
        ElseIf x.Cells(i + 1, 11).Value < x.Cells(3, 16).Value Then
            x.Cells(3, 16).Value = x.Cells(i + 1, 11).Value
            x.Cells(3, 15).Value = x.Cells(i + 1, 9).Value
        End If
    Next i
    
    x.Cells(4, 16).Value = x.Cells(2, 12).Value
    x.Cells(4, 15).Value = x.Cells(2, 9).Value
    For i = 2 To lastrow
        If x.Cells(i + 1, 12).Value > x.Cells(4, 16).Value Then
            x.Cells(4, 16).Value = x.Cells(i + 1, 12).Value
            x.Cells(4, 15).Value = x.Cells(i + 1, 9).Value
        End If
    Next i
    
Next x
End Sub
