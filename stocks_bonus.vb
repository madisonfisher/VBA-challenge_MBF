Sub stocks_bonus():

'to run on the whole workbook
Dim Current As Worksheet
For Each Current In Worksheets:

Dim value1 As Double
Dim value2 As Double

 LRow2 = Current.Cells(Rows.count, 10).End(xlUp).Row
    Current.Cells(1, 17).Value = "Ticker"
    Current.Cells(1, 18).Value = "Value"
    Current.Cells(2, 16).Value = "Greatest % Increase"
    Current.Cells(3, 16).Value = "Greatest % Decrease"
    Current.Cells(4, 16).Value = "Greatest Total Volume"
    Current.Range("R2:R3").NumberFormat = "0.00%"
    
    'greatest increase
    Current.Cells(2, 18).Value = Current.Cells(2, 12).Value
    Current.Cells(2, 17).Value = Current.Cells(2, 10).Value
    For i = 2 To LRow2
    value1 = Current.Cells(2, 18).Value
    value2 = Current.Cells(i, 12).Value
        If value1 < value2 Then
            Current.Cells(2, 18).Value = Current.Cells(i, 12).Value
            Current.Cells(2, 17).Value = Current.Cells(i, 10).Value
        End If
      
    Next i
    
    'greatest decrease
    Current.Cells(3, 18).Value = Current.Cells(2, 12).Value
    Current.Cells(3, 17).Value = Current.Cells(2, 10).Value
    For i = 2 To LRow2
    value1 = Current.Cells(3, 18).Value
    value2 = Current.Cells(i, 12).Value
        If value1 > value2 Then
            Current.Cells(3, 18).Value = Current.Cells(i, 12).Value
            Current.Cells(3, 17).Value = Current.Cells(i, 10).Value
        End If
      
    Next i
    
    'greatest total volume
    Current.Cells(4, 18).Value = Current.Cells(2, 13).Value
    Current.Cells(4, 17).Value = Current.Cells(2, 10).Value
    For i = 2 To LRow2
    value1 = Current.Cells(4, 18).Value
    value2 = Current.Cells(i, 13).Value
        If value1 < value2 Then
            Current.Cells(4, 18).Value = Current.Cells(i, 13).Value
            Current.Cells(4, 17).Value = Current.Cells(i, 10).Value
        End If
      
    Next i
Next
End Sub