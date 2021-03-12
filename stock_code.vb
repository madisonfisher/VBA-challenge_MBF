Sub stocks():

'to run on the whole workbook
Dim Current As Worksheet
For Each Current In Worksheets:

    'define variables
    Dim ticker1 As String
    Dim ticker2 As String
    Dim i As Long
    Dim LRow As Long
    Dim x As Long
    Dim volume As Double
    Dim LRow2 As Double
    Dim opened As Double
    Dim closed As Double
    Dim count As Long
    Dim changed As Double
    Dim percent As Double
    Dim value1 As Double
    Dim value2 As Double
    
    LRow = Current.Cells(Rows.count, 1).End(xlUp).Row
    
    'fill in headers in every sheet
    Current.Cells(1, 10).Value = "Ticker"
    Current.Cells(1, 11).Value = "Yearly Change"
    Current.Cells(1, 12).Value = "Percent Change"
    Current.Cells(1, 13).Value = "Total Stock Volume"
    
    'set value of print out row
    x = 2
    
    For i = 2 To LRow
    ticker1 = Current.Cells(i, 1).Value
    ticker2 = Current.Cells(i + 1, 1).Value
        'compare ticker value
        If ticker1 <> ticker2 Then
            Current.Cells(x, 10).Value = ticker1
            'add in last volume
            volume = Current.Cells(i, 7).Value + volume
            Current.Cells(x, 13).Value = volume
            'value of final closed
            closed = Current.Cells(i, 6).Value
            'value of first open
            opened = Current.Cells(i - count, 3).Value
            changed = closed - opened
                'formatting the color based on changed value
                If changed > 0 Then
                    Current.Cells(x, 11).Interior.ColorIndex = 4
                ElseIf changed < 0 Then
                    Current.Cells(x, 11).Interior.ColorIndex = 3
                End If
            'defining and printing percent change
            Current.Cells(x, 11).NumberFormat = "###0.00"
            Current.Cells(x, 12).NumberFormat = "0.00%"
            'defining and printing change
            Current.Cells(x, 11).Value = changed
            If opened <> 0 Then
                percent = changed / opened
                Current.Cells(x, 12).Value = percent
            'for any tickers that open at 0
            Else
                Current.Cells(x, 12).Value = 0
            End If
            x = x + 1
            volume = 0
            count = 0
        Else
            'adding total volume when ticker is the same
            volume = Current.Cells(i, 7).Value + volume
            'keeping track of how many rows for yearly change
            count = 1 + count
        End If
    Next i
    
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