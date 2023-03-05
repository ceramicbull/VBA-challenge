Sub Stonks()
For Each ws In Worksheets
    ' Determine Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Other Starting Variables
    output_row = 2
    symbol_vol = 0
    open_value = 2
    
    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Formatting
    'Conditional Formatting for Yearly Change
    Dim Positive As FormatCondition, Negative As FormatCondition
    ws.Columns.FormatConditions.Delete
    Set Positive = ws.Columns("J:K").FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set Negative = ws.Columns("J:K").FormatConditions.Add(xlCellValue, xlLess, "=0")
    With Positive
        .Interior.ColorIndex = 4
    With Negative
        .Interior.ColorIndex = 3
    ws.Range("J1:K1").FormatConditions.Delete
    'Number Format for Percentage
    ws.Columns("K").NumberFormat = "#.##%"

    'Begin Loop
    For input_row = 2 To Last_Row
        If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1).Value Then
            'Print Ticker
            ws.Cells(output_row, 9) = ws.Cells(input_row, 1).Value
            'Final Volume Calculation
            symbol_vol = symbol_vol + ws.Cells(input_row, 7).Value
            'Print Volume
            ws.Cells(output_row, 12).Value = symbol_vol
            'Print Change
            ws.Cells(output_row, 10).Value = ws.Cells(input_row, 6).Value - ws.Cells(open_value, 3)
            'Print Percent Change
            ws.Cells(output_row, 11).Value = ws.Cells(output_row, 10).Value / ws.Cells(open_value, 3)
            'Prepare for Next Ticker & Clean-up
            output_row = output_row + 1
            symbol_vol = 0
            open_value = input_row + 1
            
        Else
        'Tabulate Volume
            symbol_vol = symbol_vol + ws.Cells(input_row, 7).Value
        End If
    Next input_row
   
    'Stock Highlights
    
    'New Last Row
        Last_Row2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    'Entries
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Find Percentages
    'Set Variables
        Dim Max_Percent As Double
        Dim Min_Percent As Double
        Max_Percent = 0
        Min_Percent = 0
        Max_Volume = 0
        Max_Stock = 0
        Min_Stock = 0
        Vol_stock = 0
    'Loop
    For i = 2 To Last_Row2
    'Determine Max/Min Percent and Ticker Row
        If ws.Cells(i, 11).Value > Max_Percent Then
            Max_Percent = ws.Cells(i, 11).Value
            Max_Stock = i
        ElseIf ws.Cells(i, 11).Value < Min_Percent Then
            Min_Percent = ws.Cells(i, 11).Value
            Min_Stock = i
        End If
    'Determine Max Volume
        If ws.Cells(i, 12).Value > Max_Volume Then
            Max_Volume = ws.Cells(i, 12).Value
            Vol_stock = i
        End If
    Next i
    'Print Values
        ws.Cells(2, 16).Value = ws.Cells(Max_Stock, 9).Value
        ws.Cells(2, 17).Value = Max_Percent
        ws.Cells(3, 16).Value = ws.Cells(Min_Stock, 9).Value
        ws.Cells(3, 17).Value = Min_Percent
        ws.Cells(4, 16).Value = ws.Cells(Vol_stock, 9).Value
        ws.Cells(4, 17).Value = Max_Volume
        ws.Range("Q2:Q3").NumberFormat = "#.##%"
    'Autofit everything
    ws.Columns("A:Q").AutoFit
    End With
    End With
Next ws

End Sub
