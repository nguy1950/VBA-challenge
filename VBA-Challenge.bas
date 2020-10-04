Attribute VB_Name = "Module1"
Sub WorksheetsLoop()
    Dim current_sheet As Worksheet
    
    For Each current_sheet In Worksheets
        'set headers for summary table
        current_sheet.Cells(1, 9).Value = "Ticker"
        current_sheet.Cells(1, 10).Value = "Yearly Change"
        current_sheet.Cells(1, 11).Value = "Percent Change"
        current_sheet.Cells(1, 12).Value = "Total Stock Volume"
        'hard solution table
        'current_sheet.Cells(2, 15).Value = "Greatest % Increase"
        'current_sheet.Cells(3, 15).Value = "Greatest % Decrease"
        'current_sheet.Cells(4, 15).Value = "Greatest Total Volume"
        'current_sheet.Cells(1, 16).Value = "Ticker"
        'current_sheet.Cells(1, 16).Value = "Value"
        
'------------------
        Dim ticker As String 'set an initial variable for ticker
        ticker = " "
    
        Dim total_vol As Double 'set an initial variable for total stock volume
        total_vol = 0
    
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim delta_price As Double   'set an initial variable for yearly change
        delta_price = 0
        Dim delta_percent As Double     'set an initial variable for percent change
        delta_percent = 0
    
    
        'track each ticker location via summary table
        Dim summary_table_row As Long
        summary_table_row = 2
        
        'calculate # of iterations
        Dim lastrow As Long
        Dim i As Long
        lastrow = current_sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set value for open / close for the first ticker
        
        open_price = current_sheet.Cells(2, 3).Value
        'close_price = current_sheet.Cells(i, 6).Value
        'MsgBox (open_price)
        
        'loop until the last row
        'For i = 2 To 20
        For i = 2 To lastrow
        
            'loop until ticker changes
            If current_sheet.Cells(i + 1, 1).Value <> current_sheet.Cells(i, 1).Value Then
                ticker = current_sheet.Cells(i, 1).Value
                
                'calculate delta price/percent
                close_price = current_sheet.Cells(i, 6).Value
                delta_price = close_price - open_price
                If open_price <> 0 Then
                    delta_percent = 100 * (delta_price / open_price)
                End If
                
                '+total volume counter
                total_volume = total_volume + current_sheet.Cells(i, 7).Value
              
                'print to the summary table
                current_sheet.Range("I" & summary_table_row).Value = ticker
                current_sheet.Range("J" & summary_table_row).Value = delta_price
                'yearly change color calculation, GREEN & RED
                If (delta_price > 0) Then
                    current_sheet.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf (delta_price <= 0) Then
                    current_sheet.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                
                'print to the summary table
                current_sheet.Range("K" & summary_table_row).Value = (CStr(delta_percent) & "%")
                current_sheet.Range("L" & summary_table_row).Value = total_vol
'-----------------
                '+summary table row counter
                summary_table_row = summary_table_row + 1
                'reset counters
                delta_price = 0
                close_price = 0
                delta_percent = 0
                total_vol = 0
                open_price = current_sheet.Cells(i + 1, 3).Value
                
            Else
                total_vol = total_vol + current_sheet.Cells(i, 7).Value
            End If
        Next i
    Next current_sheet
End Sub

