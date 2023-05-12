Attribute VB_Name = "Module1"
Sub A_test()
Attribute A_test.VB_ProcData.VB_Invoke_Func = " \n14"

' A_test Macro
For Each ws In Worksheets:
        Dim ticker As String
        Dim total_volume As Variant
        Dim open_price, close_price, yearly_change, percent_change As Double
        Dim summary_row As Integer
        Dim last_row As Long

        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        total_volume = 0
        summary_row = 1
        
        ws.Range("I" & summary_row) = "Ticker"
        ws.Range("J" & summary_row) = "Total Stock Volume"
        ws.Range("K" & summary_row) = "Yearly Change"
        ws.Range("L" & summary_row) = "Percent Change"
        
        
        summary_row = 2
        open_price = ws.Cells(2, 3).Value
        
      'Calculates each ticker symbol's change and volume
        For i = 2 To last_row
            ticker = ws.Cells(i, 1).Value
            
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            If (ticker <> ws.Cells(i + 1, 1).Value) Then
                
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                percent_change = yearly_change / open_price
                
                ws.Range("I" & summary_row) = ticker
                ws.Range("J" & summary_row) = total_volume
                ws.Range("K" & summary_row) = yearly_change
                ws.Range("L" & summary_row) = FormatPercent(percent_change)
                
       'Conditional Formating for percentage change lower or greater
                If (yearly_change <= 0) Then
                    ws.Cells(summary_row, 11).Interior.ColorIndex = 3
                Else: ws.Range("K" & summary_row).Interior.ColorIndex = 4
                End If
                
                open_price = ws.Cells(i + 1, 3).Value
                total_volume = 0
                summary_row = summary_row + 1
        End If
        
        Next i
        
    'Summary table 2
        Cells(1, 17) = "Ticker"
        Cells(1, 18) = "Value"
        Cells(2, 16) = "Greatest % increase"
        Cells(3, 16) = "Greatest % Decrease"
        Cells(4, 16) = "Greatest Total Volume"
        

     'Calculates best performer stock with ticker symbol
        Dim max_ticker As String
        Dim max_value As Double
        max_value = ws.Cells(2, 12).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 12).Value > max_value) Then
            max_ticker = ws.Cells(i, 9).Value
            max_value = ws.Cells(i, 12).Value
            End If
        Next i
        Cells(2, 17).Value = max_ticker
        Cells(2, 18).Value = FormatPercent(max_value)
        
     'Calculates the worst performer stock with ticker symbol
        Dim min_ticker As String
        Dim min_value As Double
        min_value = ws.Cells(2, 12).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 12).Value < min_value) Then
            min_ticker = ws.Cells(i, 9).Value
            min_value = ws.Cells(i, 12).Value
            End If
        Next i
        Cells(3, 17).Value = min_ticker
        Cells(3, 18).Value = FormatPercent(min_value)
        
     'Calculates maximum volume stock with ticker symbol
        Dim max_volume_ticker As String
        Dim max_volume_value As Double
        max_volume_value = ws.Cells(2, 10).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 10).Value > max_volume_value) Then
            max_volume_ticker = ws.Cells(i, 9).Value
            max_volume_value = ws.Cells(i, 10).Value
            End If
        Next i
        Cells(4, 17).Value = max_volume_ticker
        Cells(4, 18).Value = max_volume_value


Next ws

End Sub