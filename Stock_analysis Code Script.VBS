Attribute VB_Name = "Module1"
Sub Stock_analysis1()
Attribute Stock_analysis1.VB_ProcData.VB_Invoke_Func = " \n14"

' Stock_analysis1 Macro
For Each ws In Worksheets:
    'Define all variables

        Dim ticker, max_ticker, min_ticker, max_volume_ticker  As String
        Dim total_volume As Variant
        Dim open_price, close_price, yearly_change, percent_change, max_value, min_value, max_volume_value As Double
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
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        

     'Calculates best performer stock with ticker symbol
        
        max_value = ws.Cells(2, 12).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 12).Value > max_value) Then
            max_ticker = ws.Cells(i, 9).Value
            max_value = ws.Cells(i, 12).Value
            End If
        Next i
        ws.Cells(2, 15).Value = max_ticker
        ws.Cells(2, 16).Value = FormatPercent(max_value)
        
     'Calculates the worst performer stock with ticker symbol
        
        min_value = ws.Cells(2, 12).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 12).Value < min_value) Then
            min_ticker = ws.Cells(i, 9).Value
            min_value = ws.Cells(i, 12).Value
            End If
        Next i
        ws.Cells(3, 15).Value = min_ticker
        ws.Cells(3, 16).Value = FormatPercent(min_value)
        
     'Calculates maximum volume stock with ticker symbol
        
        max_volume_value = ws.Cells(2, 10).Value
        For i = 3 To summary_row
            If (ws.Cells(i, 10).Value > max_volume_value) Then
            max_volume_ticker = ws.Cells(i, 9).Value
            max_volume_value = ws.Cells(i, 10).Value
            End If
        Next i
        ws.Cells(4, 15).Value = max_volume_ticker
        ws.Cells(4, 16).Value = max_volume_value


Next ws

End Sub
