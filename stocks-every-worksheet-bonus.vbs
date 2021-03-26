'Sunwoo Kim - UCI DA Bootcamp VBA-Challenge: Homework 2 Bonus
'VBA Script to analyze stock market data on every worksheet at once

Sub stock_analysis_all_ws():
    For Each ws in Worksheets
        'Initialize variables for ticker, total stock value, opening/closing prices,
        'price change, row count, flag, and last row.
        Dim first_cell As String
        Dim second_cell As String
        Dim total_stock As Double
        Dim row_count As Integer
        Dim last_row As Long
        Dim open_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim ticker_flag As Boolean
        total_stock = 0
        row_count = 2
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ticker_flag = True

        'Naming column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'For loop to iterate through the observations
        For i = 2 to last_row
            first_cell = ws.Cells(i, 1).Value
            second_cell = ws.Cells(i+1, 1).Value
            total_stock = total_stock + ws.Cells(i, 7).Value
            'If statement with flag to assign original opening price
            If ticker_flag Then
                opening_price = ws.Cells(i, 3).Value
                ticker_flag = False
            End If
            'If statement to catch when ticker value changes
            If first_cell <> second_cell Then
                ws.Cells(row_count, 9).Value = first_cell
                closing_price = ws.Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                ws.Cells(row_count, 10).Value = yearly_change

                'Coloring cell according to yearly change
                If yearly_change < 0 Then
                    ws.Cells(row_count, 10).Interior.ColorIndex = 3
                Elseif yearly_change > 0 Then
                    ws.Cells(row_count, 10).Interior.ColorIndex = 4
                End If

                'Calculating percent change + checking for division by 0
                If opening_price <> 0 Then
                    percent_change = yearly_change/opening_price
                Else
                    percent_change = 0
                End If
                ws.Cells(row_count, 11).Value = percent_change
                ws.Cells(row_count, 11).NumberFormat = "0.00%"

                ws.Cells(row_count, 12).Value = total_stock
                total_stock = 0
                row_count = row_count + 1
                ticker_flag = True
            End If
        Next i
    Next ws
End Sub

'VBA Script to identify greatest percent increase/decrease and total volume on
'every worksheet

Sub greatest_stock_analysis_all_ws():
    For Each ws in Worksheets
        'Setting up row/column labels
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1,16).Value = "Ticker"
        ws.Cells(1,17).Value = "Value"

        'Initializing temporary max variables and indices of max decrease, increase,
        'and volume
        Dim max_increase, max_decrease, max_volume As Double
        Dim max_increase_index, max_decrease_index, max_volume_index As Integer
        Dim last_row As Integer
        last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        max_increase = ws.Cells(2,11).Value
        max_decrease = ws.Cells(2,11).Value
        max_volume = ws.Cells(2,12).Value
        max_increase_index = 2
        max_decrease_index = 2
        max_volume_index = 2
        ' Loop through all rows of summarized stock data
        For i = 2 to last_row

            'Compare current max/min to next value and update value and index if
            ' necessary
            If ws.Cells(i,11).Value > max_increase Then
                max_increase = ws.Cells(i,11).Value
                max_increase_index = i
            End If

            If ws.Cells(i,11).Value < max_decrease Then
                max_decrease = ws.Cells(i,11).Value
                max_decrease_index = i
            End If

            If ws.Cells(i,12).Value > max_volume Then
                max_volume = ws.Cells(i,12).Value
                max_volume_index = i
            End If
        Next i

        'Print out final max/min values along with corresponding ticker value
        ws.Cells(2, 16).Value = ws.Cells(max_increase_index,9).Value
        ws.Cells(2, 17).Value = max_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"


        ws.Cells(3, 16).Value = ws.Cells(max_decrease_index,9).Value
        ws.Cells(3, 17).Value = max_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"

        ws.Cells(4, 16).Value = ws.Cells(max_volume_index,9).Value
        ws.Cells(4, 17).Value = max_volume
    Next ws
End Sub
