'Sunwoo Kim - UCI DA Bootcamp VBA-Challenge: Homework 2 Bonus
'VBA Script to identify greatest percent increase/decrease and total volume

Sub greatest_stock_analysis():
    'Setting up row/column labels
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1,16).Value = "Ticker"
    Cells(1,17).Value = "Value"

    'Initializing temporary max variables and indices of max decrease, increase,
    'and volume
    Dim max_increase, max_decrease, max_volume As Double
    Dim max_increase_index, max_decrease_index, max_volume_index As Integer
    Dim last_row As Integer
    last_row = Cells(Rows.Count, 9).End(xlUp).Row
    max_increase = Cells(2,11).Value
    max_decrease = Cells(2,11).Value
    max_volume = Cells(2,12).Value
    max_increase_index = 2
    max_decrease_index = 2
    max_volume_index = 2
    ' Loop through all rows of summarized stock data
    For i = 2 to last_row

        'Compare current max/min to next value and update value and index if
        ' necessary
        If Cells(i,11).Value > max_increase Then
            max_increase = Cells(i,11).Value
            max_increase_index = i
        End If

        If Cells(i,11).Value < max_decrease Then
            max_decrease = Cells(i,11).Value
            max_decrease_index = i
        End If

        If Cells(i,12).Value > max_volume Then
            max_volume = Cells(i,12).Value
            max_volume_index = i
        End If
    Next i

    'Print out final max/min values along with corresponding ticker value
    Cells(2, 16).Value = Cells(max_increase_index,9).Value
    Cells(2, 17).Value = max_increase
    Cells(2, 17).NumberFormat = "0.00%"


    Cells(3, 16).Value = Cells(max_decrease_index,9).Value
    Cells(3, 17).Value = max_decrease
    Cells(3, 17).NumberFormat = "0.00%"

    Cells(4, 16).Value = Cells(max_volume_index,9).Value
    Cells(4, 17).Value = max_volume

End Sub
