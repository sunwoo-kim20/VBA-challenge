'Sunwoo Kim - UCI DA Bootcamp VBA-Challenge: Homework 2
'VBA Script to analyze stock market data

Sub stock_analysis():
    'Initialize variables for ticker, total stock value, opening/closing prices,
    'price change, row count, flag, and last row.
    Dim first_cell As String
    Dim second_cell As String
    Dim total_stock As Long
    Dim row_count As Integer
    Dim last_row As Long
    Dim open_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim ticker_flag As Boolean
    total_stock = 0
    row_count = 1
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    ticker_flag = True
    'Naming column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'For loop to iterate through the observations
    For i = 2 to last_row
        first_cell = Cells(i, 1).Value
        second_cell = Cells(i+1, 1).Value
        total_stock += Cells(i, 7).Value
        'If statement with flag to assign original opening price
        If ticker_flag Then
            opening_price = Cells(i, 3).Value
            ticker_flag = False
        End If
        'If statement to catch when ticker value changes
        If first_cell <> second_cell Then
            Cells(row_count, 9).Value = first_cell
            closing_price = Cells(i, 6).Value
            yearly_change = opening_price - closing_price
            Cells(row_count, 10).Value = yearly_change

            'Coloring cell according to yearly change
            If yearly_change < 0 Then
                Cells(row_count, 10).Interior.ColorIndex = 3
            Elseif yearly_change > 0 Then
                Cells(row_count, 10).Interior.ColorIndex = 4
            End If

            'Calculating percent change + checking for division by 0
            If opening_price <> 0 Then
                percent_change = yearly_change/opening_price
            Else
                percent_change = 0
            End If
            Cells(row_count, 11).Value = percent_change
            Cells(row_count, 11).NumberFormat = "0.00%"
            

        End If
    Next i
End Sub
