Sub Multiple_year_stock_data()

'Loop for each worksheet
    For Each ws In Worksheets

' Variables and counters
    
    Dim ticker As String
    Dim summary_table_rowA As Long
    Dim summary_table_rowB As Long
    Dim lastrowA As Long
    Dim lastrowB As Long
    Dim lastrow_value As Long
    Dim total_ticker_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_total_volume As Double
    summary_table_rowA = 2
    summary_table_rowB = 2
    total_ticker_volume = 0
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total_volume = 0
    
' Each worksheet active and autofit for headers
    
    ws.Activate
    Columns("I:Q").AutoFit
    
' Title for the Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

' Last row for non-blank cell in column A
    lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
                           
' Determination tickers
    ' Loop for rows
    For i = 2 To lastrowA
    'Values for total ticekers volume
     total_ticker_volume = total_ticker_volume + Cells(i, 7).Value
    ' If row is the same than the last
        If ws.Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    ' Set ticker for first column
            ticker = Cells(i, 1).Value
    ' Summary of the tickers
            Cells(summary_table_rowA, 9).Value = ticker
            Cells(summary_table_rowA, 12).Value = total_ticker_volume
    ' Reset Ticker
            total_ticker_volume = 0
            
' Determination Yearly Change
    ' Set open price
            open_price = Cells(summary_table_rowB, 3)
    ' Set close price
            close_price = Cells(i, 6)
    ' Getting yearly change
            yearly_change = close_price - open_price
            Cells(summary_table_rowA, 10).Value = yearly_change
       
' Determination of Percent Change
    ' Set when is 0
            If open_price = 0 Then
                percent_change = 0
    'Otherwise the division
                Else
                    yearly_open = Cells(summary_table_rowB, 3)
                    percent_change = yearly_change / open_price
            End If
    ' Result and change format to percentage
            Cells(summary_table_rowA, 11).Value = percent_change
            Cells(summary_table_rowA, 11).NumberFormat = "0.00%"

' Formating highlight
    ' If is positive or equal fill the with green
            If Cells(summary_table_rowA, 10).Value >= 0 Then
                Cells(summary_table_rowA, 10).Interior.ColorIndex = 4
                Else
    ' If is negative fill the cell with red
                Cells(summary_table_rowA, 10).Interior.ColorIndex = 3
            End If
    ' Setting table plus one
            summary_table_rowA = summary_table_rowA + 1
            summary_table_rowB = i + 1
            
        End If
        
        Next i
            
 ' Determination of Greatest Increase, Decrease and Total Volume
    ' Last row for non-blank cell
        lastrowB = Cells(Rows.Count, 11).End(xlUp).Row
    ' Loop rows for final table
        For i = 2 To lastrowB
    ' Calculation of Greatest % increase
            If Cells(i, 11).Value > Cells(2, 17).Value Then
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 16).Value = Cells(i, 9).Value
            End If
    ' Calculation for Greates % increase
            If Cells(i, 11).Value < Cells(3, 17).Value Then
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 16).Value = Cells(i, 9).Value
            End If
     ' Calculation of Greatest % Total Volume
            If Cells(i, 12).Value > Cells(4, 17).Value Then
                Cells(4, 17).Value = Cells(i, 12).Value
                Cells(4, 16).Value = Cells(i, 9).Value
            End If
        
            Next i
     ' Changes formats % with 2 decimals
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
          
    Next ws

End Sub
