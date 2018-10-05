Sub Stock_Analysis()
    
    'Create a loop to cycle through the worksheets in the workbook
    'Set a variable to cycle through the worksheets
    Dim ws as Worksheet

    'Start loop
    For Each ws in Worksheets

        'Create column labels for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Set variable to hold the ticker symbol
        Dim ticker_symbol As String

        'Set variable to hold total volume of stock traded
        Dim total_vol As Double
        total_vol = 0

        'Keep tracker of location for each ticker symbol in the summary table
        Dim rowcount As Long
        rowcount = 2

        'Set variable to hold year open price
        Dim year_open As Double
        year_open = 0

        'Set variable to hold year close price
        Dim year_close As Double
        year_close = 0
        
        'Set variable to hold the change in price for the year
        Dim year_change As Double
        year_change = 0

        'Set variable to hold the percent change in price for the year
        Dim percent_change As Double
        percent_change = 0

        'Set variable for total rows to loop through
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search through ticker symbols
        For i = 2 To lastrow
            
            'Conditional to grab year open price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

            'Total up the volume for each row to determine the total stock volume for the year
            total_vol = total_vol + ws.Cells(i, 7)

            'Conditional to determine if the ticker symbol is changing
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Move ticker symbol to summary table
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'Move total stock volume to the summary table
                ws.Cells(rowcount, 12).Value = total_vol

                'Grab year end price
                year_close = ws.Cells(i, 6).Value

                'Calculate the price change for the year and move it to the summary table.
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change

                'Conditional to format to highlight positive or negative change.
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Calculate the percent change for the year and move it to the summary table format as a percentage
                'Conditional for calculating percent change
                If year_open = 0 and year_close = 0 Then
                    'Starting at zero and ending at zero will be a zero increase.  Cannot use a formula because
                    'it would be dividing by zero.
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf year_open = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    'Because of this, we only need to evaluate actual price increase by dollar amount and therefore put
                    '"New Stock" as percent change.
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change                
                Else
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Add 1 to rowcount to move it to the next empty row in the summary table
                rowcount = rowcount + 1

                'Reset total stock volume, year open price, year close price, year change, year percent change
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i

        'Create a best/worst performance table
        'Titles
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Assign lastrow to count the number of rows in the summary table
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables to hold best performer, worst performer, and stock with the most volume
        Dim best_stock as String
        Dim best_value as Double

        'Set best performer equal to the first stock
        best_value = ws.Cells(2, 11).Value

        Dim worst_stock as String
        Dim worst_value as Double

        'Set worst performer equal to the first stock
        worst_value = ws.Cells(2, 11).Value

        Dim most_vol_stock as String
        Dim most_vol_value as Double

        'Set most volume equal to the first stock
        most_vol_value = ws.Cells(2, 12).Value

        'Loop to search through summary table
        for j = 2 to lastrow

            'Conditional to determine best performer
            if ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
            End If

            'Conditional to determine worst performer
            if ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
            End If

            'Conditional to determine stock with the greatest volume traded
            if ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        'Move best performer, worst performer, and stock with the most volume items to the performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        'Autofit table columns
        ws.Columns("I:L").EntireColumn.Autofit
        ws.Columns("O:Q").EntireColumn.Autofit

    Next ws

End Sub