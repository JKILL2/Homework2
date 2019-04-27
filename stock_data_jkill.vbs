Sub stock_data()
    
    'Create a loop to cycle through each tab
    Dim ws As Worksheet

    'Begin loop
    For Each ws In Worksheets

        'Create labels for the summary
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Set variable for ticker
        Dim ticker As String
        
        'Keep track of location for each ticker in summary
        Dim rowcount As Long
        rowcount = 2
        
        'Set variable to hold volume of stock
        Dim totalvolume As Double
        totalvolume = 0

        'Set variable to get the change in price
        Dim year_dif As Double
        year_dif = 0

        'Set variable to hold the percent change for the year
        Dim percent_chg As Double
        percent_chg = 0
 
        'Set variable to year opening and closing price
        Dim begin_year As Double
        begin_year = 0
        Dim end_year As Double
        end_year = 0
        
        'Set variable to loop through till the last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search through tickers
        For i = 2 To lastrow
        
        'Conditional to grab year opening price
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        begin_year = ws.Cells(i, 3).Value

        End If

        'Calculate the total stock volume for each row
        totalvolume = totalvolume + ws.Cells(i, 7)

        'Conditional to figure how much the ticker changed
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        'Move ticker to the summary
        ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

        'Move total stock volume to the summary
        ws.Cells(rowcount, 12).Value = totalvolume

        'Grab year end price
        end_year = ws.Cells(i, 6).Value

        'Calculate the price change for the year and move it to the summary
        year_dif = end_year - begin_year
        ws.Cells(rowcount, 10).Value = year_dif

        'Format cell color to indicate if (+) or (-) change
        If year_dif >= 0 Then
            ws.Cells(rowcount, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(rowcount, 10).Interior.ColorIndex = 3
        End If

        'Calculate the % change for the year and move it to the summary
        If begin_year = 0 And end_year = 0 Then
            percent_chg = 0
            ws.Cells(rowcount, 11).Value = percent_chg
            ws.Cells(rowcount, 11).NumberFormat = "0.00%"
        
        ElseIf begin_year = 0 Then
        
        'If a stock starts at 0 and increases, the % growth will be huge, therefore put "NA" as % change
        Dim NA_growth As String
            NA_growth = "NA"
            ws.Cells(rowcount, 11).Value = percent_chg
        
        Else
            percent_chg = year_dif / begin_year
            ws.Cells(rowcount, 11).Value = percent_chg
            ws.Cells(rowcount, 11).NumberFormat = "0.00%"
        End If

        'Add 1 to row to move it to the next row in the summary
        rowcount = rowcount + 1

        'Reset values to 0
        totalvolume = 0
        begin_year = 0
        end_year = 0
        year_dif = 0
        percent_chg = 0
        
        End If
            
    Next i

Next ws

End Sub
