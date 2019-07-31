Sub HardChallenge()
    'Variable to store summary table row
    Dim summary_table_row as Integer
    summary_table_row = 2

    'Variable for last row in sheet
    Dim last_row as Long

    'Variables for calculations
    Dim open_price, close_price, total_volume, max_pct_increase as Double
    Dim max_pct, min_pct, max_volume As Double
    Dim row_index as Integer


    ' Loop through all sheets
    For Each ws In Worksheets
        'MsgBox(ws.Name)
        ws.Activate

        'First, create new columns in each worksheet
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"


        'find last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'set variables for new sheets
        total_volume = Cdbl(0)
        summary_table_row = 2
        row_index = 0

        'Loop through all rows in sheet to set ticker stats
        For r = 2 to last_row
            total_volume = total_volume + Cdbl(Cells(r, 7).Value)

            If Cells(r,1).Value<> Cells(r-1, 1).Value Then
                'first row of ticker, let's set the opening price
                open_price = Cells(r,3).Value
            ElseIf Cells(r,1).Value <> Cells(r+1, 1).Value Then
                'last row of ticker, let set closing price
                close_price = Cells(r,6).Value

                'set summary data
                Range("I" & summary_table_row).Value = Cells(r,1).Value
                Range("J" & summary_table_row).Value = close_price - open_price 'price difference
                If open_price = 0 Then
                    Range("K" & summary_table_row).Value = "N/A"
                Else
                    Range("K" & summary_table_row).Value = (close_price - open_price) / open_price '% change
                    Range("K" & summary_table_row).NumberFormat ="#.00%" 'format as % with two decimal places
                End If
                
                Range("L" & summary_table_row).Value = total_volume 'total volume

                'set cell colors
                If (close_price - open_price) < 0 Then
                    Range("J" & summary_table_row).Interior.ColorIndex = 3 'red
                Else
                    Range("J" & summary_table_row).Interior.ColorIndex = 4 'green
                End If

                'reset variables
                total_volume = 0
                summary_table_row = summary_table_row + 1

            End if

        Next r

        ' Set Greatest Stats
        
        'find last row of ticker summary data
        last_row = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

        'find max pct increase from summary data and set values in sheet
        max_value = WorksheetFunction.Max(Range("K2:K" & last_row).Value)
        
        'find row index of max value (need to add one as we're starting with the second row)
        row_index = WorksheetFunction.Match(max_value, Range("K2:K" & last_row), 0) + 1 

        'Set Greatest Increase % Value
        Range("P2").Value = Range("I" & row_index).Value 'ticker
        Range("Q2").Value = max_value 'value
        Range("Q2").NumberFormat ="#.00%" 'format as % with two decimal places

        'find min pct increase from summary data and set values in sheet
        min_value = WorksheetFunction.Min(Range("K2:K" & last_row).Value)
        
        'find row index of min value (need to add one as we're starting with the second row)
        row_index = WorksheetFunction.Match(min_value, Range("K2:K" & last_row), 0) + 1 
        
        'Set Greatest Decrease% Value
        Range("P3").Value = Range("I" & row_index).Value 'ticker
        Range("Q3").Value = min_value 'value
        Range("Q3").NumberFormat ="#.00%" 'format as % with two decimal places

        'find max volume from summary data and set values in sheet
        max_volume= WorksheetFunction.Max(Range("L2:L" & last_row).Value)
        
        'find row index of max volume (need to add one as we're starting with the second row)
        row_index = WorksheetFunction.Match(max_volume, Range("L2:L" & last_row), 0) + 1 
        
        'Set volume Value
        Range("P4").Value = Range("I" & row_index).Value 'ticker
        Range("Q4").Value = max_volume 'value
        
  
        ' Autofit to display data nicely
        ws.Columns("A:Q").AutoFit

        'MsgBox("Next WS")
    Next ws
    ' set view back to first worksheet
    Worksheets(1).Activate
End Sub