Attribute VB_Name = "Module1"
Sub Ticker():

    For Each ws In Worksheets

        'Identifies the last row of the spreadsheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Defines variables
        Dim ticker_symbol As String
        Dim spot As Integer
        Dim year_close As Double
        Dim year_open As Double
        Dim yearly_change As Double
    
        'Creates labels
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'Initiates the counter for total_volume
        total_volume = 0
    
        'Inititates the counter for the location in the output columns
        spot = 2
    
        'Cycles through the rows until the end of the spreadsheet
        For i = 1 To lastrow
    
            'Adds up the stock volumes
            If i > 1 Then
                total_volume = total_volume + ws.Cells(i, 7)
            
                'Identifies the greatest volume
                If total_volume > greatest_volume Then
                    greatest_volume = total_volume
                    ws.Cells(4, 17).Value = greatest_volume
                    ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                End If
        
            End If
            
            'This all happens when the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                If i > 1 Then
                
                    'Creates the list of unique ticker symbols
                    ticker_symbol = ws.Cells(i, 1).Value
                    ws.Cells(spot, 9).Value = ticker_symbol
                   
                    'Identifies the closing value
                    year_close = ws.Cells(i, 6)
           
                    'Calculates the yearly change
                    yearly_change = year_close - year_open
                    ws.Cells(spot, 10) = yearly_change
                
                        'Conditionally formats the Yearly Change: positive = green / negative = red
                        If ws.Cells(spot, 10).Value < 0 Then
                            ws.Cells(spot, 10).Interior.ColorIndex = 3
                        Else
                            ws.Cells(spot, 10).Interior.ColorIndex = 4
                        End If
                
                    'Calculates the Percent Change and formats it as a percent
                    percent_change = yearly_change / year_open
                    ws.Cells(spot, 11) = percent_change
                    ws.Cells(spot, 11).NumberFormat = "0.00%"
                
                    'Identifies the Greatest Percent Increase
                    If percent_change > greatest_percent Then
                    
                        greatest_percent = percent_change
                        ws.Cells(2, 17).Value = greatest_percent
                        ws.Cells(2, 17).NumberFormat = "0.00%"
                        ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
                    
                    End If
                
                    'Identifies the Greatest Percent Decrease
                    If percent_change < least_percent Then
                    
                        least_percent = percent_change
                        ws.Cells(3, 17).Value = least_percent
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                        ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                
                    End If
            
                    'Create the list of total volume
                    ws.Cells(spot, 12) = total_volume
                    total_volume = 0
                
                    'Moves down the output column
                    spot = spot + 1
           
                End If
           
                'Identifies the opening value
                year_open = ws.Cells(i + 1, 3)
                
            End If
    
        Next i

    'Adjusts the Column Widths to the contents
    ws.Columns(9).AutoFit
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
    ws.Columns(15).AutoFit
    ws.Columns(16).AutoFit
    ws.Columns(17).AutoFit

    'Resets the check values before moving to the next worksheet
    greatest_volume = 0
    greatest_percent = 0
    least_percent = 0

    Next ws

End Sub



