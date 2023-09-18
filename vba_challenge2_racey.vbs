Sub stock_ticker()

' Declare Current as a worksheet object variable.
Dim ws As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each ws In Worksheets

    ' lastRow to use in loop
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set ticker symbol variable
    Dim ticker_symbol As String
    
    ' Set variable for ticker count/row
    Dim ticker_row As Integer
    ticker_row = 2
    
    ' Set variables for open, close, yearly, percent change prices & stock volume total
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim ticker_total As Double
    
    
    ' Set column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set calculated value table
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' From Microsoft, to fit column widths to text in headers
    ws.Range("I1:L1").Columns.AutoFit
    ws.Range("O2:O4").Columns.AutoFit

    ' Set inital values for open price and ticker total
    open_price = Cells(2, 3).Value
    ticker_total = 0
    
    ' Loop through all of the stocks
    For i = 2 To lastRow
    
    
    
        ' Check to see if the row is within the same ticker symbol or we have a new one
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set ticker symbol to variable
            ticker_symbol = ws.Cells(i, 1).Value
            
            ' print ticker symbol in list
            ws.Cells(ticker_row, 9).Value = ticker_symbol
            
            ' find yearly change in stock price & set to 2 decimals
            
            close_price = ws.Cells(i, 6).Value
            ws.Cells(ticker_row, 10).Value = close_price - open_price
            ws.Range("A2:A" & lastRow).NumberFormat = "0.00"

            ' find percent change, check for open price - 0
            If open_price <> 0 Then
                percent_change = ((close_price - open_price) / open_price)
            
            ' If open price is 0 then set percent to 100% to avoid error
            Else
                percent_change = 1
                
            End If
            ws.Cells(ticker_row, 11).Value = percent_change
            
            ' set to percentage
            ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
            
            ' conditional formatting - green for positive, red for negative
            If percent_change > 0 Then
                ws.Cells(ticker_row, 11).Interior.ColorIndex = 4
            ElseIf percent_change < 0 Then
                ws.Cells(ticker_row, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(ticker_row, 11).Interior.ColorIndex = 2
            End If
            
            ' set ticker volume total
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            ws.Cells(ticker_row, 12).Value = ticker_total
            
            ' Add one to the ticker count/row
            ticker_row = ticker_row + 1
            
            ' set next ticker's open price and reset ticker volume total
            open_price = ws.Cells(i + 1, 3).Value
            ticker_total = 0
            
        ' If we have the same ticker symbol
        Else
            
            ' add to ticker volume total
            ticker_total = ticker_total + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    ' Last row for ticker symbol & change table
    Dim lastRow2 As Long
    lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Set inital values for values in summary table
    Dim max_inc As Double
    Dim min_inc As Double
    Dim max_vol As Double
    
    max_inc = ws.Cells(2, 11).Value
    max_dec = ws.Cells(2, 11).Value
    max_vol = ws.Cells(2, 12).Value
    
    ' Loop for summary table
    For i = 2 To lastRow2
        
        ' Check if increase is greater than current greatest increase
        If ws.Cells(i, 11).Value > max_inc Then
            max_inc = ws.Cells(i, 11).Value

        ' Print greatest increase ticker and value    
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = max_inc
        ws.Range("Q2").NumberFormat = "0.00%"
            
        ' if not greater than current greatest increase, keep same value stored
        Else
            max_inc = max_inc
        
        End If
        
        ' Check if decrease is greater than current greatest decrease
        If ws.Cells(i, 11).Value < max_dec Then
            max_dec = ws.Cells(i, 11).Value

        ' Print greatest decrease ticker and value    
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = max_dec
        ws.Range("Q3").NumberFormat = "0.00%"
                    
        ' if not greater than current greatest decrease, keep same value stored
        Else
            max_dec = max_dec
        
        End If
        
        ' Check if ticker volume is greater than current greatest total volume
        If ws.Cells(i, 12).Value > max_vol Then
            max_vol = ws.Cells(i, 12).Value

        ' Print greatest ticker volume and value    
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = max_vol
     
        ' if not greater than current greatest ticker volume, keep same value stored
        Else
            max_vol = max_vol
        
        End If


    Next i

Next

End Sub