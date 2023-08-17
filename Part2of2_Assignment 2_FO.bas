Attribute VB_Name = "Module2"
Sub secondSummary()

For Each ws In Worksheets

lastrow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

        'Identify highest percent change for the second summary table
        gr_per_inc = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
        'Identify the location of the matching ticker
        ticker_val = WorksheetFunction.Match(gr_per_inc, ws.Range("K2:K" & lastrow), 0)
        'Finding the matching ticker to the location
        ticker_val = ticker_val + 1
        maxPosition = ws.Cells(ticker_val, 9).Value
        'Print maxPosition into the corresponding cell
        ws.Cells(2, 16).Value = maxPosition
        
        'Identify lowest percent change for the second summary table
        gr_per_dec = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
        'Identify the location of the matching ticker
        ticker_two = WorksheetFunction.Match(gr_per_dec, ws.Range("K2:K" & lastrow), 0)
        'Finding the matching ticker to the location
        ticker_two = ticker_two + 1
        minPosition = ws.Cells(ticker_two, 9).Value
        'Print minPosition into the corresponding cell
        ws.Cells(3, 16).Value = minPosition

        'Formatting greatest increase and decrease to percentage
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'Identify greatest total volume for the second summary table
        gr_vol = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        'Identify the location of the matching ticker
        ticker_three = WorksheetFunction.Match(gr_vol, ws.Range("L2:L" & lastrow), 0)
        'Finding the matching ticker to the location
        ticker_three = ticker_three + 1
        gr_total = ws.Cells(ticker_three, 9).Value
        
        'Print the values in the correct cells
        ws.Cells(2, 17).Value = gr_per_inc
        ws.Cells(3, 17).Value = gr_per_dec
        ws.Cells(4, 17).Value = gr_vol
        ws.Cells(4, 16).Value = gr_total
       
        
Next ws

End Sub

