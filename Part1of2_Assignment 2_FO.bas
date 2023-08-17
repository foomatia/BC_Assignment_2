Attribute VB_Name = "Module1"

'PART I of II of the complete code

Sub TickerSummary()
    
    'Definining the variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim Total_Stock_Volume As Double
    Dim yr_open As Double
    Dim yr_close As Double
    Dim yr_high As Double
    Dim yr_low As Double
    Dim yr_date As Long
    Dim vol As Long
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim gr_per_inc As Double
    Dim gr_per_dec As Double
    Dim gr_total As Double
    
    
    'Creating a loop through all worksheets in this workbook
    For Each ws In ThisWorkbook.Worksheets
    
    On Error Resume Next
    
    'Define the headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Keeping track of the location for each ticker
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'Defining last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Pull opening value of first ticker before loop starts
    yr_open = ws.Cells(2, 3).Value
    
    'Loop through all ticker items
    For i = 2 To lastrow
    
    'Check if it is still the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Define ticker location
        ticker = ws.Cells(i, 1).Value
        'Define vol location
        vol = ws.Cells(i, 7).Value
        
        'Add to the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        'Print the ticker name in Column I
        ws.Range("I" & Summary_Table_Row).Value = ticker
        
        'Print the Total Stock Volume in column L
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
          
        'Reset total stock volume
        Total_Stock_Volume = 0
        
        
        'Pull closing value of ticker
        yr_close = ws.Cells(i, 6).Value
        
        'Calculate the year change
        YearChange = yr_close - yr_open
        
        'Color coding negative and positive numbers in the Yearly Change column
        If YearChange >= 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        'Percent change calculation
        PercentChange = (YearChange / yr_open)
        
        'Print the Percent Change in column K
        ws.Range("K" & Summary_Table_Row).Value = PercentChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'Pull opening value of next ticker
        yr_open = ws.Cells(i + 1, 3).Value
        
        'Print the Yearly Change in column J
        ws.Range("J" & Summary_Table_Row).Value = YearChange
        
        
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
    Else
        'Add to the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    End If
    
Next i
Next ws
    
End Sub


        





