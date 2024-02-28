Sub Challenge3()
        'Declare Variables
        Dim LastRow As Long
        Dim shortRow As Integer
        Dim printRow As Integer
        Dim counter As Integer
        Dim opening As Double
        Dim closing As Double
        Dim totalStock As Double
        Dim inTicker As String
        Dim inValue As Double
        Dim deTicker As String
        Dim deValue As Double
        Dim toTicker As String
        Dim toValue As Double
    
    'For statement to loop through all worksheets
    For Each ws In Worksheets
        'Initialize Variables for each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        printRow = 2
        counter = 0
        opening = 0
        closing = 0
        totalStock = 0
        shortRow = 0
        inTicker = ws.Cells(2, 9).Value
        inValue = ws.Cells(2, 11).Value
        deTicker = ws.Cells(2, 9).Value
        deValue = ws.Cells(2, 11).Value
        toTicker = ws.Cells(2, 9).Value
        toValue = ws.Cells(2, 12).Value
        
        'Print headers for the new data in all worksheets
        ws.Range("I1, P1").Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 17).Value = "Value"
        
        'For statement that loops through the main list of each worksheet
        For i = 2 To LastRow
            
            'Saves the first opening value for each ticker
            If counter = 0 Then
                opening = ws.Cells(i, 3).Value
            End If
          
            'Identifies when a new ticker value is found
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Sets the last closing value
                closing = ws.Cells(i, 6).Value
                'Adds the last closing value
                totalStock = totalStock + ws.Cells(i, 7).Value
                'Print one time every Ticker for each worksheet
                ws.Cells(printRow, 9).Value = ws.Cells(i, 1).Value
                'Print diference between opening and closing of every year
                ws.Cells(printRow, 10).Value = closing - opening
                'Print percentage change
                ws.Cells(printRow, 11).Value = (closing - opening) / opening
                'Print Total Stock Volume
                ws.Cells(printRow, 12).Value = totalStock
                'Moves to the next print row so the code doesn't print over the same cell
                printRow = printRow + 1
                'Restart counters
                counter = 0
                totalStock = 0
            'While in the same ticker add 1 to the counter and adds the new Total Stock value
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                counter = counter + 1
                totalStock = totalStock + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        'Save the reduced list length in variable shortRow
        shortRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'For loop to go through the reduce list in each worksheet
        For j = 2 To shortRow
        
            'Calculate greatest % increase
            If inValue < ws.Cells(j, 11).Value Then
                inTicker = ws.Cells(j, 9).Value
                inValue = ws.Cells(j, 11).Value
            End If
            
            'Calculate greatest % decrease
            If deValue > ws.Cells(j, 11).Value Then
                deTicker = ws.Cells(j, 9).Value
                deValue = ws.Cells(j, 11).Value
            End If
            
            'Calculate greatest total volume
            If toValue < ws.Cells(j, 12).Value Then
                toTicker = ws.Cells(j, 9).Value
                toValue = ws.Cells(j, 12).Value
            End If
            
            'Apply conditional formating for yearly change
             If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
             ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            'Apply conditional formating for percent change
             If ws.Cells(j, 11).Value >= 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 4
             ElseIf ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            End If
            
        Next j
        
        'Print columns in greatest list
        ws.Cells(2, 16).Value = inTicker
        ws.Cells(2, 17).Value = inValue
        ws.Cells(3, 16).Value = deTicker
        ws.Cells(3, 17).Value = deValue
        ws.Cells(4, 16).Value = toTicker
        ws.Cells(4, 17).Value = toValue
        
        'Format Columns/Cells
        'Give the Percent Change column percent format. Reference from URL: https://stackoverflow.com/questions/36654624/set-a-range-to-an-entire-column-with-index-number
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'Size columns accordingly. Reference from URL: https://stackoverflow.com/questions/2048295/how-to-auto-size-column-width-in-excel-during-text-entry
        ws.Columns.AutoFit
        
    Next ws

End Sub