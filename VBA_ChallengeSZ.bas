Attribute VB_Name = "SZCode"
Sub StockData()
'written with assistance of study group, office hours & professor

'Loop through all sheets
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

    'declare variables
     
     Dim ticker As String
     Dim rowindex As Integer
     Dim open_price As Double
     Dim close_price As Double
     Dim yearly_change As Double
     Dim percent_change As Double
     Dim totalstockvol As Double
     
     rowindex = 2
     totalstockvol = 0
     
    'Fill in header values - start at I
    
     Range("I1").Value = "Ticker"
     Range("J1").Value = "Yearly Change"
     Range("K1").Value = "Percent Change"
     Range("L1").Value = "Total Stock Volume"
     
     'set first openprice
     open_price = Cells(2, 3).Value
       
     'loop through all the tickers
     'start at row 2 due to header and go to end of data
     For i = 2 To Range("A1").CurrentRegion.End(xlDown).Row
     
        'check if cell below is the same ticker, if not then:
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
           'set ticker name & put in table
            ticker = Cells(i, 1).Value
            Cells(rowindex, 9).Value = ticker
        
            'Set close price
            close_price = Cells(i, 6).Value
        
            'compute yearly change & put in table
            yearly_change = close_price - open_price
            Cells(rowindex, 10).Value = yearly_change
        
            'Increase total stock volume & put in table
            totalstockvol = totalstockvol + Cells(i, 7)
            Cells(rowindex, 12).Value = totalstockvol
           
            'Compute percent change
            If (open_price = 0 And close_price = 0) Then
                percent_change = 0
            ElseIf (open_price = 0 And close_price <> 0) Then
                percent_change = 1
            Else
                percent_change = yearly_change / open_price
                
                'put in table
                Cells(rowindex, 11).Value = percent_change
                Cells(rowindex, 11).NumberFormat = "0.00%"
            End If
            
            'move to next row in table
            rowindex = rowindex + 1
            
            'reset values for next loop
            open_price = Cells(i + 1, 3).Value
            totalstockvol = 0
            
            'If cell tickers are the same:
            Else
                totalstockvol = totalstockvol + Cells(i, 7).Value
         
         End If
    Next i
    
    'Find last row based off part 1 created table
    Lastrow = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Colors!
    For i = 2 To Lastrow
        If Cells(i, 10) >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
        
        
'Challenge Part
    'Set up table with headers to fill out (did it as Range last time, doing it cells this time to practice)
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Loop through to find greatest values and match with ticker
    'Fill out table with values & ticker
    For i = 2 To Lastrow
        If Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & Lastrow)) Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & Lastrow)) Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & Lastrow)) Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        End If
    Next i
    
Next ws

End Sub
