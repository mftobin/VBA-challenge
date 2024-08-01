Sub tickercolumn()
For Each WS In Worksheets
    Dim WorksheetName As String
    ' definining the last row
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    'not sure if this is something I need
    WorksheetName = WS.Name
    'labeling columns I-L, leaving column H blank
    'I did not insert columns because there were already blank columns available
    'to insert a column, I would have put in "ws.Range("H1").EntireColumn.Insert"
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Quarterly Change ($)"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"
    WS.Range("M1").Value = "Total Opening Price"
    WS.Range("N1").Value = "Total Closing Price"
    
    Dim Stock_Ticker As String
    Dim Quarterly_Change As Integer
    Dim Total_Opening_Price As LongLong
    Dim Total_Closing_Price As LongLong
    Dim Percent_Change As Integer
    Dim Total_Stock_Volume As LongLong
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    
    'Defining the first column as Ticker and condensing that information into column I
    
    'Loop through all the stocks
    For i = 2 To LastRow
        
        'check if we are still within the same stock symbol, if not...
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
            'set the stock ticker
            Stock_Ticker = WS.Cells(i, 1).Value
            
            'set the total opening price
            Total_Opening_Price = Total_Opening_Price + WS.Cells(i, 3).Value
            
            'set the total closing price
            Total_Closing_Price = Total_Closing_Price + WS.Cells(i, 6).Value
            
            'add the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value

            'print the stock symbol in the summary table
            WS.Range("I" & Summary_Table_Row).Value = Stock_Ticker
            
            'print the total opening price in the summary table
            WS.Range("M" & Summary_Table_Row).Value = Total_Opening_Price
            
            'print the total closing price in the summary talbe
            WS.Range("N" & Summary_Table_Row).Value = Total_Closing_Price
            
            'print the total stock volume in the summary table
            WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
            
            'add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset the total opening price
            Total_Opening_Price = 0
            
            'reset the total closing price
            Total_Closing_Price = 0
        
            'reset the total stock volume
            Total_Stock_Volume = 0
            
        'If the cell immediately following a row is the same stock symbol...
        Else
            'add to the total opening price
            Total_Opening_Price = Total_Opening_Price + WS.Cells(i, 3).Value

            'add to the total closing price
            Total_Closing_Price = Total_Closing_Price + WS.Cells(i, 6).Value
            
            'add to the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
        End If

    Next i
    'Now that our summary table has the stock symbols, the total volume, and the total opening/closing prices,
        'let's calculate quarterly change and percentage change
        'I'm sure I should have been able to do all of this in a nested for loop but I couldn't figure it out and my brain hurts
     
Next WS

End Sub

Sub Summary_Quarterly()
For Each WS In Worksheets
    Dim WorksheetName As String
    ' definining the last row
    LastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
    'not sure if this is something I need
    WorksheetName = WS.Name
    'labeling columns I-L, leaving column H blank
    'I did not insert columns because there were already blank columns available
    'to insert a column, I would have put in "ws.Range("H1").EntireColumn.Insert"

    Dim Quarterly_Change As Integer
    Dim Total_Opening_Price As LongLong
    Dim Total_Closing_Price As LongLong
    Dim Percent_Change As Integer
    Dim Total_Stock_Volume As LongLong
    For i = 2 To LastRow
        Quarterly_Change = WS.Cells(i, 14).Value - WS.Cells(i, 13).Value
        WS.Cells(i, 10).Value = Quarterly_Change
        WS.Cells(i, 10).NumberFormat = "0.00"
        If WS.Cells(i, 10) > 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf WS.Cells(i, 10) < 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 3
        Else
            WS.Cells(i, 10).Interior.ColorIndex = 0
    End If
    
    Next i

Next WS

End Sub

Sub Summary_Percent()
For Each WS In Worksheets
    Dim WorksheetName As String
    ' definining the last row
    LastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
    'not sure if this is something I need
    WorksheetName = WS.Name
    'labeling columns I-L, leaving column H blank
    'I did not insert columns because there were already blank columns available
    'to insert a column, I would have put in "ws.Range("H1").EntireColumn.Insert"

    Dim Quarterly_Change As Double
    Dim Total_Opening_Price As LongLong
    Dim Total_Closing_Price As LongLong
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As LongLong
    For i = 2 To LastRow
        Percent_Change = WS.Cells(i, 10).Value / WS.Cells(i, 13).Value
        WS.Cells(i, 11).Value = Percent_Change
        WS.Cells(i, 11).NumberFormat = "0.00%"
    Next i

Next WS

End Sub


Sub CalculatedValues()
Dim Ticker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim TotalVolume As LongLong


For Each WS In Worksheets
    WS.Range("Q1").Value = "Ticker"
    WS.Range("R1").Value = "Value"
    WS.Range("P2").Value = "Greatest % Increase"
    WS.Range("P3").Value = "Greatest % Decrease"
    WS.Range("P4").Value = "Greatest Total Volume"
    Dim WorksheetName As String
    Dim GreatestTotal As LongLong

    ' definining the last row
    LastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
    'not sure if this is something I need
    WorksheetName = WS.Name
    
   For i = 2 To LastRow
        GreatestIncrease = Application.WorksheetFunction.Max(WS.Range("J2:J" & LastRow))
        WS.Range("R2") = GreatestIncrease
        WS.Range("R2").NumberFormat = "0.00"
        If WS.Cells(i, 10).Value = GreatestIncrease Then
            WS.Range("Q2").Value = WS.Cells(i, 9).Value
        End If
        GreatestDecrease = Application.WorksheetFunction.Min(WS.Range("J2:J" & LastRow))
        WS.Range("R3") = GreatestDecrease
        WS.Range("R3").NumberFormat = "0.00"
        If WS.Cells(i, 10).Value = GreatestDecrease Then
            WS.Range("Q3").Value = WS.Cells(i, 9).Value
        End If
        GreatestTotal = Application.WorksheetFunction.Max(WS.Range("L2:L" & LastRow))
        WS.Range("R4") = GreatestTotal
        If WS.Cells(i, 12).Value = GreatestTotal Then
            WS.Range("Q4").Value = WS.Cells(i, 9).Value
        End If
    Next i
    WS.Columns("A:R").AutoFit
    WS.Range("M:N").EntireColumn.Hidden = True

Next WS
End Sub


