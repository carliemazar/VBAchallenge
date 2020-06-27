Attribute VB_Name = "Module1"
Sub StockData():
    'Create the loops within one loop that will go through every sheet ".ws"
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
      
        'Add header labels
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'Generate variables
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        
        'need to reset volume for the loop
        
        Volume = 0
        Dim Row As Double
        
        'start the loops on the row that the data is on
        
        Row = 2
        Dim Column As Integer
        
        'start loops per ws on the first column, then add additional lengths to equal desired column
        Column = 1
        
        'need to classify i as a larger integer, using "Long"
        Dim i As Long
        
    'Loop through all data to output the ticker symbol if change, the yearly change, percent change, and total stock volume
         
         
        'Establish the lastrow and open price variable to use in the loops
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Open_Price = Cells(2, Column + 2).Value

        For i = 2 To LastRow
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                'Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                
                'Recall close price to establish yearly change
                Close_Price = Cells(i, Column + 5).Value
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                
                'Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    
                    'Format cells to percentage
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                    
                End If
                
              'Within the same for loop, calculate total stock volume
              
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                'Increment the rows while looping
                
                Row = Row + 1
                
                'open price reset
                
                Open_Price = Cells(i + 1, Column + 2)
                
                ' reset total volume
                
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
            
        Next i
        
    'Format coloring of the Yearly Change column
        
        'Establish a last row for yearly change to loop through
        
        ycLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
        For j = 2 To ycLastRow
        
        'Change color to green for positive change
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 4
                
        'Change color to red for negative change
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        'Create headers for greatest % increase/decrease, and total volume - with corresponding ticker and value
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
    'Now loop through all data to pull
    
        'Loop for Greatest % increase and Decrease
        
        For c = 2 To ycLastRow
            If Cells(c, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & ycLastRow)) Then
                Cells(2, Column + 15).Value = Cells(c, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(c, Column + 10).Value
                
                'format for percentage
                Cells(2, Column + 16).NumberFormat = "0.00%"
                
            ElseIf Cells(c, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & ycLastRow)) Then
                Cells(3, Column + 15).Value = Cells(c, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(c, Column + 10).Value
                
                'format for percentage
                Cells(3, Column + 16).NumberFormat = "0.00%"
                
            'Loop for Greatest Total Volume
            
            ElseIf Cells(c, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & ycLastRow)) Then
                Cells(4, Column + 15).Value = Cells(c, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(c, Column + 11).Value
            End If
        Next c
        
        'must continue looping through the entire workbook, so indicate to move onto next sheet
        
    Next WS
        
End Sub

