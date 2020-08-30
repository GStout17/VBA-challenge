Attribute VB_Name = "Module1"
Sub VBA_Challenge():

'Defining variables
Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
        'Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        'Add titles/headings to columns
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'variables (define rows and or columns?)
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Column = 1
        
        'Set Initial Open Price
        Open_Price = Cells(2, 3).Value
         'Loop through all ticker symbol
        
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker_Name = Cells(i, 1).Value
                
                Cells(Row, 9).Value = Ticker_Name
                
                Close_Price = Cells(i, 6).Value
                
                Yearly_Change = Close_Price - Open_Price
                
                Cells(Row, 10).Value = Yearly_Change
                
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                'Add Total Volume
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                
                Row = Row + 1
        
                Open_Price = Cells(i + 1, 3)
                ' reset total volume
                Volume = 0
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        'conditional formating/coloring
        For j = 2 To YCLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        'Title headers 'Columns giving me a tricky time, define them above

        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"

        ' Look through each rows 'Lets use Z as are noew placeholder
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
       
        
    Next WS
        
End Sub

            
    




