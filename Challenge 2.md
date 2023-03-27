
Sub stockloop()

    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Integer
    Dim increase_number As Double
    Dim decrease_number As Double
    Dim volume_number As Long
    

    For Each ws In Worksheets
  
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
  
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Summary_Table_Row = 2

        For i = 2 To lastRow
    
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                yearly_open = ws.Cells(i, 3).Value
    
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
                yearly_close = ws.Cells(i, 6).Value
                yearly_change = yearly_close - yearly_open
                percent_change = yearly_change / yearly_open
       
                ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1).Value
                ws.Range("L" & Summary_Table_Row).Value = Total
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
                Summary_Table_Row = Summary_Table_Row + 1
                Total = 0
    
            Else
                Total = Total + ws.Cells(i, 7).Value
    
             End If


            If ws.Cells(Summary_Table_Row, 10).Value <= 0 Then
    
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

            Else
    
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

            End If

        Next i

        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    

        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
    
        ws.Range("O1:Q4").Interior.ColorIndex = 6
    
    Next ws
  
End Sub
