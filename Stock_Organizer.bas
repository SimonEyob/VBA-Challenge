Attribute VB_Name = "Module1"
Sub StockOrgainzer()
   For Each ws In Worksheets
    Dim LastRow, LastColumn, Vol_Total, Summary_Table_Row, EndYear, BeginYear As Double
    Dim Price_diff_percentage As Double
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    Vol_Total = 0
    Summary_Table_Row = 2
    
    Counter = 0
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    
    
    For i = 2 To LastRow
      
    If ws.Cells(i + 1, 1) = ws.Cells(i, 1) And Counter = 0 Then

      ' Add to the Brand Total
     Vol_Total = Vol_Total + ws.Cells(i, 7).Value
     
     BeginYear = ws.Cells(i, 3)
     Counter = Counter + 1
     
     ElseIf ws.Cells(i + 1, 1) = ws.Cells(i, 1) Then

      ' Add to the Brand Total
     Vol_Total = Vol_Total + ws.Cells(i, 7).Value
     
     Counter = Counter + 1
    
  
  ElseIf ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Volume Total
      Vol_Total = Vol_Total + ws.Cells(i, 7)
      
      'Store the end year price and percentage'
      EndYear = ws.Cells(i, 6)
      
      Price_diff = EndYear - BeginYear
      
      If Price_diff = 0 Or BeginYear = 0 Then
      Price_diff_percentage = 0
      
      Else
      Price_diff_percentage = Round(Price_diff, 3) / (BeginYear)
      
      End If
      

      ' Print the Credit Card Brand in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ws.Range("J" & Summary_Table_Row).Value = Price_diff
      ws.Range("K" & Summary_Table_Row).Value = Price_diff_percentage
      

      ' Print the Volume stock Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Vol_Total
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Vol_Total = 0
      Counter = 0


    ' If the cell immediately following a row is the same brand...
            End If
    Next i
    
    NewLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            
    For j = 2 To NewLastRow
         
         ws.Cells(j, 11).Style = "Percent"
        
        If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(j, 10) < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
        
            End If
        
       
    Next j
    NewLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    
    ws.Range("P2:P3").Style = "Percent"
    ws.Range("P2") = WorksheetFunction.Max(ws.Range("K2:K" & NewLastRow))
    ws.Range("P3") = WorksheetFunction.Min(ws.Range("K2:K" & NewLastRow))
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & NewLastRow))
    
    TickerRow_Max_Increase = WorksheetFunction.Match(ws.Range("P2"), ws.Range("K2:K" & NewLastRow), 0)
    ws.Range("O2") = ws.Cells(TickerRow_Max_Increase + 1, 9)
    
   TickerRow_Max_Decrease = WorksheetFunction.Match(ws.Range("P3"), ws.Range("K2:K" & NewLastRow), 0)
    ws.Range("O3") = ws.Cells(TickerRow_Max_Decrease + 1, 9)
    
    TickerRow_Max = WorksheetFunction.Match(ws.Range("P4"), ws.Range("L2:L" & NewLastRow), 0)
    ws.Range("O4") = ws.Cells(TickerRow_Max + 1, 9)
    
 
    
    
    
Next ws
End Sub
