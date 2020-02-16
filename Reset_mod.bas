Attribute VB_Name = "Module2"
Sub Clearmod()
For Each ws In Worksheets
NewLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
ws.Range("I1:P1" & NewLastRow).Clear
Next ws

End Sub
