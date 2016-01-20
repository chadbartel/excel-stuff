Sub updatePivotTables()

Dim PT As PivotTable
Dim WS As Worksheet

  For Each WS In ThisWorkbook.Worksheets
    For Each PT In WS.PivotTables
      On Error Resume Next
      PT.RefreshTable
    Next PT
  Next WS

  MsgBox ("All Pivot Tables refreshed!")
  
End Sub
