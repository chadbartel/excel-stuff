Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim lr As Long
    Dim keyCells As Range
    
    lr = Cells(Rows.Count, "F").End(xlUp).Row
    Set keyCells = Range("F:F")
    
    If Not Application.Intersect(keyCells, Range(Target.Address)) Is Nothing Then
        lr = Cells(Rows.Count, "F").End(xlUp).Row
        Range("A" & lr & ":F" & lr).Interior.Color = RGB(204, 255, 255)
    End If
    
End Sub
