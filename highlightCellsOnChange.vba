Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim lr As Long
    Dim rc As String
    Dim keyCells1 As Range
    Dim keyCells2 As Range
    
    lr = Cells(Rows.Count, "F").End(xlUp).Row
    Set keyCells1 = Range("F:F")
    Set keyCells2 = Range("G:G")
    
    If Not Application.Intersect(keyCells1, Range(Target.Address)) Is Nothing Then
        lr = Cells(Rows.Count, "F").End(xlUp).Row
        Range("A" & lr & ":F" & lr).Interior.Color = RGB(204, 255, 255)
    ElseIf Not Application.Intersect(keyCells2, Range(Target.Address)) Is Nothing Then
        rc = Target.Row
        Range("H" & rc & ":L" & rc).Interior.Color = vbYellow
    End If
    
End Sub
