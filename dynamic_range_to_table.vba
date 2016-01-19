Dim rng As Range

    Set rng = Range("A1").CurrentRegion 'Change to match your range
    ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "MyTable"
    

' or


Sub A_SelectAllMakeTable()
    Dim tbl As ListObject
    Dim rng As Range

    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium15"
End Sub
