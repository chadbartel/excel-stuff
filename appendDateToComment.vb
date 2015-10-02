' This macro appends to a comment on a cell that has
' "Emailed" and add the current date on a new line.
Sub appendDateToComment()

' Declare our variables
Dim ws As Worksheet
Dim lr As Long
Dim oldCom As Comment
Dim newCom As Comment
Dim chk As Long
Dim firstAddress As Long

' Set out variables equal to their values
Set ws = ThisWorkbook.Sheets("Comment Update Sheet")
' This is looking for the last row in column "I"
Set lr = ws.Cells(Rows.Count, 9).End(xlUp).Row

With ws.Range("I2:I" & lr)
    Set chk = .Find("Emailed", LookIn:=xlWhole, SearchOrder:=xlByRows, _
        MatchCase:=True)
    ' If chk finds a match...
    If Not chk Is Nothing Then
        firstAddress = chk.Address
        Do
            oldCom = chk.Comment
            newCom = oldCom.Text(vbNewLine & VBA.DateTime.Month(Now) _
                & "/" & VBA.DateTime.Day(Now) & "/" & VBA.DateTime.Year(Now))
            chk.Comment = newCom
        Loop While Not chk Is Nothing And chk.Address <> firstAddress
    End If
End With

End Sub
