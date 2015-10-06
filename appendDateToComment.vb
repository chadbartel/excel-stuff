' This macro appends to a comment on a cell that has
' "Emailed" and add the current date on a new line.
Sub appendDateToComment()

' Declare our variables
Dim r As Range
Dim c As Range
Dim sCommentCurrent As String
Dim sCommentAdd As String
Dim sCommentNew As String

' Store the selected cells in a var
Set r = Range(ActiveCell, ActiveCell.End(xlDown))

' Set the new "Emailed" comment string
sCommentNew = "Emailed:" & vbNewLine & VBA.DateTime.Month(Now) _
        & "/" & VBA.DateTime.Day(Now) & "/" & VBA.DateTime.Year(Now)

' Loop through selected range
For Each c In r
    ' The added comment won't change, this should save space
    sCommentAdd = vbNewLine & VBA.DateTime.Month(Now) & "/" _
        & VBA.DateTime.Day(Now) & "/" & VBA.DateTime.Year(Now)
        
    On Error Resume Next
    sCommentCurrent = c.Comment.Text
    sCommentAdd = sCommentCurrent & sCommentAdd
    c.Comment.Text Text:=sCommentAdd
    
    If Err.Number = 91 Then
        Err.Clear
        c.AddComment
        c.Comment.Text sCommentNew
    End If
    
    c.Comment.Shape.TextFrame.AutoSize = True
    
Next c

End Sub
