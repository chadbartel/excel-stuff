Sub MACD_data_validation_v1()

' Declare MACD workbook variable
Dim macdwb As Workbook
Set macdwb = ThisWorkbook

' Declare worksheets on which data validation will be performed
Dim uploadws As Worksheet
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
Dim locationws As Worksheet
Set locationws = macdwb.Worksheets("Locations")

'Since the size of the location worksheet doesn't change...
lastRowLoc = 488

' Declare the ranges you will be validating
Dim snrange As Range
lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
Set snrange = uploadws.Range("U2:U" & lastRowSN)

' Declare static comment strings
Dim sCommentCurrent As String
Dim sCommentNew As String

' Create static values for static comment strings
sCommentNew = "Data Validation Message:" & vbNewLine

' Call each data validation script
Application.Run ("Deactivated_SN_validate")
Application.Run ("Duplicate_SN_validate")
Application.Run ("Discrepancy_DG_validate")
Application.Run ("Discrepancy_DO_validate")
Application.Run ("Discrepancy_DR_validate")

End Sub


Sub Deactivated_SN_validate()

Dim macdwb As Workbook
Dim decomws As Worksheet
Dim lastRowDecom As Long
Dim uploadws As Worksheet
Dim lastRowSN As Long
Dim snrange As Range
Dim sCommentCurrent As String
Dim sCommentDecom As String
Dim i As Integer
Dim j As Integer

Set macdwb = ThisWorkbook
Set decomws = macdwb.Worksheets("Deactivated")
lastRowDecom = decomws.Cells(Rows.Count, 21).End(xlUp).Row
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
Set snrange = uploadws.Range("U2:U" & lastRowSN)
sCommentDecom = "This device has been decommissioned."

For i = 2 To lastRowSN
    cellVal = uploadws.Range("U" & i).Value
    For j = 2 To lastRowDecom
        If cellVal = decomws.Cells(j, 21).Value Then
            If uploadws.Range("U" & i).Comment Is Nothing Then
                With uploadws.Range("U" & i).AddComment
                    .Text sCommentDecom
                    .Shape.TextFrame.AutoSize = True
                End With
            Else
                sCommentDecom = uploadws.Range("U" & i).Comment.Text & _
                    vbNewLine & sCommentDecom
                uploadws.Range("U" & i).Comment.Text sCommentDecom
                uploadws.Range("U" & i).Comment.Shape.TextFrame.AutoSize = True
            End If
        End If
    Next j
Next i

End Sub


Sub Duplicate_SN_validate()

Dim macdwb As Workbook
Dim uploadws As Worksheet
Dim lastRowSN As Long
Dim snrange As Range
Dim sCommentCurrent As String
Dim sCommentDupe As String
Dim countvar As Long
Dim i As Integer

Set macdwb = ThisWorkbook
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
Set snrange = uploadws.Range("U2:U" & lastRowSN)
sCommentDupe = "This SN has already appeared in your upload."

For i = 2 To lastRowSN
    countvar = Application.WorksheetFunction.CountIf(snrange, _
        uploadws.Range("U" & i))
    If countvar > 1 Then
        If uploadws.Range("U" & i).Comment Is Nothing Then
            With uploadws.Range("U" & i).AddComment
                .Text sCommentDupe
                .Shape.TextFrame.AutoSize = True
            End With
        Else
            sCommentDupe = uploadws.Range("U" & i).Comment.Text & _
                vbNewLine & sCommentDupe
            uploadws.Range("U" & i).Comment.Text Text:=sCommentDupe
            uploadws.Range("U" & i).Comment.Shape.TextFrame.AutoSize = True
        End If
    End If
Next i

End Sub


Sub Discrepancy_DG_validate()

Dim macdwb As Workbook
Dim uploadws As Worksheet
Dim lastRowSN As Long
Dim locationws As Worksheet
Dim lastRowLoc As Long
Dim sCommentCurrent As String
Dim sCommentDG As String
Dim i As Integer
Dim j As Integer

Set macdwb = ThisWorkbook
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
Set locationws = macdwb.Worksheets("Locations")
lastRowLoc = locationws.Cells(Rows.Count, 15).End(xlUp).Row
lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
sCommentDG = "Customer Name (G) doesn't match Address Name (D)."

For i = 2 To lastRowSN
    cellVal = uploadws.Range("D" & i).Value
    For j = 2 To lastRowLoc
        On Error Resume Next
        If cellVal = locationws.Range("N" & j).Value And _
            uploadws.Range("G" & i).Value <> locationws.Range("B" & j).Value Then
            If uploadws.Range("U" & i).Comment Is Nothing Then
                With uploadws.Range("U" & i).AddComment
                    .Text sCommentDG
                    .Shape.TextFrame.AutoSize = True
                End With
            Else
                sCommentDG = uploadws.Range("U" & i).Comment.Text & _
                    vbNewLine & sCommentDG
                uploadws.Range("U" & i).Comment.Text Text:=sCommentDG
                uploadws.Range("U" & i).Comment.Shape.TextFrame.AutoSize = True
            End If
        End If
    Next j
Next i

End Sub


'Sub Discrepancy_DO_validate()
'
'Dim macdwb As Workbook
'Dim uploadws As Worksheet
'Dim locationws As Worksheet
'Dim lastRowLoc As Long
'Dim sCommentCurrent As String
'Dim sCommentDO As String
'Dim i As Integer
'Dim j As Integer
'
'Set macdwb = ThisWorkbook
'Set uploadws = macdwb.Worksheets("Upload " & _
'    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
'Set locationws = macdwb.Worksheets("Locations")
'lastRowLoc = locationws.Cells(Rows.Count, 15).End(xlUp).Row
'lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
'sCommentDO = "Contact User ID (O) doesn't match Address Name (D)."
'
'For i = 2 To lastRowSN
'    cellVal = uploadws.Range("D" & i).Value
'    contactid = uploadws.Range("O" & i).Value
'    For j = 2 To lastRowLoc
'        On Error Resume Next
'        If cellVal = locationws.Range("N" & j).Value And _
'            contactid <> locationws.Range("O" & j).Value Then
'            If uploadws.Range("U" & i).Comment Is Nothing Then
'                With uploadws.Range("U" & i).AddComment
'                    .Text sCommentDO
'                    .Shape.TextFrame.AutoSize = True
'                End With
'            Else
'                sCommentDG = uploadws.Range("U" & i).Comment.Text & _
'                    vbNewLine & sCommentDO
'                uploadws.Range("U" & i).Comment.Text Text:=sCommentDO
'                uploadws.Range("U" & i).Comment.Shape.TextFrame.AutoSize = True
'            End If
'        End If
'    Next j
'Next i
'
'End Sub


Sub Discrepancy_DR_validate()

Dim macdwb As Workbook
Dim uploadws As Worksheet
Dim lastRowSN As Long
Dim locationws As Worksheet
Dim lastRowLoc As Long
Dim sCommentCurrent As String
Dim sCommentDR As String
Dim i As Integer
Dim j As Integer

Set macdwb = ThisWorkbook
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
Set locationws = macdwb.Worksheets("Locations")
lastRowLoc = locationws.Cells(Rows.Count, 15).End(xlUp).Row
lastRowSN = uploadws.Cells(Rows.Count, 21).End(xlUp).Row
sCommentDR = "Data discrepancy between column D and R."

For i = 2 To lastRowSN
    cellVal = uploadws.Range("D" & i).Value
    costcenter = uploadws.Range("R" & i).Value
    For j = 2 To lastRowLoc
        On Error Resume Next
        If cellVal = locationws.Range("N" & j).Value And _
            costcenter <> locationws.Range("M" & j).Value Then
            If uploadws.Range("U" & i).Comment Is Nothing Then
                With uploadws.Range("U" & i).AddComment
                    .Text sCommentDR
                    .Shape.TextFrame.AutoSize = True
                End With
            Else
                sCommentDR = uploadws.Range("U" & i).Comment.Text & _
                    vbNewLine & sCommentDR
                uploadws.Range("U" & i).Comment.Text Text:=sCommentDR
                uploadws.Range("U" & i).Comment.Shape.TextFrame.AutoSize = True
            End If
        End If
    Next j
Next i

End Sub
