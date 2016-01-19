Sub MACD_Sheet1_Copy()

' Declare MACD workbook variable
Dim macdwb As Workbook
Set macdwb = ThisWorkbook

' Declare sheet where we're given SNs
Dim emails As Worksheet
Set emails = macdwb.Sheets("Sheet1")

' Declare variables to look for in MP8032 sheet
Dim faultyemail As Range
Dim replaceemail As Range
Dim faultrow As Long
Dim replacerow As Long

Set faultyemail = emails.Range("B7")
Set replaceemail = emails.Range("B8")

Dim mp8032s As Worksheet
Set mp8032s = macdwb.Sheets("MP8032")

' Declare column containing SN values
Dim snrange As Range
Set snrange = mp8032s.Range("AG:AG")

' Declare worksheet where we're copying data
Dim uploadws As Worksheet
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))
    
' Find the row of each SN we're looking for in MP8032
With snrange
    Set f = .Find(faultyemail.Value, LookIn:=xlValues)
    Set r = .Find(replaceemail.Value, LookIn:=xlValues)
    If Not f Is Nothing Then
        faultrow = f.Row
        mp8032s.Range("M" & faultrow & ":CH" & faultrow).Copy
        lastRow = uploadws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        uploadws.Range("A" & lastRow).PasteSpecial xlPasteValues
    End If
    If Not r Is Nothing Then
        replacerow = r.Row
        mp8032s.Range("M" & replacerow & ":CH" & replacerow).Copy
        lastRow = uploadws.Cells(Rows.Count, 1).End(xlUp).Row + 1
        uploadws.Range("A" & lastRow).PasteSpecial xlPasteValues
    End If
End With

Application.Run ("MACD_Switch_Loc_Details")

End Sub




Sub MACD_Switch_Loc_Details()


' Declare MACD workbook variable
Dim macdwb As Workbook
Set macdwb = ThisWorkbook

' Declare worksheet where we're copying data
Dim uploadws As Worksheet
Set uploadws = macdwb.Worksheets("Upload " & _
    VBA.DateTime.Month(Now) & "-" & VBA.DateTime.Day(Now))

' Get last row from sheet
Dim lastRow As Long
lastRow = uploadws.Cells(Rows.Count, 1).End(xlUp).Row

' Switch faulty printer location details with replacement
uploadws.Range("D" & lastRow - 1 & ":G" & lastRow - 1).Copy _
    uploadws.Range("D" & lastRow + 1)
uploadws.Range("D" & lastRow & ":G" & lastRow + 1).Copy _
    uploadws.Range("D" & lastRow - 1)
uploadws.Range("D" & lastRow + 1 & ":G" & lastRow + 1).ClearContents

uploadws.Range("O" & lastRow - 1 & ":R" & lastRow - 1).Copy _
    uploadws.Range("O" & lastRow + 1)
uploadws.Range("O" & lastRow & ":R" & lastRow + 1).Copy _
    uploadws.Range("O" & lastRow - 1)
uploadws.Range("O" & lastRow + 1 & ":R" & lastRow + 1).ClearContents

uploadws.Range("AE" & lastRow - 1 & ":AM" & lastRow - 1).Copy _
    uploadws.Range("AE" & lastRow + 1)
uploadws.Range("AE" & lastRow & ":AM" & lastRow + 1).Copy _
    uploadws.Range("AE" & lastRow - 1)
uploadws.Range("AE" & lastRow + 1 & ":AM" & lastRow + 1).ClearContents

End Sub
