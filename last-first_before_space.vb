Function FirstBeforeChar(ByVal Cell As Range, Optional ByVal Char As String = " ")

    FirstBeforeChar = Trim(Left(Application.WorksheetFunction.Substitute(Cell, Char, Application.WorksheetFunction.Rept(Char, Len(Cell))), Len(Cell)))

End Function


Function LastBeforeChar(ByVal Cell As Range, Optional ByVal Char As String = " ")

    LastBeforeChar = Trim(Right(Application.WorksheetFunction.Substitute(Cell, Char, Application.WorksheetFunction.Rept(Char, Len(Cell))), Len(Cell)))

End Function
