Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Überprüfen, ob sich die relevanten Zellen geändert haben
    If Not Intersect(Target, Sh.Range("F2,F4,F6,I6")) Is Nothing Then
        Call SetHyperlinkAGZeichnungforSheet(ActiveSheet)
    End If

End Sub