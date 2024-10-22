Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Sub SetHyperlinkAGZeichnung()
    Dim basePath As String
    Dim articleFolder As String
    Dim firstOptionPath As String
    Dim secondOptionPath As String
    Dim fallbackPath As String
    Dim fileName As String
    Dim cell As Range
    
    ' Stamm-Pfad zusammenstellen
    basePath = "\\MS01\Datenpfad\Betriebsorganisation\Fertigungsdaten\"
    articleFolder = Left(ActiveSheet.Range("F4").Value, 1) & "\" & ActiveSheet.Range("F4").Value & "\" & ActiveSheet.Range("F2").Value & "\"
    
    ' Dateiname basierend auf den Feldern F2, F6 und I6
    fileName = ActiveSheet.Range("F2").Value & "-" & ActiveSheet.Range("F6").Value & "-AG" & ActiveSheet.Range("I6").Value & ".pdf"
    
    ' Erster Option-Pfad
    firstOptionPath = basePath & articleFolder & "Zeichnungsdaten\" & fileName
    
    ' Zweiter Option-Pfad (Fallback, falls erster Pfad nicht existiert)
    secondOptionPath = basePath & articleFolder & "Zeichnungsdaten\" & ActiveSheet.Range("F2").Value & "-" & ActiveSheet.Range("F6").Value & ".pdf"
    
    ' Dritter Option-Pfad (fallback auf JPG-Pfad)
    fallbackPath = "\\MS01\Datenpfad\Fauser\Zeichnungen\" & ActiveSheet.Range("F2").Value & ".jpg"
    
    ' Überprüfen, ob die Datei existiert und den Hyperlink in Zelle I7 setzen
    If FileExists(firstOptionPath) Then
        ActiveSheet.Range("I7").Hyperlinks.Add Anchor:=ActiveSheet.Range("I7"), Address:=firstOptionPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    ElseIf FileExists(secondOptionPath) Then
        ActiveSheet.Range("I7").Hyperlinks.Add Anchor:=ActiveSheet.Range("I7"), Address:=secondOptionPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    Else
        ActiveSheet.Range("I7").Hyperlinks.Add Anchor:=ActiveSheet.Range("I7"), Address:=fallbackPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    End If
End Sub