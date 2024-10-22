Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Sub UpdateHyperlinksInAGSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If Left(ws.Name, 2) = "AG" Then
            SetHyperlinkAGZeichnungforSheet ws
        End If
    Next ws
End Sub

Sub SetHyperlinkAGZeichnungforSheet(ws As Worksheet)
    Dim basePath As String
    Dim articleFolder As String
    Dim firstOptionPath As String
    Dim secondOptionPath As String
    Dim fallbackPath As String
    Dim fileName As String
    Dim info2 As String
    Dim artikelNummer As String
    Dim errMsg As String
    
    ' Basis-Pfad setzen
    basePath = "\\MS01\Datenpfad\Betriebsorganisation\Fertigungsdaten\"
    
    ' Fehlerüberprüfung für notwendige Felder
    errMsg = ""
    If ws.Range("F2").Value = "" Then errMsg = errMsg & "Feld F2 auf " & ws.Name & "; "
    If ws.Range("F6").Value = "" Then errMsg = errMsg & "Feld F6 auf " & ws.Name & "; "
    If ws.Range("I6").Value = "" Then errMsg = errMsg & "Feld I6 auf " & ws.Name & "; "
    
    ' Falls Daten fehlen, zeige eine Warnmeldung und beende die Subroutine
    If errMsg <> "" Then
        MsgBox "Erforderliche Zelleninformationen fehlen auf dem Blatt " & ws.Name & ": " & errMsg, vbExclamation, "Datenfehler"
        Exit Sub
    End If

    ' Info2 und Artikelnummer aus den entsprechenden Arbeitsblättern holen
    info2 = ThisWorkbook.Sheets("Stammdaten").Range("B17").Value
    artikelNummer = ws.Range("F2").Value

    ' Stamm-Pfad zusammenstellen basierend auf dem ersten Buchstaben von Info2, dem gesamten Info2-Wert und der Artikelnummer
    articleFolder = Left(info2, 1) & "\" & info2 & "\" & artikelNummer & "\"

    ' Dateiname basierend auf den Feldern F2, F6 und I6
    fileName = artikelNummer & "-" & ws.Range("F6").Value & "-AG" & ws.Range("I6").Value & ".pdf"

    ' Erster Option-Pfad
    firstOptionPath = basePath & articleFolder & "Zeichnungsdaten\" & fileName

    ' Zweiter Option-Pfad (Fallback, falls erster Pfad nicht existiert)
    secondOptionPath = basePath & articleFolder & "Zeichnungsdaten\" & artikelNummer & "-" & ws.Range("F6").Value & ".pdf"

    ' Dritter Option-Pfad (fallback auf JPG-Pfad)
    fallbackPath = "\\MS01\Datenpfad\Fauser\Zeichnungen\" & artikelNummer & ".jpg"

    ' Überprüfen, ob die Datei existiert und den Hyperlink in Zelle I7 setzen
    If FileExists(firstOptionPath) Then
        ws.Range("I7").Hyperlinks.Add Anchor:=ws.Range("I7"), Address:=firstOptionPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    ElseIf FileExists(secondOptionPath) Then
        ws.Range("I7").Hyperlinks.Add Anchor:=ws.Range("I7"), Address:=secondOptionPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    Else
        ws.Range("I7").Hyperlinks.Add Anchor:=ws.Range("I7"), Address:=fallbackPath, TextToDisplay:="Arbeitsgang-Zeichnung"
    End If
End Sub

