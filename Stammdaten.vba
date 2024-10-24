Private Sub Worksheet_Change(ByVal Target As Range)
    ' Überprüfe, ob die Änderung in Zelle B3 des Arbeitsblatts "Stammdaten" erfolgte
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        ' Rufe die Funktion zur Aktualisierung der Daten auf
        RefreshData Me.Range("B3").Value
    End If
End Sub

Sub RefreshData(orderNumber As String)
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim zeichnungsNummer As String
    Dim lastChar As String
    Set ws = ThisWorkbook.Sheets("Stammdaten")
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    ' Verbindungszeichenfolge unter Verwendung der SQL-Server-Authentifizierung
    conn.Open "Provider=SQLOLEDB;Data Source=MS01\mothessql;Initial Catalog=ISDATA;User ID=sa;Password=sa;"

    ' Lösche den Inhalt der relevanten Zellen im Arbeitsblatt "Stammdaten"
    ws.Range("B5:C20").ClearContents

    ' Setze die SQL-Abfrage mit der neuen Auftragsnummer
    sql = "SELECT fag.TXT05 AS Hauptordner, " & _
          "OR_ORDER.NAME AS Auftragsnummer, " & _
          "OR_PROJ.NAME AS Projekt, " & _
          "OR_ORDER.DESCR AS Bezeichnung, " & _
          "OR_ORDER.ARTNO AS Artikelnummer, " & _
          "OR_ORDER.DRAWNO AS Zeichnungsnummer, " & _
          "OR_ORDER.DRAWIND AS [Index], " & _
          "OR_ORDER.INFO1 AS Werkstoff, " & _
          "OR_ORDER.TYPE AS Fertigungstyp, " & _
          "OR_ORDER.DELIVERY AS Liefertermin, " & _
          "OR_ORDER.IDENT AS Teilenummer, " & _
          "OR_ORDER.PPARTS AS Sollstückzahl, " & _
          "OR_ORDER.COMMNO AS Materialeingangsnummer, " & _
          "CU_COMP.NAME AS Kunde, " & _
          "CU_COMP.INFO2 AS Info2 " & _
          "FROM OR_ORDER " & _
          "LEFT JOIN PA_POSIT pos ON (pos.POSTNAME = OR_ORDER.NAME) " & _
          "LEFT JOIN PA_PAPER pap ON (pap.PANO = pos.PANO) " & _
          "LEFT OUTER JOIN fag_detail fag ON fag.FKNO = pap.PANO AND fag.TYP = 3 " & _
          "LEFT JOIN CU_COMP ON OR_ORDER.KCONO = CU_COMP.CONO " & _
          "LEFT JOIN OR_PROJ ON OR_ORDER.PRONO = OR_PROJ.NO " & _
          "WHERE OR_ORDER.NAME = '" & orderNumber & "' "

    ' Öffne ein Recordset
    rs.Open sql, conn, 1, 3  ' adOpenKeyset, adLockOptimistic

    ' Stelle sicher, dass Daten vorhanden sind
    If Not rs.EOF Then
        ws.Range("B5").Value = rs.Fields("Auftragsnummer").Value
        ws.Range("B6").Value = rs.Fields("Projekt").Value
        ws.Range("B7").Value = rs.Fields("Bezeichnung").Value
        ws.Range("B8").Value = rs.Fields("Teilenummer").Value
        ws.Range("B9").Value = rs.Fields("Artikelnummer").Value
        zeichnungsNummer = rs.Fields("Zeichnungsnummer").Value
        
        If IsNull(rs.Fields("Index").Value) Or rs.Fields("Index").Value = "" Then
            lastChar = Right(zeichnungsNummer, 1)
            ws.Range("B11").Value = lastChar  ' Setze das letzte Zeichen der Zeichnungsnummer als Index
            If InStr(zeichnungsNummer, " ") > 0 Then
                zeichnungsNummer = Left(zeichnungsNummer, InStr(1, zeichnungsNummer, " ") - 1)
            End If
        Else
            ws.Range("B11").Value = rs.Fields("Index").Value
        End If
        
        ws.Range("B10").Value = zeichnungsNummer
        ws.Range("B12").Value = rs.Fields("Werkstoff").Value
        ws.Range("C12").Value = rs.Fields("Materialeingangsnummer").Value
        ws.Range("B13").Value = rs.Fields("Fertigungstyp").Value
        ws.Range("B14").Value = rs.Fields("Liefertermin").Value
        ws.Range("B15").Value = rs.Fields("Sollstückzahl").Value
        ws.Range("B16").Value = rs.Fields("Kunde").Value
        ws.Range("B17").Value = rs.Fields("Info2").Value
        ws.Range("B20").Value = rs.Fields("Hauptordner").Value
        

        ' Berechne den Artikelordner-Pfad
        Dim basePath As String
        Dim info2 As String
        Dim artikelNummer As String
        basePath = "\\MS01\Datenpfad\Betriebsorganisation\Fertigungsdaten\"
        info2 = ws.Range("B17").Value
        artikelNummer = ws.Range("B9").Value
        Dim articleFolder As String
        articleFolder = basePath & Left(info2, 1) & "\" & info2 & "\" & artikelNummer & "\"
        ws.Range("B19").Value = articleFolder
        
        Call UpdateHyperlinksInAGSheets
    End If

    ' Schließe das Recordset und die Verbindung
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


