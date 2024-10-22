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

    ' Setze die SQL-Abfrage mit der neuen Auftragsnummer
    sql = "SELECT OR_ORDER.NAME AS Auftragsnummer, " & _
          "OR_ORDER.PRONO AS Projekt, " & _
          "OR_ORDER.DESCR AS Bezeichnung, " & _
          "OR_ORDER.ARTNO AS Artikelnummer, " & _
          "OR_ORDER.DRAWNO AS Zeichnungsnummer, " & _
          "OR_ORDER.DRAWIND AS [Index], " & _
          "OR_ORDER.INFO1 AS Werkstoff, " & _
          "OR_ORDER.TYPE AS Fertigungstyp, " & _
          "OR_ORDER.DELIVERY AS Liefertermin, " & _
          "OR_ORDER.IDENT AS Teilenummer, " & _
          "OR_ORDER.PPARTS AS Sollstückzahl, " & _
          "CU_COMP.NAME AS Kunde, " & _
          "CU_COMP.INFO2 AS Info2 " & _
          "FROM OR_ORDER " & _
          "LEFT JOIN CU_COMP ON OR_ORDER.KCONO = CU_COMP.CONO " & _
          "WHERE OR_ORDER.NAME = '" & orderNumber & "';"

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
        ws.Range("B13").Value = rs.Fields("Fertigungstyp").Value
        ws.Range("B14").Value = rs.Fields("Liefertermin").Value
        ws.Range("B15").Value = rs.Fields("Sollstückzahl").Value
        ws.Range("B16").Value = rs.Fields("Kunde").Value
        ws.Range("B17").Value = rs.Fields("Info2").Value
    End If

    ' Schließe das Recordset und die Verbindung
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
