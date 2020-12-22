VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLiefergegenstand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLiefergegenstand

Option Compare Database
Option Explicit

Private mrstRecordset As DAO.RecordSet
Private mbolVerbatim As Boolean

' dient als Zwischenspeicher für die ID des aktuell
' mit dem Cursor ausgewählten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarLiefergegenstandClipboard As Variant

    Private mvarLiefergegenstandID As Variant
    Private mvarAftrID As Variant
    Private mvarBezeichnung As Variant
    Private mvarLieferscheinLink As Variant
    Private mvarSeriennummer As Variant
    Private mvarAnzahl As Variant
    Private mvarHerstellerkennzeichen As Variant
    Private mvarUnterauftragnehmerAngebotLink As Variant
    Private mvarAngebotNetto As Variant
    Private mvarPreisBrutto As Variant
    Private mvarLieferdatum As Variant
    Private mvarLieferschein As Variant
    Private mvarBemerkung As Variant
    Private mvarZielAftrID As Variant
    Private mvarZielLieferschein As Variant
    Private mvarZielLinklieferschein As Variant
    Private mvarZielLieferdatum As Variant
    Private mvarLiefergegenstandLink As Variant
    Private mvarIstReserve As Variant
    Private mvarAftrTitel As Variant
    Private mvarBWIKey As Variant
    Private mvarLeistungsbeschreibungLink As Variant
    Private mvarParentKey As Variant

Public Sub Default()
    mvarLiefergegenstandID = Null
    mvarAftrID = Null
    mvarBezeichnung = Null
    mvarLieferscheinLink = Null
    mvarSeriennummer = Null
    mvarAnzahl = Null
    mvarHerstellerkennzeichen = Null
    mvarUnterauftragnehmerAngebotLink = Null
    mvarAngebotNetto = Null
    mvarPreisBrutto = Null
    mvarLieferdatum = Null
    mvarLieferschein = Null
    mvarBemerkung = Null
    mvarZielAftrID = Null
    mvarZielLieferschein = Null
    mvarZielLinklieferschein = Null
    mvarLieferdatum = Null
    mvarLiefergegenstandLink = Null
    mvarIstReserve = Null
    mvarAftrTitel = Null
    mvarBWIKey = Null
    mvarLeistungsbeschreibungLink = Null
    mvarParentKey = Null
End Sub

' wird im Unterformular ausgeführt, lädt den aktuell ausgewählten
' Datensatz ins Objekt
' ist Abhängig von der Abfrage qryLiefergegenstand und
' der Variablen mvarLiefergegenstandClipboard
Public Sub Laden(ByVal varID As Variant)
    Debug.Print "clsLiefergegenstand.Laden ausführen"
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb

    ' das Objekt in den Default-Zustand versetzen, wenn im Unterformular
    ' kein Datensatz mit dem Cursor ausgewählt wurde,
    ' zu ExitProc springen
    If IsEmpty(varID) Or IsNull(varID) Then
        mvarLiefergegenstandClipboard = Null
        Default
        Debug.Print "clsLiefergegenstand.Laden: varID ist leer, " _
            + "Datensatz Defaultwert zugewiesen"
        GoTo ExitProc
    End If
    
    ' der Zwischenspeichervariable 'mvarLiefergegenstandClipboard' den ausgewählten
    ' Datensatz zuweisen - wird für die Prozedur Speichern benötigt
    mvarLiefergegenstandClipboard = varID
    ' Debug.Print "clsLiefergegenstand.Laden: mvarLiefergegenstandClipboard = " & mvarLiefergegenstandClipboard
        
        ' zu durchsuchende Zielabfrage benennen
        Dim strAbfrageTitel As String
        strAbfrageTitel = "qryLiefergegenstand"
        
        ' zu durchsuchendes Zielfeld benennen
        Dim strSucheZiel As String
        strSucheZiel = "LiefergegenstandID"
        
        ' prüfen ob zu durchsuchende Abfrage existiert,
        ' wenn nicht: goto ExitProc

        Dim RecordSet As Object
        Dim bolQueryExistiert As Boolean
        bolQueryExistiert = False
    
        For Each RecordSet In dbsCurrentDB.QueryDefs
            If RecordSet.Name = strAbfrageTitel Then
                bolQueryExistiert = True
            End If
        Next RecordSet
        
        If bolQueryExistiert = False Then
            GoTo ExitProc
        End If
        
        ' prüfen ob varID null ist, wenn ja gehe zu ExitProc
        If IsNull(varID) Or IsEmpty(varID) Then
            Debug.Print "varID ist leer, Prozedur 'AuftragLaden' abgebrochen"
            Default
            GoTo ExitProc
        End If
    
        ' Zielfeld in Zielabfrage nach varID durchsuchen
        Set mrstRecordset = dbsCurrentDB.OpenRecordset(strAbfrageTitel, dbOpenDynaset)
        mrstRecordset.FindFirst (strSucheZiel & " = " & CLng(varID))

    ' Attribute von mrstRecordset in Objekt laden
    With mrstRecordset
        mvarLiefergegenstandID = !LiefergegenstandID
        mvarAftrID = !aftrIdKey
        mvarBezeichnung = !bezeichnung
        mvarLieferscheinLink = !linkLieferschein
        mvarSeriennummer = !seriennummer
        mvarAnzahl = !anzahl
        mvarHerstellerkennzeichen = !herstellerkennzeichen
        mvarUnterauftragnehmerAngebotLink = !Uanangebot
        mvarAngebotNetto = !angebotNetto
        mvarPreisBrutto = !preisBrutto
        mvarLieferdatum = !lieferdatum
        mvarLieferschein = !lieferschein
        mvarBemerkung = !Bemerkung
        mvarZielAftrID = !zielAftrID
        mvarZielLieferschein = !zielLieferschein
        mvarZielLinklieferschein = !zielLinkLieferschein
        mvarZielLieferdatum = !zielLieferdatum
        mvarLiefergegenstandLink = !LiefergegenstandLink
        mvarIstReserve = !IstReserve
        mvarAftrTitel = !AftrTitel
        mvarBWIKey = !BWIKey
        mvarLeistungsbeschreibungLink = !Leistungsbeschreibung
        mvarParentKey = !ParentKey
    End With
    
    ' dbsCurrentDB schließen
ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
    End Sub

' durch Schaltfläche 'Suchen' ausführen,
' erzeugt Abfrage unter Berücksichtigung des Suchbegriffs
Public Sub Suchen(ByVal varSuchbegriff As Variant)
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim qdfAbfrage As DAO.QueryDef
    Set qdfAbfrage = dbsCurrentDB.CreateQueryDef
    
    Dim strAbfrageTitel As String
    strAbfrageTitel = "qryLiefergegenstand"
    
    If IsNull(varSuchbegriff) Then
        varSuchbegriff = "*"
    End If
    
    ' überprüft, ob die Abfrage qryLiefergegenstand bereits geöffnet ist,
    ' wenn ja Abfrage schließen
    If SysCmd(acSysCmdGetObjectState, acQuery, strAbfrageTitel) = 1 Then
        Debug.Print "clsLiefergegenstand.Suchen: " _
            & "qryLiefergegenstand geschlossen"
        DoCmd.Close acQuery, strAbfrageTitel, acSaveYes
    End If
    
    'prüfen ob Abfrage existiert, wenn ja Abfrage löschen
    Dim RecordSet As Object
    For Each RecordSet In dbsCurrentDB.QueryDefs
        If RecordSet.Name = strAbfrageTitel Then
            DoCmd.DeleteObject acQuery, strAbfrageTitel
        End If
    Next RecordSet
    
    ' SQL-Code erzeugen
    ' durchsucht die Felder aftrIDKey, bezeichnung, seriennummer, herstellerkennzeichen, BWIKey und ParentKey
    With qdfAbfrage
        .SQL = "SELECT tblLiefergegenstand.*, tblAuftrag.AftrTitel, tblAuftrag.BWIKey, tblAuftrag.Leistungsbeschreibung, tblAuftrag.ParentKey" _
                & " FROM tblLiefergegenstand LEFT JOIN tblAuftrag ON tblLiefergegenstand.aftrIdKey = tblAuftrag.AftrID" _
                & " WHERE tblLiefergegenstand.aftrIdKey LIKE '*" & varSuchbegriff & "*' OR tblLiefergegenstand.bezeichnung LIKE '*" & varSuchbegriff & "*' OR tblLiefergegenstand.seriennummer LIKE '*" & varSuchbegriff & "*' OR tblLiefergegenstand.herstellerkennzeichen LIKE '*" & varSuchbegriff & "' OR tblLiefergegenstand.zielAftrID LIKE '*" & varSuchbegriff & "*' OR tblAuftrag.BWIKey LIKE '*" & varSuchbegriff & "*' OR tblAuftrag.ParentKey LIKE '*" & varSuchbegriff & "*'" _
                & " ORDER BY tblAuftrag.BWIKey, tblLiefergegenstand.aftrIdKey, tblLiefergegenstand.LiefergegenstandID" _
                & " ;"
        .Name = strAbfrageTitel
    End With
    
    ' Query speichern
    With dbsCurrentDB.QueryDefs
        .Append qdfAbfrage
        .Refresh
    End With
    
ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
        qdfAbfrage.Close
        Set qdfAbfrage = Nothing
End Sub

' mit der 'Speichern' Schaltfläche verknüpfen,
' schreibt die Attribute des Objekts in den recordset
' wählt den recordset in Abhängigkeit von der Variablen
' mvarLiefergegenstandClipboard
Public Sub Speichern()
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim strAbfrageTitel As String
    ' strAbfrageTitel = "qryLiefergegenstand"
    strAbfrageTitel = "qryLiefergegenstand"
    
    'prüft ob Abfrage existiert, wenn nicht: goto ExitProc
        Dim RecordSet As Object
        Dim bolQueryExistiert As Boolean
        bolQueryExistiert = False
    
        For Each RecordSet In dbsCurrentDB.QueryDefs
            If RecordSet.Name = strAbfrageTitel Then
                bolQueryExistiert = True
            End If
        Next RecordSet
        
        If bolQueryExistiert = False Then
            Debug.Print "clsLiefergegenstand.Speichern: " _
                & strAbfrageTitel + " existiert nicht, " _
                 & "clsLiefergegenstand.Speichern abgebrochen"
            GoTo ExitProc
        End If
        
    ' prüft ob mvarLiefergegenstandClipboard null oder leer ist,
    ' wenn ja wirf Fehlermeldung (MsgBox) und gehe zu ExitProc
        Dim varClipboard As Variant
        varClipboard = mvarLiefergegenstandClipboard
        Debug.Print "clsLiefergegenstand.Speichern: " _
            & "varClipboard = " & varClipboard
        
        If IsNull(varClipboard) Or IsEmpty(varClipboard) Then
            Debug.Print "clsLiefergegenstand.Speichern: " _
                + varClipboard & " ist null, Prozedur Speichern abgebrochen"
            MsgBox "Kein Datensatz ausgewählt." _
                & vbCrLf & "Speichern nicht erfolgreich.", _
                vbCritical, "Speichern"
            GoTo ExitProc
        End If
    
    If mbolVerbatim = True Then
        Debug.Print "clsLiefergegenstand.Speichern: " _
            & "mvarLiefergegenstandClipboard = " _
            & mvarLiefergegenstandClipboard
    End If
    
    ' Abfrage qryLiefergegenstand nach mvarAuftragClipboard durchsuchen
        Dim strSucheAttribut As String
        strSucheAttribut = "LiefergegenstandID"

        Set mrstRecordset = dbsCurrentDB.OpenRecordset(strAbfrageTitel, dbOpenDynaset)
        mrstRecordset.FindFirst (strSucheAttribut & " = " & CLng(varClipboard))
        
    If mbolVerbatim = True Then
        ' MsgBox "mvarLiefergegenstandID: " & mvarLiefergegenstandID
        MsgBox "mvarLiefergegenstandID: " & mvarLiefergegenstandClipboard
    End If
    
    ' Debug.Print "clsLiefergegenstand.Speichern: mrstRecordset!LiefergegenstandID = " _
        & mrstRecordset!LiefergegenstandID
        
    ' Objektattribute in recordset schreiben
        With mrstRecordset
            .Edit
                !aftrIdKey = mvarAftrID
                !bezeichnung = mvarBezeichnung
                !linkLieferschein = mvarLieferscheinLink
                !seriennummer = mvarSeriennummer
                !anzahl = mvarAnzahl
                !herstellerkennzeichen = mvarHerstellerkennzeichen
                !Uanangebot = mvarUnterauftragnehmerAngebotLink
                !angebotNetto = mvarAngebotNetto
                !preisBrutto = mvarPreisBrutto
                !lieferdatum = mvarLieferdatum
                !lieferschein = mvarLieferschein
                !Bemerkung = mvarBemerkung
                !zielAftrID = mvarZielAftrID
                !zielLieferschein = mvarZielLieferschein
                !zielLinkLieferschein = mvarZielLinklieferschein
                !zielLieferdatum = mvarZielLieferdatum
                !LiefergegenstandLink = mvarLiefergegenstandLink
                !IstReserve = mvarIstReserve
                .Update
        End With
        
        MsgBox "Datensatz gespeichert", vbInformation, "Speichern"
        
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

' Überprüft, ob die Zielabfrage existiert
' ermittelt die höchste LiefergegenstandID und erhöht diese um 1
' erstellt einen neuen, leeren Datensatz in der Zielabfrage
' überträgt die LiefergegenstandID in mvarLiefergegenstandClipboard
' ruft gobjLiefergegenstand.Speichern auf, um die Attribute in den neuen
' Datensatz zu schreiben
Public Sub DatensatzHinzufuegen()
    
    ' Felder zurücksetzen
    ' clsLiefergegenstand.Default
    
    ' Database instanziieren
        Dim dbsCurrentDB As DAO.Database
        Set dbsCurrentDB = CurrentDb
        
    ' Abfragetitel festlegen
        Dim strAbfrageTitel As String
        strAbfrageTitel = "qryLiefergegenstand"
    
    ' prüft ob Abfrage existiert, wenn nicht: goto ExitProc
        Dim RecordSet As Object
        Dim bolQueryExistiert As Boolean
        bolQueryExistiert = False
    
        For Each RecordSet In dbsCurrentDB.QueryDefs
            If RecordSet.Name = strAbfrageTitel Then
                bolQueryExistiert = True
            End If
        Next RecordSet
        
        If bolQueryExistiert = False Then
            Debug.Print "clsLiefergegenstand.DatensatzHinzufuegen: " _
                & strAbfrageTitel + " existiert nicht, " _
                 & "clsLiefergegenstand.DatensatzHinzufuegen abgebrochen"
            GoTo ExitProc
        End If
    
    ' höchste LiefergegenstandID ermitteln und um 1 erhöhen
        Dim lngLiefergegenstandID As Long
        Dim strFeld As String
        Dim strTabelle As String
        strFeld = "LiefergegenstandID"
        strTabelle = "tblLiefergegenstand"
        
        lngLiefergegenstandID = DMax(strFeld, strTabelle) + 1

    If mbolVerbatim = True Then
        MsgBox "lngLiefergegenstandID = " & lngLiefergegenstandID, _
            vbOKOnly, "Gespeichert"
    End If
        
    ' mit Hilfe von strAbfrageTitel den Liefergegenstand-Recordset
    ' instanziieren
        Set mrstRecordset = dbsCurrentDB.OpenRecordset(strAbfrageTitel, _
            dbOpenDynaset)
        
    ' neuen Datensatz erzeugen
        mrstRecordset.AddNew
            mrstRecordset!LiefergegenstandID = lngLiefergegenstandID
            mrstRecordset!aftrIdKey = mvarAftrID
            mrstRecordset!bezeichnung = mvarBezeichnung
            mrstRecordset!anzahl = mvarAnzahl
            ' mrstRecordset!Bemerkung = mvarBemerkung
            mrstRecordset!Bemerkung = "TEST"
            mrstRecordset!IstReserve = mvarIstReserve
        mrstRecordset.Update
    
    mvarLiefergegenstandClipboard = lngLiefergegenstandID
    
    ' Abfrage qryLiefergegenstand aktualisieren
    ' gobjLiefergegenstand.Suchen (mvarLiefergegenstandClipboard)
    MsgBox "Liefergegenstand erzeugt", vbInformation, "Datensatz erzeugen"
    
    If mbolVerbatim = True Then
        Debug.Print "clsLiefergegenstand.DatensatzHinzufuegen: " _
            & "Datensatz erzeugt"
    End If
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Private Sub Class_Initialize()
    Debug.Print "Klasse 'clsLiefergegenstand' instanziiert"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse 'clsLiefergegenstand' terminiert"
End Sub

' Die Pruefe-Prozeduren sollen ermöglichen, dass null Werte in den
' Recordset geschrieben und von dort ausgelesen als solche ausgelesen
' werden können, gleichzeitig sollen sie verhindern, dass falsche
' Datentypen eingegeben werden können

' Prüft, ob der übergebene Wert vom Typ String ist
' und überführt ihn in diesen Typ
' Private Function PruefeString(ByVal varInput As Variant) As String
Private Function PruefeString(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
    
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeString: " _
                & "Property ist null oder leer"
        End If
        
        PruefeString = varInput
        Exit Function
    End If
    
    PruefeString = CStr(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ String ist
' und überführt ihn in diesen Typ
' nach dem Speichern muss das Formular akutalisiert werden,
' um den Link nutzen zu können
Private Function PruefeLink(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeLink: " _
                & "Property ist null oder leer"
        End If
            
        PruefeLink = varInput
        Exit Function
    End If
    
    ' Prüfen, ob varInput bereits im Link-Format (#...#) vorliegt,
    ' wenn ja, dann nicht mit # einschließen -> verhindert ungültige
    ' Pfade (##...##)
    If Left(varInput, 1) = "#" And Right(varInput, 1) = "#" Then
        PruefeLink = CStr(varInput)
        Exit Function
    End If
        
    PruefeLink = "#" + CStr(varInput) + "#"
    
End Function

' Prüft, ob der übergebene Wert vom Typ Date ist
' und überführt ihn in diesen Typ
'Private Function PruefeDatum(ByVal varInput As Variant) As Date
Private Function PruefeDatum(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand: " _
                & "Property ist null oder leer"
        End If
        
        PruefeDatum = varInput
        Exit Function
    End If
    
    PruefeDatum = CDate(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Currency ist
' und überführt ihn in diesen Typ
Private Function PruefeWaehrung(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeWaehrung: " _
                & "Property ist null oder leer"
        End If
        
        PruefeWaehrung = varInput
        Exit Function
    End If
    
    PruefeWaehrung = CCur(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Integer ist
' und überführt ihn in diesen Typ
Private Function PruefeInteger(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeInteger: " _
                & "Property ist null oder leer"
        End If
            
        PruefeInteger = varInput
        Exit Function
    End If
    
    PruefeInteger = CInt(varInput)
End Function

' Prüft, ob der übergegebene Wert vom Typ Boolean ist
' und überführt ihn in diesen Typ
Private Function PruefeBoolean(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
    
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeBoolean: " _
                & "Property ist null oder leer"
        End If
        
        PruefeBoolean = varInput
        Exit Function
    End If
    
    PruefeBoolean = CBool(varInput)
End Function

' Prüft, ob der übergegebene Wert vom Typ Long ist
' und überführt ihn in diesen Typ
Private Function PruefeLong(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand: " _
                & "Property ist null oder leer"
        End If
        
        PruefeLong = varInput
        Exit Function
    End If
    
    PruefeLong = CLng(varInput)
End Function

Public Property Get LiefergegenstandID() As Variant
    LiefergegenstandID = PruefeLong(mvarLiefergegenstandID)
End Property

Public Property Let LiefergegenstandID(ByVal varLiefergegenstandID As Variant)
    mvarLiefergegenstandID = PruefeLong(varLiefergegenstandID)
End Property

Public Property Get AftrID() As Variant
    AftrID = PruefeString(mvarAftrID)
End Property

Public Property Let AftrID(ByVal varAftrID As Variant)
    mvarAftrID = PruefeString(varAftrID)
End Property

Public Property Get bezeichnung() As Variant
    bezeichnung = PruefeString(mvarBezeichnung)
End Property

Public Property Let bezeichnung(ByVal varBezeichnung As Variant)
    mvarBezeichnung = PruefeString(varBezeichnung)
End Property

Public Property Get LieferscheinLink() As Variant
    LieferscheinLink = PruefeString(mvarLieferscheinLink)
End Property

Public Property Let LieferscheinLink(ByVal varLieferscheinLink As Variant)
    mvarLieferscheinLink = PruefeLink(varLieferscheinLink)
End Property

Public Property Get seriennummer() As Variant
    seriennummer = PruefeString(mvarSeriennummer)
End Property

Public Property Let seriennummer(ByVal varSeriennummer As Variant)
    mvarSeriennummer = PruefeString(varSeriennummer)
End Property

Public Property Get anzahl() As Variant
    anzahl = PruefeInteger(mvarAnzahl)
End Property

Public Property Let anzahl(ByVal varAnzahl As Variant)
    mvarAnzahl = PruefeInteger(varAnzahl)
End Property

Public Property Get herstellerkennzeichen() As Variant
    herstellerkennzeichen = PruefeString(mvarHerstellerkennzeichen)
End Property

Public Property Let herstellerkennzeichen(ByVal varHerstellerkennzeichen As Variant)
    mvarHerstellerkennzeichen = PruefeString(varHerstellerkennzeichen)
End Property

Public Property Get UnterauftragnehmerAngebotLink() As Variant
    UnterauftragnehmerAngebotLink = PruefeLink(mvarUnterauftragnehmerAngebotLink)
End Property

Public Property Let UnterauftragnehmerAngebotLink(ByVal varUnterauftragnehmerAngebotLink As Variant)
    mvarUnterauftragnehmerAngebotLink = PruefeLink(varUnterauftragnehmerAngebotLink)
End Property

Public Property Get angebotNetto() As Variant
    angebotNetto = PruefeWaehrung(mvarAngebotNetto)
End Property

Public Property Let angebotNetto(ByVal varAngebotNetto As Variant)
    mvarAngebotNetto = PruefeWaehrung(varAngebotNetto)
End Property

Public Property Get preisBrutto() As Variant
    preisBrutto = PruefeWaehrung(mvarPreisBrutto)
End Property

Public Property Let preisBrutto(ByVal varPreisBrutto As Variant)
    mvarPreisBrutto = PruefeWaehrung(varPreisBrutto)
End Property

Public Property Get lieferdatum() As Variant
    lieferdatum = PruefeDatum(mvarLieferdatum)
End Property

Public Property Let lieferdatum(ByVal varLieferdatum As Variant)
    mvarLieferdatum = PruefeDatum(varLieferdatum)
End Property

Public Property Get lieferschein() As Variant
    lieferschein = PruefeString(mvarLieferschein)
End Property

Public Property Let lieferschein(ByVal varLieferschein As Variant)
    mvarLieferschein = PruefeString(varLieferschein)
End Property

Public Property Get Bemerkung() As Variant
    Bemerkung = PruefeString(mvarBemerkung)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    mvarBemerkung = PruefeString(varBemerkung)
End Property

Public Property Get zielAftrID() As Variant
    zielAftrID = PruefeString(mvarZielAftrID)
End Property

Public Property Let zielAftrID(ByVal varZielAftrID As Variant)
    mvarZielAftrID = PruefeString(varZielAftrID)
End Property

Public Property Get zielLieferschein() As Variant
    zielLieferschein = PruefeString(mvarZielLieferschein)
End Property

Public Property Let zielLieferschein(ByVal varZielLieferschein As Variant)
    mvarZielLieferschein = PruefeString(varZielLieferschein)
End Property

Public Property Get zielLinkLieferschein() As Variant
    zielLinkLieferschein = PruefeLink(mvarZielLinklieferschein)
End Property

Public Property Let zielLinkLieferschein(ByVal varZielLinkLieferschein As Variant)
    mvarZielLinklieferschein = PruefeLink(varZielLinkLieferschein)
End Property

Public Property Get zielLieferdatum() As Variant
    zielLieferdatum = PruefeDatum(mvarZielLieferdatum)
End Property

Public Property Let zielLieferdatum(ByVal varZielLieferdatum As Variant)
    mvarZielLieferdatum = PruefeDatum(varZielLieferdatum)
End Property

Public Property Get LiefergegenstandLink() As Variant
    LiefergegenstandLink = PruefeLink(mvarLiefergegenstandLink)
End Property

Public Property Let LiefergegenstandLink(ByVal varLiefergegenstandLink As Variant)
    mvarLiefergegenstandLink = PruefeLink(varLiefergegenstandLink)
End Property

Public Property Get IstReserve() As Variant
    IstReserve = PruefeBoolean(mvarIstReserve)
End Property

Public Property Let IstReserve(ByVal varIstReserve As Variant)
    mvarIstReserve = PruefeBoolean(varIstReserve)
End Property

Public Property Get AftrTitel() As Variant
    AftrTitel = PruefeString(mvarAftrTitel)
End Property

Public Property Get BWIKey() As Variant
    BWIKey = PruefeString(mvarBWIKey)
End Property

Public Property Get LeistungsbeschreibungLink() As Variant
    LeistungsbeschreibungLink = PruefeLink(mvarLeistungsbeschreibungLink)
End Property

Public Property Get ParentKey() As Variant
    ParentKey = PruefeString(mvarParentKey)
End Property

Public Property Get Verbatim() As Boolean
    Verbatim = PruefeBoolean(mbolVerbatim)
End Property

Public Property Let Verbatim(ByVal bolVerbatim As Boolean)
    mvarVerbatim = PruefeBoolean(bolVerbatim)
End Property




