VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAuftrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsAuftrag

Option Compare Database
Option Explicit

' Hauptdatensatz der Klasse
Private mdbsCurrentDb As DAO.Database
Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher für die ID des aktuell
' mit dem Cursor ausgewählten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' steuert Debug-Meldungen
Private mbolVerbatim As Boolean

    Private mvarAftrID As Variant
    Private mvarAftrTitel As Variant
    Private mvarStatusKey As Variant
    Private mvarOwnerKey As Variant
    Private mvarPrioriteatKey As Variant
    Private mvarParentKey As Variant
    Private mvarBemerkung As Variant
    Private mvarBeginnSoll As Variant
    Private mvarEndeSoll As Variant
    Private mvarErstellt As Variant
    Private mvarKunde As Variant

Private Sub Default()
    mvarAftrID = Null
    mvarAftrTitel = Null
    mvarStatusKey = Null
    mvarOwnerKey = Null
    mvarPrioriteatKey = Null
    mvarParentKey = Null
    mvarBemerkung = Null
    mvarBeginnSoll = Null
    mvarEndeSoll = Null
    mvarErstellt = Null
    mvarKunde = Null
End Sub

' wird im Unterformular ausgeführt, lädt den aktuell ausgewählten
' Datensatz ins Objekt
Public Sub Laden(ByVal varID As Variant)
    Debug.Print "clsAuftragNew.Laden ausführen"
    
    Set mdbsCurrentDb = CurrentDb

    ' das Objekt in den Default-Zustand versetzen,
    ' wenn im Unterformular kein Datensatz mit dem
    ' Cursor ausgewählt wurde zu ExitProc springen
    If IsEmpty(varID) Or IsNull(varID) Then
        mvarClipboard = Null
        Default
        Debug.Print "clsAuftragNew.Laden: varID ist leer, " _
            & "Datensatz Defaultwert zugewiesen"
        GoTo ExitProc
    End If
    
    ' der Zwischenspeichervariable 'mvarClipboard' den ausgewählten
    ' Datensatz zuweisen - wird für die Prozedur Speichern benötigt
    mvarClipboard = varID
        
        ' zu durchsuchende Zielabfrage benennen
        Dim strAbfrageTitel As String
        strAbfrageTitel = "qryAuftragAuswahl"
        
        ' zu durchsuchendes Zielfeld benennen
        Dim strSucheZiel As String
        strSucheZiel = "AftrID"
    
        ' prüfen ob die qryAuftragAuswahl existiert,
        ' wenn nicht: goto ExitProc

        Dim RecordSet As Object
        Dim bolQueryExistiert As Boolean
        bolQueryExistiert = False
    
        For Each RecordSet In mdbsCurrentDb.QueryDefs
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
    
        ' Abfrage qryAuftragAuswahl nach varID durchsuchen
        Set mrstRecordset = mdbsCurrentDb.OpenRecordset(strAbfrageTitel, dbOpenDynaset)
        mrstRecordset.FindFirst (strSucheZiel & " = '" & CStr(varID) & "'")
        
    ' Attribute von mrstRecordset in Objekt laden
    With mrstRecordset
        mvarAftrID = !AftrID
        mvarAftrTitel = !AftrTitel
        mvarStatusKey = !StatusKey
        mvarOwnerKey = !OwnerKey
        mvarPrioriteatKey = !PrioritaetKey
        mvarParentKey = !ParentKey
        mvarBemerkung = !Bemerkung
        mvarBeginnSoll = !BeginnSoll
        mvarEndeSoll = !EndeSoll
        mvarErstellt = !Erstellt
        mvarKunde = !kunde
    End With
    
    ' mdbsCurrentDB schließen
ExitProc:
        mdbsCurrentDb.Close
        Set mdbsCurrentDb = Nothing
    End Sub

' mit der 'Speichern' Schaltfläche verknüpfen,
' schreibt die Attribute des Objekts in den recordset
Public Sub Speichern()
    Set mdbsCurrentDb = CurrentDb
    
    Dim strAbfrageTitel As String
    strAbfrageTitel = "qryAuftragAuswahl"
    
    'prüft ob Abfrage existiert, wenn nicht: goto ExitProc
        Dim RecordSet As Object
        Dim bolQueryExistiert As Boolean
        bolQueryExistiert = False
    
        For Each RecordSet In mdbsCurrentDb.QueryDefs
            If RecordSet.Name = strAbfrageTitel Then
                bolQueryExistiert = True
            End If
        Next RecordSet
        
        If bolQueryExistiert = False Then
            Debug.Print "clsAuftragNew.Speichern: " _
                & strAbfrageTitel & " existiert nicht, " _
                & "clsAuftragNew.Speichern abgebrochen"
            GoTo ExitProc
        End If

    ' prüfe ob mvarClipboard null oder leer ist,
    ' wenn ja wirf Fehlermeldung (MsgBox) und gehe zu ExitProc
        Dim varClipboard As Variant
        varClipboard = mvarClipboard
        
        If IsNull(varClipboard) Or IsEmpty(varClipboard) Then
            Debug.Print "clsAuftragNew.Speichern: " _
                + varClipboard & " is null, clsAuftragNew.Speichern abgebrochen"
            MsgBox "Kein Datensatz ausgewählt." & vbCrLf & "Speichern nicht erfolgreich.", vbCritical, "Speichern"
            GoTo ExitProc
        End If
        
    ' Abfrage qryAuftragAuswahl nach mvarClipboard durchsuchen
        Dim strSucheAttribut As String
        strSucheAttribut = "AftrID"

        Set mrstRecordset = mdbsCurrentDb.OpenRecordset(strAbfrageTitel, dbOpenDynaset)
        mrstRecordset.FindFirst (strSucheAttribut & " = '" & CStr(varClipboard) & "'")
        
    ' Objektattribute in recordset schreiben
        With mrstRecordset
            .Edit
                ' !AftrID = mvarAftrID
                !AftrTitel = mvarAftrTitel
                !EAkurzKey = mvarEAkurzKey
                !StatusKey = mvarStatusKey
                !OwnerKey = mvarOwnerKey
                !PrioritaetKey = mvarPrioriteatKey
                !ParentKey = mvarParentKey
                !Bemerkung = mvarBemerkung
                !BeginnSoll = mvarBeginnSoll
                !EndeSoll = mvarEndeSoll
                !Erstellt = mvarErstellt
                !kunde = mvarKunde
            .Update
        End With
        
        MsgBox "Datensatz gespeichert", vbInformation, "Speichern"
        
ExitProc:
    mdbsCurrentDb.Close
    Set mdbsCurrentDb = Nothing
End Sub

' durch Schaltfläche 'Suchen' ausführen,
' erzeugt Abfrage unter Berücksichtigung des Suchbegriffs
Public Sub Suchen(ByVal varSuchbegriff As Variant)
    Set mdbsCurrentDb = CurrentDb
    
    Dim qdfAbfrage As DAO.QueryDef
    Set qdfAbfrage = mdbsCurrentDb.CreateQueryDef
    
    Dim strAbfrageTitel As String
    strAbfrageTitel = "qryAuftragAuswahl"
    
    If IsNull(varSuchbegriff) Then
        varSuchbegriff = "*"
    End If
    
    ' überprüft, ob die Abfrage qryAuftragAuswahl bereits geöffnet ist,
    ' wenn ja Abfrage schließen
    If SysCmd(acSysCmdGetObjectState, acQuery, strAbfrageTitel) = 1 Then
        Debug.Print "clsAuftragNew.Suchen: " _
            & "qryAuftragAuswahl geschlossen"
        DoCmd.Close acQuery, strAbfrageTitel, acSaveYes
    End If
    
    'prüfen ob Abfrage existiert, wenn ja Abfrage löschen
    Dim RecordSet As Object
    For Each RecordSet In mdbsCurrentDb.QueryDefs
        If RecordSet.Name = strAbfrageTitel Then
            DoCmd.DeleteObject acQuery, strAbfrageTitel
        End If
    Next RecordSet
    
    ' SQL-Code erzeugen
    With qdfAbfrage
        .SQL = " SELECT qryAuftrag.*" _
                & " FROM qryAuftrag" _
                & " WHERE qryAuftrag.AftrID LIKE '*" & varSuchbegriff & "*' OR qryAuftrag.AftrTitel LIKE '*" & varSuchbegriff & "*' OR qryAuftrag.BWIKey LIKE '*" & varSuchbegriff & "*' OR qryAuftrag.ParentKey LIKE '*" & varSuchbegriff & "*'" _
                & " ;"
        .Name = strAbfrageTitel
    End With
    
    ' Query speichern
    With mdbsCurrentDb.QueryDefs
        .Append qdfAbfrage
        .Refresh
    End With
    
ExitProc:
        mdbsCurrentDb.Close
        Set mdbsCurrentDb = Nothing
        qdfAbfrage.Close
        Set qdfAbfrage = Nothing
End Sub

Public Function AddRecordset()
    Set mdbsCurrentDb = CurrentDb
    
    Dim strTabelleTitel As String
    strTabelleTitel = "tblAuftrag"
    
    ' pr�fe ob Tabelle existiert, wenn nicht
    ' gehe zu Exit Proc
    If basSupport.TabelleExistiert("tblAuftrag") = False Then
        Debug.Print "clsAuftrag.Add: tblAuftrag existiert nicht, Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    If mbolVerbatim = True Then
        Debug.Print "clsAuftrag.Add: tblAuftrag existiert"
    End If
    
    ' Titel des neuen Datensatzes abfragen, wenn leer dann ExitProc
    Dim varRecordsetName As Variant
    varRecordsetName = InputBox("Titel des Datensatzes eingeben", "Auftrag erzeugen")
    
    If basSupport.PflichtfeldIstLeer(varRecordsetName) = True Then
        Debug.Print "clsAuftrag.AddRecordset: RecordsetName ist leer, " _
            & "Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    varRecordsetName = CStr(varRecordsetName)
    
    If mbolVerbatim = True Then
        Debug.Print "clsAuftrag.AddRecordset!varRecordsetName: " + varRecordsetName
    End If
    
    ' Neuen Datensatz erzeugen
    Set mrstRecordset = mdbsCurrentDb.OpenRecordset(strTabelleTitel, dbOpenDynaset)
    
    mrstRecordset.AddNew
        mrstRecordset!AftrID = varRecordsetName
    mrstRecordset.Update
    
    ' Datensatzname an Zwischenablage �bergeben
    mvarClipboard = varRecordsetName
    
    If mbolVerbatim = True Then
        Debug.Print "clsAuftrag.AddRecordset: clsAuftrag!mvarClipboard = " + mvarClipboard
    End If
    
    ' Abschlussmeldung
    MsgBox "'" & mvarClipboard & "' erzeugt", vbOKOnly, "Auftrag hinzuf�gen"
        
ExitProc:
    mrstRecordset.Close
    Set mrstRecordset = Nothing
    mdbsCurrentDb.Close
    Set mdbsCurrentDb = Nothing
End Function


Private Sub Class_Initialize()
    Debug.Print "Klasse clsAuftrag instanziiert."
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsAuftrag terminiert."
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
        ' Debug.Print "Property ist null oder leer"
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
        ' Debug.Print "Property ist null oder leer"
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
        ' Debug.Print "Property ist null oder leer"
        PruefeDatum = varInput
        Exit Function
    End If
    
    PruefeDatum = CDate(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Currency ist
' und überführt ihn in diesen Typ
Private Function PruefeWaehrung(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeWaehrung = varInput
        Exit Function
    End If
    
    PruefeWaehrung = CCur(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Integer ist
' und überführt ihn in diesen Typ
Private Function PruefeInteger(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeWahrung = varInput
        Exit Function
    End If
    
    PruefeInteger = CInt(varInput)
End Function


' Get holt den Wert aus dem Zwischenspeicher und übergibt ihn in
' das Property des Objekts

' Let überträgt des Propertys in den Zwischenspeicher, durch die Prozedur 'Speichern'
' wird der Zwischenspeicher in den Recordset geschrieben

Public Property Let ID(ByVal varInput As Variant)
    mstrAftrID = PruefeString(varInput)
End Property

Public Property Get ID() As Variant
    ID = PruefeString(mvarAftrID)
End Property

Public Property Let Titel(ByVal varTitel As Variant)
    mvarAftrTitel = PruefeString(varTitel)
End Property

Public Property Get Titel() As Variant
    Titel = PruefeString(mvarAftrTitel)
End Property

' Sollte nur durch den Import eines ICD-Reports geändert werden
' Public Property Let ICDStatus(ByVal strInput as String)
'     mstrStatusKey = strInput
' End Property

Public Property Get ICDStatus() As Variant
    ICDStatus = PruefeString(mvarStatusKey)
End Property

' Public Property Let Prio(ByVal strInput as String)
'     mintPrioriteatKey = strInput
' End Property

' Public Property Get Prio() As String
'     Prio = mintPrioriteatKey
' End Property

Public Property Let Parent(ByVal varParent As Variant)
    mvarParentKey = PruefeString(varParent)
End Property

Public Property Get Parent() As Variant
    Parent = PruefeString(mvarParentKey)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    mvarBemerkung = PruefeString(varBemerkung)
End Property

Public Property Get Bemerkung() As Variant
    Bemerkung = PruefeString(mvarBemerkung)
End Property

' Public Property Let BeginnSoll(ByVal varBeginnSoll As Variant)
'      mvarBeginnSoll = PruefeDatum(varBeginnSoll)
'  End Property

Public Property Get BeginnSoll() As Variant
    BeginnSoll = PruefeDatum(mvarBeginnSoll)
End Property

' Public Property Let EndeSoll(ByVal varEndeSoll As Variant)
'     mvarEndeSoll = PruefeDatum(varEndeSoll)
' End Property

Public Property Get EndeSoll() As Variant
    EndeSoll = PruefeDatum(mvarEndeSoll)
End Property

' Public Property Let Erstellt(ByVal dtmInput as Date)
'     mdtmErstellt = dtmInput
' End Property

' Public Property Get Erstellt() As Date
'     Erstellt = mdtmErstellt
' End Property

' Public Property Let kunde(ByVal strInput as String)
'     mstrKunde = strInput
' End Property

' Public Property Get kunde() As String
'     kunde = mstrKunde
' End Property

Public Property Let Verbatim(ByVal bolInput As Boolean)
    mbolVerbatim = basSupport.PruefeBoolean(bolInput)
End Property

Public Property Get Verbatim() As Boolean
    Verbatim = basSupport.PruefeBoolean(mbolVerbatim)
End Property
