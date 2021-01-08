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
' Private mdbsCurrentDb As DAO.Database
' Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher für die ID des aktuell
' mit dem Cursor ausgewählten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' steuert Debug-Meldungen
Private mbolverbatim As Boolean
    
Dim mastrColumnToVariable(2, 10) As Variant
    
Private Sub Class_Initialize()
    Debug.Print "clsAuftrag.Class_Initialize: Klasse clsAuftrag instanziiert."
    
    mastrColumnToVariable(0, 0) = "AftrID"
        mastrColumnToVariable(1, 0) = "string"
        mastrColumnToVariable(2, 0) = Null
    mastrColumnToVariable(0, 1) = "AftrTitel"
        mastrColumnToVariable(1, 1) = "string"
        mastrColumnToVariable(2, 1) = Null
    mastrColumnToVariable(0, 2) = "StatusKey"
        mastrColumnToVariable(1, 2) = "string"
        mastrColumnToVariable(2, 2) = Null
    mastrColumnToVariable(0, 3) = "OwnerKey"
        mastrColumnToVariable(1, 3) = "string"
        mastrColumnToVariable(2, 3) = Null
    mastrColumnToVariable(0, 4) = "PrioritaetKey"
        mastrColumnToVariable(1, 4) = "string"
        mastrColumnToVariable(2, 4) = Null
    mastrColumnToVariable(0, 5) = "ParentKey"
        mastrColumnToVariable(1, 5) = "string"
        mastrColumnToVariable(2, 5) = Null
    mastrColumnToVariable(0, 6) = "Bemerkung"
        mastrColumnToVariable(1, 6) = "string"
        mastrColumnToVariable(2, 6) = Null
    mastrColumnToVariable(0, 7) = "BeginnSoll"
        mastrColumnToVariable(1, 7) = "date"
        mastrColumnToVariable(2, 7) = Null
    mastrColumnToVariable(0, 8) = "EndeSoll"
        mastrColumnToVariable(1, 8) = "date"
        mastrColumnToVariable(2, 8) = Null
    mastrColumnToVariable(0, 9) = "Erstellt"
        mastrColumnToVariable(1, 9) = "date"
        mastrColumnToVariable(2, 9) = Null
    mastrColumnToVariable(0, 10) = "Kunde"
        mastrColumnToVariable(1, 10) = "string"
        mastrColumnToVariable(2, 10) = Null
        
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsAuftrag.Class_Terminate: Klasse clsAuftrag terminiert."
End Sub

Private Sub default()
    ' l�schen
End Sub

' default muss neu definiert werden
Public Sub SelectRecordset(ByVal varRecordsetName As Variant, Optional ByVal bolVerbatim As Boolean = False)

    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strObjectTitle As String
    strObjectTitle = "tblAuftrag"

    ' the type of the containing table, feasible values:
    ' table, query
    Dim strObjectType As String
    strObjectType = "table"

    ' name of the column that contains the recordset
    Dim strKeyColumn As String
    strKeyColumn = "AftrID"

    ' define database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' define recordset
    Dim rstRecordset As DAO.RecordSet
    
    ' loop variable
    Dim intLoopVar As Integer
    
    ' debug message: returns input of the procedure
    If bolVerbatim = True Then
        Debug.Print "basSupport.SelectRecordset: varRecordsetName = " & varRecordsetName
    End If

    ' if no recordset is selected set output to default value
    ' end exitProc
    If IsNull(varRecordsetName) Then
        ' default
        Debug.Print "clsAuftrag.SelectRecordset: varRecordsetName ist leer, " _
            & "Defaultwert anzeigen"
        GoTo ExitProc
    End If
        
    ' check if referenced query exists, when negative
    ' return error + exitProc
    If basSupport.ObjectExists(strObjectTitle, strObjectType, True) = False Then
        Debug.Print "clsAuftrag.SelectRecordset: " & strObjectTitle & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    End If
    
    ' check if recordset exists
    ' likely obsolete, since the recordset is chosen from a recordset selection
    If basSupport.RecordsetExists(strObjectTitle, strKeyColumn, varRecordsetName) <> True Then
        Debug.Print "clsAuftrag.SelectRecordset: " & varRecordsetName & " existiert nicht. Prozedur abgebrochen."
        MsgBox varRecordsetName & " existiert nicht. Prozedur abgebrochen.", vbCritical, "Datensatz nicht gefunden"
        GoTo ExitProc
    End If

    ' select recordset from table or query
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strObjectTitle, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = '" & CStr(varRecordsetName) & "'")

    ' assign recordset to array
    For intLoopVar = LBound(mastrColumnToVariable, 2) To UBound(mastrColumnToVariable, 2)
        mastrColumnToVariable(2, intLoopVar) = rstRecordset.Fields(mastrColumnToVariable(0, intLoopVar))
    Next
    
    ' debug message: returns data content of the array
    If bolVerbatim = True Then
        For intLoopVar = LBound(mastrColumnToVariable, 2) To UBound(mastrColumnToVariable, 2)
            Debug.Print "clsAuftrag.SelectRecordset: " & mastrColumnToVariable(0, intLoopVar) & vbTab & mastrColumnToVariable(2, intLoopVar)
        Next
    End If
    
    ' assigne AftrID to Clipboard
    ' likely obsolete
    mvarClipboard = mastrColumnToVariable(2, 0)

    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
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
                !Kunde = mvarKunde
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
    mvarClipboard = basSupport.AddRecordsetParent("tblAuftrag", "AftrID", "Auftrag", "AuftragID eingeben", "Auftrag erfassen", True)
End Function

Public Property Let AftrID(ByVal varAftrID As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrID")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varAftrID, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get AftrID() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrID")

    AftrID = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let AftrTitel(ByVal varAftrTitel As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrTitel")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varAftrTitel, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get AftrTitel() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "varTitel")
    
    AftrTitel = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

' Sollte nur durch den Import eines ICD-Reports geändert werden
Public Property Let StatusKey(ByVal varStatusKey As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "StatusKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varStatusKey, mastrColumnToVariable(1, intItemPosition, mbolverbatim))
End Property

Public Property Get StatusKey() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "StatusKey")
    
    ICDStatus = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let PrioritaetKey(ByVal varPrioritaetKey As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "PrioritaetKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varPrioritaetKey, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get PrioritaetKey() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "PrioritaetKey")
    
    PrioritaetKey = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let ParentKey(ByVal varParentKey As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "ParentKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varParentKey, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get ParentKey() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "ParentKey")
    
    ParentKey = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Bemerkung")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varBemerkung, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get Bemerkung() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Bemerkung")
    
    Bemerkung = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let BeginnSoll(ByVal varBeginnSoll As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "BeginnSoll")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varBemerkung, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get BeginnSoll() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "BeginnSoll")
    
    BeginnSoll = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let EndeSoll(ByVal varEndeSoll As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "EndeSoll")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varEndeSoll, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get EndeSoll() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "EndeSoll")
    
    EndeSoll = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let Erstellt(ByVal varErstellt As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Erstellt")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varErstellt, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get Erstellt() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Erstellt")
    
    Erstellt = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let Kunde(ByVal varKunde As Variant)
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Kunde")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varKunde, mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Get Kunde() As Variant
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Kunde")
    
    Kunde = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition), mbolverbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolverbatim = CBool(varVerbatim)
End Property

Public Property Get Verbatim() As Variant
    Verbatim = CBool(mbolverbatim)
End Property
