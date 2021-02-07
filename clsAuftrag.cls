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

Private mvarClipboard As Variant
  
Dim mastrColumnToVariable(2, 10) As Variant
    
Private Sub Class_Initialize()

    ' initialize the variables
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
            
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Class_Initialize: Klasse clsAuftrag instanziiert."
    End If
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsAuftrag.Class_Terminate: Klasse clsAuftrag terminiert."
End Sub

Private Sub default()
    ' l�schen
End Sub

' default muss neu definiert werden
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strObjectTitle As String
    strObjectTitle = "tblAuftrag"

    ' the type of the containing table,
    ' feasible values are: table, query
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
    If gconVerbatim = True Then
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
    If basSupport.ObjectExists(strObjectTitle, strObjectType) = False Then
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
    If gconVerbatim = True Then
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

' writes class properties to recordset
Public Sub SaveRecordset(ByVal strRecordsetName As String)
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstAftr As DAO.RecordSet
    Set rstAftr = dbsCurrentDB.OpenRecordset("tblAuftrag", dbOpenDynaset)
    rstAftr.FindFirst (mastrColumnToVariable(0, 0) & " = '" & strRecordsetName & "'")
    
    Dim inti As Integer
    With rstAftr
        .Edit
            For inti = LBound(mastrColumnToVariable, 2) + 1 To UBound(mastrColumnToVariable, 2)
                CallByName rstAftr, mastrColumnToVariable(0, inti), VbLet, mastrColumnToVariable(2, inti)
            Next
        .Update
    End With
    
    ' success message
    MsgBox "Datensatz '" & strRecordsetName & "' gespeichert", vbInformation, "Gespeichert"
    
ExitProc:
    rstAftr.Close
    dbsCurrentDB.Close
    
    Set rstAftr = Nothing
    Set dbsCurrentDB = Nothing
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

' create recordset, return True if
' 1. referenced table dosn't exist,
' 2. recordset name is empty,
' 3. recordset name is taken
Public Function AddRecordset(ByVal RecordsetName As String) As Boolean
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    
    ' error state
    Dim bolError As Boolean
    bolError = False
    
    Dim strTableName As String
    Dim strKeyColumn As String
    Dim strRecordsetName As String
    Dim strArtifact As String
    Dim strDialogMessage As String
    Dim strDialogTitle As String
    
    strTableName = "tblAuftrag"
    strKeyColumn = "AftrID"
    strRecordsetName = RecordsetName
    strArtifact = "Auftrag"
    strDialogMessage = "Auftrag ID eingeben"
    strDialogTitle = "Auftrag erfassen"
    
    ' check if strTableName exists
    If basSupport.ObjectExists(strTableName, "table") = False Then
        Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        bolError = True
        GoTo ExitProc
    Else:
        If gconVerbatim = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert."
        End If
    End If
    
    ' check if recordset name is empty, if true then messagebox + exit procedure
    If basSupport.PflichtfeldIstLeer(strRecordsetName) = True Then
        Debug.Print "clsAuftrag.Addrecordset: " & strArtifact & _
            " ist leer. Prozedur abgebrochen."
        MsgBox strArtifact & " ist leer. Prozedur wird abgebrochen.", vbCritical, "Fehler"
        bolError = True
        GoTo ExitProc
    End If
    
    ' check if recordset already exists
    If basSupport.RecordsetExists(strTableName, strKeyColumn, strRecordsetName) = True Then
        Debug.Print "clsAuftrag.AddRecordset: " & strRecordsetName _
            & " existiert bereits. Prozedur abgebrochen."
        ' error message via messagebox
        MsgBox "Ein Datensatz mit dem Namen '" & strRecordsetName & _
            "' existiert bereits." & vbCrLf & "Prozedur abgebrochen.", _
            vbCritical, "Doppelter Eintrag"
        bolError = True
        GoTo ExitProc
    End If

    ' create recordset
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)

        rstRecordset.AddNew
            rstRecordset.Fields(strKeyColumn) = strRecordsetName
        rstRecordset.Update
        
    ' completion message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AddRecordset: " & strRecordsetName & " created."
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing
    
ExitProc:
    AddRecordset = bolError
End Function

Public Property Let AftrID(ByVal varAftrID As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrID (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrID")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varAftrID, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get AftrID() As Variant
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrID (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrID")

    AftrID = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let AftrTitel(ByVal varAftrTitel As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrTitel")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varAftrTitel, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get AftrTitel() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "AftrTitel")
    
    AftrTitel = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

' Sollte nur durch den Import eines ICD-Reports geändert werden
Public Property Let StatusKey(ByVal varStatusKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "StatusKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varStatusKey, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get StatusKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "StatusKey")
    
    StatusKey = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let PrioritaetKey(ByVal varPrioritaetKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "PrioritaetKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varPrioritaetKey, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get PrioritaetKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "PrioritaetKey")
    
    PrioritaetKey = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let ParentKey(ByVal varParentKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "ParentKey")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varParentKey, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get ParentKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "ParentKey")
    
    ParentKey = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Bemerkung")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varBemerkung, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get Bemerkung() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Bemerkung")
    
    Bemerkung = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let BeginnSoll(ByVal varBeginnSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "BeginnSoll")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varBemerkung, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get BeginnSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "BeginnSoll")
    
    BeginnSoll = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let EndeSoll(ByVal varEndeSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "EndeSoll")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varEndeSoll, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get EndeSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "EndeSoll")
    
    EndeSoll = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let Erstellt(ByVal varErstellt As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Erstellt")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varErstellt, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get Erstellt() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Erstellt")
    
    Erstellt = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Let kunde(ByVal varKunde As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Let) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Kunde")
    
    mastrColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varKunde, mastrColumnToVariable(1, intItemPosition))
End Property

Public Property Get kunde() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Get) ausfuehren"
    End If
    
    Dim intItemPosition
    intItemPosition = basSupport.FindItemArray(mastrColumnToVariable, "Kunde")
    
    kunde = basSupport.CheckDataType(mastrColumnToVariable(2, intItemPosition), mastrColumnToVariable(1, intItemPosition))
End Property
