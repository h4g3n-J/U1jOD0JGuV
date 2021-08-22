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

' clipboard, relevant for saving and loading methods
Private mvarClipboard As Variant

' Private Const cintColumns As Integer = 2
' Private Const cintRows As Integer = 11
' Dim mavarAuftragAttributes(cintRows, cintColumns) As Variant
Dim mavarAuftragAttributes As Variant
    
Private Sub Class_Initialize()

    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Class_Initialize: Klasse clsAuftrag instanziieren."
    End If
    
    ' load default values
    mavarAuftragAttributes = DefaultValue
    
End Sub

Private Sub Class_Terminate()
    If gconVerbatim Then
        Debug.Print "clsAuftrag.Class_Terminate: Klasse clsAuftrag terminiert."
    End If
End Sub

' default class values, executed with clsAuftrag.Class_Initialize
Private Function DefaultValue()
    
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.DefaultValue ausfuehren"
    End If
    
    ' see modul header to change dimensions
    ' Dim varAuftragSettings(cintRows, cintColumns) As Variant
    Dim varAuftragSettings(11, 2) As Variant
    
    varAuftragSettings(0, 0) = "Attribut name"
        varAuftragSettings(0, 1) = "data type"
        varAuftragSettings(0, 2) = "value"
    varAuftragSettings(1, 0) = "AftrID"
        varAuftragSettings(1, 1) = "string"
        varAuftragSettings(1, 2) = Null
    varAuftragSettings(2, 0) = "AftrTitel"
        varAuftragSettings(2, 1) = "string"
        varAuftragSettings(2, 2) = Null
    varAuftragSettings(3, 0) = "StatusKey"
        varAuftragSettings(3, 1) = "string"
        varAuftragSettings(3, 2) = Null
    varAuftragSettings(4, 0) = "OwnerKey"
        varAuftragSettings(4, 1) = "string"
        varAuftragSettings(4, 2) = Null
    varAuftragSettings(5, 0) = "PrioritaetKey"
        varAuftragSettings(5, 1) = "string"
        varAuftragSettings(5, 2) = Null
    varAuftragSettings(6, 0) = "ParentKey"
        varAuftragSettings(6, 1) = "string"
        varAuftragSettings(6, 2) = Null
    varAuftragSettings(7, 0) = "Bemerkung"
        varAuftragSettings(7, 1) = "string"
        varAuftragSettings(7, 2) = Null
    varAuftragSettings(8, 0) = "BeginnSoll"
        varAuftragSettings(8, 1) = "date"
        varAuftragSettings(8, 2) = Null
    varAuftragSettings(9, 0) = "EndeSoll"
        varAuftragSettings(9, 1) = "date"
        varAuftragSettings(9, 2) = Null
    varAuftragSettings(10, 0) = "Erstellt"
        varAuftragSettings(10, 1) = "date"
        varAuftragSettings(10, 2) = Null
    varAuftragSettings(11, 0) = "Kunde"
        varAuftragSettings(11, 1) = "string"
        varAuftragSettings(11, 2) = Null
        
    DefaultValue = varAuftragSettings
End Function

' set to default when no recordset is selected
' check if source (query) exists
' check if recordset exists
' assign recordset to mavarAuftragAttributes
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)
    
    ' error handler in case no recordset is selected
    If IsNull(varRecordsetName) Then
        ' verbatim message
        If gconVerbatim = True Then
            Debug.Print "clsAuftrag.SelectRecordset: varRecordsetName ist leer, " _
                & "Defaultwert anzeigen"
        End If
        ' load default Value
        mavarAuftragAttributes = DefaultValue
        GoTo ExitProc
    End If
    
    ' recordset source settings
    ' source type, feasible values: table, query
    Dim strSourceType As String
    strSourceType = "table"
    ' source name
    Dim strSourceName As String
    strSourceName = "tblAuftrag"
    ' name of scanned column
    Dim strKeyColumn As String
    strKeyColumn = "AftrID"
        
    ' check if source exists
    ' negative: return error + exitProc
    If ObjectExists(strSourceName, strSourceType) = False Then
        Debug.Print "clsAuftrag.SelectRecordset: " & strSourceName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    End If
    
    ' check if recordset exists
    ' likely obsolete, since the recordset is chosen from selection of recordsets
    If RecordsetExists(strSourceName, strKeyColumn, varRecordsetName) <> True Then
        Debug.Print "clsAuftrag.SelectRecordset: " & varRecordsetName & " existiert nicht. Prozedur abgebrochen."
        MsgBox varRecordsetName & " existiert nicht. Prozedur abgebrochen.", vbCritical, "Datensatz nicht gefunden"
        GoTo ExitProc
    End If
    
    ' define database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    ' define recordset
    Dim rstRecordset As DAO.RecordSet

    ' select recordset
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strSourceName, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = '" & CStr(varRecordsetName) & "'")

    ' load recordset to class properties
    Dim inti As Integer
    For inti = LBound(mavarAuftragAttributes, 1) + 1 To UBound(mavarAuftragAttributes, 1)
        ' property value [mavarAuftragAttributes(inti, 2)] = recordset value [mavarAuftragAttributes(inti, 0)]
        mavarAuftragAttributes(inti, 2) = rstRecordset.Fields(mavarAuftragAttributes(inti, 0))
    Next
    
    ' set local varbatim message on or off
    Dim bolVerbatim As Boolean
    bolVerbatim = False
    
    ' verbatim message: return data from class properties
    If bolVerbatim = True Then
        For inti = LBound(mavarAuftragAttributes, 1) + 1 To UBound(mavarAuftragAttributes, 1)
            Debug.Print "clsAuftrag.SelectRecordset: " & mavarAuftragAttributes(inti, 0) & vbTab & mavarAuftragAttributes(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

' write class properties to recordset
Public Sub SaveRecordset(ByVal varRecordsetName As Variant)
    
    Dim strTargetTableName As String
    strTargetTableName = "tblAuftrag"
    
    Dim strKeyColumn As String
    strKeyColumn = mavarAuftragAttributes(1, 0)
    
    ' verbatim message
    If gconVerbatim Then
            Debug.Print "clsAuftrag.SaveRecordset ausfuehren"
    End If
    
    ' error handler, case no item is selected a.k.a. varRecordset is empty
    If IsNull(varRecordsetName) Then
        ' verbatim message
        Debug.Print "clsAuftrag.SaveRecordset: kein Recordset ausgew�hlt, Speichern abbrechen"
        ' error message
        MsgBox "Kein Datensatz ausgewaehlt", vbCritical, "Speichern abgebrochen"
        ' abort procedure
        GoTo ExitProc
    End If
    
    ' declare current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' declare Auftrag
    Dim rstAftr As DAO.RecordSet
    Set rstAftr = dbsCurrentDB.OpenRecordset(strTargetTableName, dbOpenDynaset)
    
    ' find fist recordset where AftrID = varRecordsetName
    rstAftr.FindFirst (strKeyColumn & " = '" & varRecordsetName & "'")
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' move through textboxes, transform datatype and write properties to recordset
    Dim inti As Integer
    With rstAftr
        .Edit
            For inti = LBound(mavarAuftragAttributes, 1) + 1 To UBound(mavarAuftragAttributes, 1)
                ' local verbatim
                If bolLocalVerbatim Then
                    Debug.Print "clsAuftrag.SaveRecordset: " & mavarAuftragAttributes(inti, 0) & " = " & mavarAuftragAttributes(inti, 2)
                End If
                ' attribute name: mavarAuftragAttributes(inti, 0), _
                ' property value: mavarAuftragAttributes(inti, 2), _
                ' data type: mavarAuftragAttributes(inti, 1)
                CallByName rstAftr, mavarAuftragAttributes(inti, 0), VbLet, basSupport.CheckDataType(mavarAuftragAttributes(inti, 2), mavarAuftragAttributes(inti, 1))
            Next
        .Update
    End With
    
    ' success message
    MsgBox "Datensatz '" & varRecordsetName & "' gespeichert", vbInformation, "Gespeichert"
    
    ' clean up
    rstAftr.Close
    Set rstAftr = Nothing
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

' inaktiv
' durch Schaltflaeche 'Suchen' ausfuehren,
' erzeugt Abfrage unter Beruecksichtigung des Suchbegriffs
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
Public Function AddRecordset(ByVal strRecordsetName As String) As Boolean
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAuftrag.AddRecordset ausfuehren"
    End If
    
    ' settings
    Dim strTargetName As String
    strTargetName = "tblAuftrag"
    
    Dim strKeyColumn As String
    strKeyColumn = "AftrID"
    
    ' error state
    Dim bolError As Boolean
    bolError = False
    
    Dim strDialogMessage As String
    strDialogMessage = "Auftrag ID eingeben"
    
    Dim strDialogTitle As String
    strDialogTitle = "Auftrag erfassen"
    
    ' check if strTargetName exists
    If basSupport.ObjectExists(strTargetName, "table") = False Then
        Debug.Print "basSupport.AddrecordsetParent: " & strTargetName & " existiert nicht. Prozedur abgebrochen."
        bolError = True
        GoTo ExitProc
    Else:
        If gconVerbatim = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strTargetName & " existiert."
        End If
    End If
    
    Dim strArtifact As String
    strArtifact = "Auftrag"
    
    ' check if recordset name is empty, if true then messagebox + exit procedure
    If basSupport.PflichtfeldIstLeer(strRecordsetName) = True Then
        Debug.Print "clsAuftrag.Addrecordset: " & strArtifact & _
            " ist leer. Prozedur abgebrochen."
        MsgBox strArtifact & " ist leer. Prozedur abgebrochen.", vbCritical, "Fehler"
        bolError = True
        GoTo ExitProc
    End If
    
    ' check if recordset already exists
    If RecordsetExists(strTargetName, strKeyColumn, strRecordsetName) = True Then
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
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTargetName, dbOpenDynaset)

    rstRecordset.AddNew
        rstRecordset.Fields(strKeyColumn) = strRecordsetName
    rstRecordset.Update
        
    ' completion message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AddRecordset: " & strRecordsetName & " erzeugt."
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
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "AftrID")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varAftrID, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get AftrID() As Variant
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrID (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "AftrID")

    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    AftrID = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let AftrTitel(ByVal varAftrTitel As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "AftrTitel")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varAftrTitel, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get AftrTitel() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "AftrTitel")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    AftrTitel = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

' Sollte nur durch den Import eines ICD-Reports geändert werden
Public Property Let StatusKey(ByVal varStatusKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "StatusKey")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varStatusKey, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get StatusKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "StatusKey")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    StatusKey = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let PrioritaetKey(ByVal varPrioritaetKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "PrioritaetKey")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varPrioritaetKey, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get PrioritaetKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "PrioritaetKey")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    PrioritaetKey = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let ParentKey(ByVal varParentKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "ParentKey")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varParentKey, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get ParentKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "ParentKey")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    ParentKey = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Bemerkung")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varBemerkung, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get Bemerkung() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Bemerkung")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    Bemerkung = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let BeginnSoll(ByVal varBeginnSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "BeginnSoll")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varBeginnSoll, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get BeginnSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "BeginnSoll")
    
    ' get value [mavarAuftragAttributes(intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    BeginnSoll = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let EndeSoll(ByVal varEndeSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "EndeSoll")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varEndeSoll, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get EndeSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "EndeSoll")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    EndeSoll = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let Erstellt(ByVal varErstellt As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Erstellt")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varErstellt, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get Erstellt() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Erstellt")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    Erstellt = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let kunde(ByVal varKunde As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Kunde")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varKunde, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get kunde() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "Kunde")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform corresponding to data type [mavarAuftragAttributes(intItemPosition, 1)]
    kunde = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Let OwnerKey(ByVal varOwnerKey As Variant)
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Owner (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "OwnerKey")
    
    ' transform input [mavarAuftragAttributes(intItemPosition, 1)], assign to attribut value [mavarAuftragAttributes( intItemPosition, 2)]
    mavarAuftragAttributes(intItemPosition, 2) = basSupport.CheckDataType(varOwnerKey, mavarAuftragAttributes(intItemPosition, 1))
    
End Property

Public Property Get OwnerKey() As Variant
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarAuftragAttributes, "OwnerKey")
    
    ' get value [mavarAuftragAttributes( intItemPosition, 2)], transform data type [mavarAuftragAttributes(intItemPosition, 1)]
    OwnerKey = basSupport.CheckDataType(mavarAuftragAttributes(intItemPosition, 2), mavarAuftragAttributes(intItemPosition, 1))
    
End Property

' checks if a specific table or query exists
' returns true if positive
' strModus feasible values: "table", "query"
Private Function ObjectExists(ByVal strObjectName, strModus As String) As Boolean
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim RecordSet As Object
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "execute clsAuftrag.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAuftrag.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
    End If
    
    Select Case strModus
        Case "table"
            For Each RecordSet In dbsCurrentDB.TableDefs
                If RecordSet.Name = strObjectName Then
                    bolObjectExists = True
                End If
            Next RecordSet
        Case "query"
            For Each RecordSet In dbsCurrentDB.QueryDefs
                If RecordSet.Name = strObjectName Then
                    bolObjectExists = True
                End If
            Next RecordSet
    End Select
    
    ObjectExists = bolObjectExists
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Function

Private Function RecordsetExists(ByVal varTblName As Variant, ByVal varFieldName As Variant, ByVal varRecordsetName As Variant) As Boolean

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAuftrag.RecordsetExists"
    End If

    Dim bolStatus As Boolean
    bolStatus = False
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAuftrag.RecordsetExists: varTblName = " & varTblName & " , varFieldName = " & " , varRecordsetName = " & varRecordsetName
    End If
    
    varTblName = CStr(varTblName)
    varFieldName = CStr(varFieldName)
    varRecordsetName = CStr(varRecordsetName)
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(varTblName, dbOpenDynaset)
    
    If DCount(varFieldName, varTblName, varFieldName & " Like '" & varRecordsetName & "'") > 0 Then
        bolStatus = True
    End If
    
    RecordsetExists = bolStatus
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.RecordsetExists executed"
    End If
    
End Function


