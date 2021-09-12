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

' container for selected recordset, is also relevant for properties procedures
' look-up DefaultProperties for array composition
Dim mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean
    
Private Sub Class_Initialize()

    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Class_Initialize: Klasse clsAuftrag instanziieren."
    End If
    
    ' load default values
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    If gconVerbatim Then
        Debug.Print "clsAuftrag.Class_Terminate: Klasse clsAuftrag terminiert."
    End If
End Sub

' default class values, executed with clsAuftrag.Class_Initialize
Private Function DefaultProperties()
    
    mbolVerbatim = False
    
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAuftrag.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 11
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "value"
    varDefaultProperties(1, 0) = "AftrID"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "AftrTitel"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "StatusKey"
        varDefaultProperties(3, 1) = "string"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "OwnerKey"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "PrioritaetKey"
        varDefaultProperties(5, 1) = "string"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "ParentKey"
        varDefaultProperties(6, 1) = "string"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "Bemerkung"
        varDefaultProperties(7, 1) = "string"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "BeginnSoll"
        varDefaultProperties(8, 1) = "date"
        varDefaultProperties(8, 2) = Null
    varDefaultProperties(9, 0) = "EndeSoll"
        varDefaultProperties(9, 1) = "date"
        varDefaultProperties(9, 2) = Null
    varDefaultProperties(10, 0) = "Erstellt"
        varDefaultProperties(10, 1) = "date"
        varDefaultProperties(10, 2) = Null
    varDefaultProperties(11, 0) = "Kunde"
        varDefaultProperties(11, 1) = "string"
        varDefaultProperties(11, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)
        
    DefaultProperties = varDefaultProperties
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.DefaultProperties executed"
    End If
    
End Function

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' command message
    If mbolVerbatim Then
        Debug.Print "execute clsAuftrag.SelectRecordset"
    End If

    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblAuftrag"
    
    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"
    
    ' name of the column that contains the sought recordset
    Dim strKeyColumn As String
    strKeyColumn = "AftrID"
    
    ' define database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' define recordset
    Dim rstRecordset As DAO.RecordSet
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' error handler if no recordset is selected
    If IsNull(varRecordsetName) Then
        ' verbatim message
        If bolLocalVerbatim = True Then
            Debug.Print "clsAuftrag.SelectRecordset: varRecordsetName is empty, " _
                & "prompt default value"
        End If
        ' load default Value
        mvarRecordsetClipboard = DefaultProperties
        GoTo ExitProc
    End If
    
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsAuftrag.SelectRecordset: " & strTableName & " does not exist. Procedure canceled."
        GoTo ExitProc
    End If
    
    ' select recordset from table or query
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = '" & CStr(varRecordsetName) & "'")
    
    ' assign recordset to properties
    ' mvarRecordsetClipboard(inti, 2): property value
    ' mvarRecordsetClipboard(inti, 0): property name
    Dim inti As Integer
    For inti = LBound(mvarRecordsetClipboard, 1) + 1 To UBound(mvarRecordsetClipboard, 1)
        mvarRecordsetClipboard(inti, 2) = rstRecordset.Fields(mvarRecordsetClipboard(inti, 0))
    Next
    
    ' verbatim message: return data content from array
    ' mvarRecordsetClipboard(inti, 2): property value
    ' mvarRecordsetClipboard(inti, 0): property name
    If bolLocalVerbatim = True Then
        For inti = LBound(mvarRecordsetClipboard, 1) To UBound(mvarRecordsetClipboard, 1)
            Debug.Print "clsAuftrag.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
End Sub

Public Property Let AftrID(ByVal varAftrID As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrID (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "AftrID")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varAftrID, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get AftrID() As Variant
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrID (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "AftrID")

    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    AftrID = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let AftrTitel(ByVal varAftrTitel As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "AftrTitel")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varAftrTitel, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get AftrTitel() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.AftrTitel (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "AftrTitel")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    AftrTitel = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

' Sollte nur durch den Import eines ICD-Reports geÃ¤ndert werden
Public Property Let StatusKey(ByVal varStatusKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "StatusKey")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varStatusKey, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get StatusKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.StatusKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "StatusKey")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    StatusKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let PrioritaetKey(ByVal varPrioritaetKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "PrioritaetKey")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPrioritaetKey, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get PrioritaetKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.PrioritaetKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "PrioritaetKey")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    PrioritaetKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let ParentKey(ByVal varParentKey As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "ParentKey")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varParentKey, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get ParentKey() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.ParentKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "ParentKey")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    ParentKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Bemerkung")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varBemerkung, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get Bemerkung() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Bemerkung (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Bemerkung")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let BeginnSoll(ByVal varBeginnSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "BeginnSoll")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varBeginnSoll, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get BeginnSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BeginnSoll (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "BeginnSoll")
    
    ' get value [mvarRecordsetClipboard(intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    BeginnSoll = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let EndeSoll(ByVal varEndeSoll As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "EndeSoll")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varEndeSoll, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get EndeSoll() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.EndeSoll (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "EndeSoll")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    EndeSoll = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let Erstellt(ByVal varErstellt As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Erstellt")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varErstellt, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get Erstellt() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Erstellt (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Erstellt")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    Erstellt = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let kunde(ByVal varKunde As Variant)

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Kunde")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varKunde, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get kunde() As Variant

    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "Kunde")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intItemPosition, 1)]
    kunde = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Let OwnerKey(ByVal varOwnerKey As Variant)
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Owner (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "OwnerKey")
    
    ' transform input [mvarRecordsetClipboard(intItemPosition, 1)], assign to attribut value [mvarRecordsetClipboard( intItemPosition, 2)]
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varOwnerKey, mvarRecordsetClipboard(intItemPosition, 1))
    
End Property

Public Property Get OwnerKey() As Variant
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "OwnerKey")
    
    ' get value [mvarRecordsetClipboard( intItemPosition, 2)], transform data type [mvarRecordsetClipboard(intItemPosition, 1)]
    OwnerKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
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

' search varWanted in two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAuftrag.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsAuftrag.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsAuftrag.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' array style B: (intField, intIndex)
        Case "B"
            Do While avarArray(intField, intIndex) <> varWanted
                If intIndex = UBound(avarArray, 2) Then
                    Debug.Print "clsAuftrag.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsAuftrag.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.FindItemInArray executed"
    End If
    
End Function

' enables to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAuftrag.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsAuftrag.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsAuftrag.CheckDataType: varInput ist null"
        End If
        
        ' return input
        CheckDataType = varInput
        Exit Function
    End If
    
    ' declare output
    Dim varOutput As Variant
    
    Select Case varMode
        Case "string"
            varOutput = (varInput)
        Case "link"
            ' check if varInput is already in link format (#...#),
            ' if not convert to link format
            If Left(varInput, 1) = "#" And Right(varInput, 1) = "#" Then
                varOutput = CStr(varInput)
            Else
                varOutput = "#" + CStr(varInput) + "#"
            End If
        Case "date"
            varOutput = CDate(varInput)
        Case "currency"
            varOutput = CCur(varInput)
        Case "integer"
            varOutput = CInt(varInput)
        Case "boolean"
            varOutput = CBool(varInput)
    End Select
    
    CheckDataType = varOutput
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.CheckDataType executed"
    End If

End Function

Public Sub CreateRecordset()
    ' Error Code 1: "AftrID" input is missing
    ' Error Code 2: "AftrID" is taken, choose another value
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAuftrag.CreateRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAuftrag"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "AftrID"
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        ' error message
        MsgBox "AftrID ist ein Pflichtfeld, bitte Wert zuweisen!", vbCritical, "Fehlermeldung"
        Debug.Print "Error: clsAuftrag.CreateRecordset aborted, Error Code 1"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    ' check if primary key value is taken
    If DCount(strPrimaryKey, strTableName, strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'") > 0 Then
        MsgBox mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & " wird bereits verwendet, bitte einen anderen Wert wählen.", vbCritical, "Fehlermeldung"
        Debug.Print "Error: clsAuftrag.CreateRecordset aborted, Error Code 2"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    rstRecordset.AddNew
        rstRecordset!AftrID = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrID"), intValuePosition)
        rstRecordset!AftrTitel = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrTitel"), intValuePosition)
        rstRecordset!StatusKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "StatusKey"), intValuePosition)
        rstRecordset!OwnerKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "OwnerKey"), intValuePosition)
        rstRecordset!PrioritaetKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "PrioritaetKey"), intValuePosition)
        rstRecordset!ParentKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "ParentKey"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BeginnSoll = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BeginnSoll"), intValuePosition)
        rstRecordset!EndeSoll = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "EndeSoll"), intValuePosition)
        rstRecordset!Erstellt = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Erstellt"), intValuePosition)
        rstRecordset!kunde = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "kunde"), intValuePosition)
    rstRecordset.Update
    
    MsgBox "Datensatz erzeugt", vbOKOnly, "Auftrag erstellen"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAuftrag.CreateRecordset executed"
    End If
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Public Sub DeleteRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAuftrag.DeleteRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAuftrag"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "AftrID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsAuftrag.DeleteRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    ' delete recordset
    rstRecordset.Delete
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAuftrag.DeleteRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    
End Sub

Public Sub SaveRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAuftrag.SaveRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAuftrag"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "AftrID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsAuftrag.SaveRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    Dim strRecordsetName As String
    strRecordsetName = rstRecordset.Name
    
    rstRecordset.Edit
        rstRecordset!AftrTitel = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrTitel"), intValuePosition)
        rstRecordset!StatusKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "StatusKey"), intValuePosition)
        rstRecordset!OwnerKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "OwnerKey"), intValuePosition)
        rstRecordset!PrioritaetKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "PrioritaetKey"), intValuePosition)
        rstRecordset!ParentKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "ParentKey"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BeginnSoll = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BeginnSoll"), intValuePosition)
        rstRecordset!EndeSoll = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "EndeSoll"), intValuePosition)
        rstRecordset!Erstellt = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Erstellt"), intValuePosition)
        rstRecordset!kunde = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "kunde"), intValuePosition)
    rstRecordset.Update
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAuftrag.SaveRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

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
