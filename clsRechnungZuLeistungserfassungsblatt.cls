VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRechnungZuLeistungserfassungsblatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsRechnungZuLeistungserfassungsblatt

Option Compare Database
Option Explicit

' conainter for selected recordset, is also relevant for properties procedures
' lookup DefaultProperties for array composition
Private mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

Private Sub Class_Initialize()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "initialize clsRechnungZuLeistungserfassungsblatt"
    End If
    
    Dim strQryChecksumName As String
    strQryChecksumName = "qryChecksumRechnungZuLeistungserfassungsblatt"
    
    Dim strObjectType As String
    strObjectType = "query"
    
    If Not (ObjectExists(strQryChecksumName, strObjectType)) Then
        BuildCheckSum
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt initialized"
    End If
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsRechnungZuLeistungserfassungsblatt terminated"
End Sub

Private Function DefaultProperties()

    mbolVerbatim = False

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 4
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "ID"
        varDefaultProperties(1, 1) = "long"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "RefRechnungNr"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "RefLeistungserfassungsblattID"
        varDefaultProperties(3, 1) = "string"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "Bemerkung"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)

    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.DefaultProperties executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsRechnungZuLeistungserfassungsblatt.CheckDataType: varInput ist null"
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
        Case "long"
            varOutput = CLng(varInput)
    End Select
    
    CheckDataType = varOutput
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.CheckDataType executed"
    End If

End Function

' search varWanted in two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsRechnungZuLeistungserfassungsblatt.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsRechnungZuLeistungserfassungsblatt.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsRechnungZuLeistungserfassungsblatt.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.FindItemInArray executed"
    End If
    
End Function

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
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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

Public Property Let ID(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "ID"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Let ID"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Let ID executed"
    End If
    
End Property

Public Property Get ID() As Variant

    Dim strPropertyName As String
    strPropertyName = "ID"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Get ID"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    ID = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Get ID executed"
    End If
    
End Property

Public Property Let RefRechnungNr(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "RefRechnungNr"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Let RefRechnungNr"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Let RefRechnungNr executed"
    End If
    
End Property

Public Property Get RefRechnungNr() As Variant

    Dim strPropertyName As String
    strPropertyName = "RefRechnungNr"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Get RefRechnungNr"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    RefRechnungNr = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Get RefRechnungNr executed"
    End If
    
End Property

Public Property Let RefLeistungserfassungsblattID(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "RefLeistungserfassungsblattID"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Let RefLeistungserfassungsblattID"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Let RefLeistungserfassungsblattID executed"
    End If
    
End Property

Public Property Get RefLeistungserfassungsblattID() As Variant

    Dim strPropertyName As String
    strPropertyName = "RefLeistungserfassungsblattID"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Get RefLeistungserfassungsblattID"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    RefLeistungserfassungsblattID = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Get RefLeistungserfassungsblattID executed"
    End If
    
End Property

Public Property Let Bemerkung(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Let Bemerkung"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Let Bemerkung executed"
    End If
    
End Property

Public Property Get Bemerkung() As Variant

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt Get Bemerkung"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt Get Bemerkung executed"
    End If
    
End Property

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' command message
    If mbolVerbatim Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.SelectRecordset"
    End If

    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblRechnungZuLeistungserfassungsblatt"
    
    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"
    
    ' name of the column that contains the primary key
    Dim strKeyColumn As String
    strKeyColumn = "ID"
    
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
            Debug.Print "clsRechnungZuLeistungserfassungsblatt.SelectRecordset: varRecordsetName is empty, " _
                & "prompt default value"
        End If
        ' load default Value
        mvarRecordsetClipboard = DefaultProperties
        Exit Sub
    End If
    
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.SelectRecordset: " & strTableName & " does not exist. Procedure canceled."
        Exit Sub
    End If
    
    ' select recordset from table or query
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = " & varRecordsetName)
    
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
            Debug.Print "clsRechnungZuLeistungserfassungsblatt.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
        
    ' event message
    If mbolVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.SelectRecordset executed"
    End If
    
End Sub

Public Sub CreateRecordset()
' Error Code 1: relationship between recordsets was recorded earlier
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.CreateRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblRechnungZuLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "ID"
    
    Dim strField01Name As String
    strField01Name = "RefRechnungNr"
    
    Dim strField02Name As String
    strField02Name = "RefLeistungserfassungsblattID"
    
    Dim strDomainName As String
    strDomainName = "qryChecksumRechnungZuLeistungserfassungsblatt"
    
    Dim strFieldName03 As String
    strFieldName03 = "checksum"
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    Dim intValuePosition As Integer
    intValuePosition = 2
        
    ' get property values from mvarRecordsetClipboard
    Dim varField01Value As Variant
    varField01Value = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strField01Name), intValuePosition)
    
    Dim varField02Value As Variant
    varField02Value = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strField02Name), intValuePosition)
    
    ' check if input is null
    If IsNull(varField01Value) Or IsNull(varField02Value) Then
        Debug.Print "Error: clsRechnungZuLeistungserfassungsblatt.CreateRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
    
    ' check if relationship already exists
    Dim strDummy As String
    strDummy = varField01Value & varField02Value
    
    ' check if relationship RechnungZuLeistungserfassungsblatt is taken
    If DCount("[" & strFieldName03 & "]", strDomainName, "[" & strFieldName03 & "] Like '" & strDummy & "'") > 0 Then
        MsgBox "Das Leistungserfassungsblatt '" & varField02Value & "' wurde bereits mit der Rechnung Nr '" & varField01Value & "' verknüpft.", vbCritical, "Speichern"
        Debug.Print "Error: clsRechnungZuLeistungserfassungsblatt.CreateRecordset, Error Code 1"
        GoTo ExitProc
    End If
    
    ' generate primary key
    Dim lngPrimaryKey As Long
    lngPrimaryKey = Nz(DMax(strPrimaryKey, strTableName)) + 1
    
    ' create recordset
    rstRecordset.AddNew
            rstRecordset!ID = lngPrimaryKey
            rstRecordset!RefRechnungNr = varField01Value
            rstRecordset!RefLeistungserfassungsblattID = varField02Value
            rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
            ' class changed? Do not miss to update the class properties!
        rstRecordset.Update
    
    ' success message
    MsgBox "Das Leistungserfassungsblatt '" & varField02Value & "' wurde der Rechnung Nr '" & varField01Value & "' zugeordnet.", vbInformation, "Beziehung erfasst"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.CreateRecordset executed"
    End If
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Public Sub DeleteRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.DeleteRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblRechnungZuLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "ID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsRechnungZuLeistungserfassungsblatt.DeleteRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    ' delete recordset
    rstRecordset.Delete
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.DeleteRecordset executed"
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
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.SaveRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblRechnungZuLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "ID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsRechnungZuLeistungserfassungsblatt.SaveRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    Dim strRecordsetName As String
    strRecordsetName = rstRecordset.Name
    
    rstRecordset.Edit
        rstRecordset!RefRechnungNr = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "RefRechnungNr"), intValuePosition)
        rstRecordset!RefLeistungserfassungsblattID = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "RefLeistungserfassungsblattID"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
    rstRecordset.Update
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.SaveRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Private Sub BuildCheckSum()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsRechnungZuLeistungserfassungsblatt.SearchEinzelauftrag"
    End If
    
    Dim strDomainName As String
    strDomainName = "tblRechnungZuLeistungserfassungsblatt"
    
    Dim strCheckField01 As String
    strCheckField01 = "RefRechnungNr"
    
    Dim strCheckField02 As String
    strCheckField02 = "RefLeistungserfassungsblattID"
    
    Dim strQueryName As String
    strQueryName = "qryChecksumRechnungZuLeistungserfassungsblatt"
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT [" & strDomainName & "]![" & strCheckField01 & "] & [" & strDomainName & "]![" & strCheckField02 & "] AS checksum" & _
                    " FROM " & strDomainName & ";"
    End With
    
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With
    
ExitProc:
    qdfQuery.Close
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsRechnungZuLeistungserfassungsblatt.SearchEinzelauftrag executed"
    End If
    
End Sub



