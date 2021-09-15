VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLeistungserfassungsblatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLeistungserfassungsblatt

Option Compare Database
Option Explicit

' container for selected recordset, is also relevant for properties procedures
' lookup DefaultProperties for array composition
Private mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

Private Function DefaultProperties()

    mbolVerbatim = False
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 5
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "LeistungserfassungsblattID"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "RechnungNr"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "Bemerkung"
        varDefaultProperties(3, 1) = "string"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "BelegID"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "Brutto"
        varDefaultProperties(5, 1) = "currency"
        varDefaultProperties(5, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)
        
    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt.DefaultProperties executed"
    End If

End Function

' search varWanted in two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsLeistungserfassungsblatt.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsLeistungserfassungsblatt.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsLeistungserfassungsblatt.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt.FindItemInArray executed"
    End If
    
End Function

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' command message
    If mbolVerbatim Then
        Debug.Print "execute clsLeistungserfassungsblatt.SelectRecordset"
    End If

    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblLeistungserfassungsblatt"
    
    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"
    
    ' name of the column that contains the sought recordset
    Dim strKeyColumn As String
    strKeyColumn = "LeistungserfassungsblattID"
    
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
            Debug.Print "clsLeistungserfassungsblatt.SelectRecordset: varRecordsetName is empty, " _
                & "prompt default value"
        End If
        ' load default Value
        mvarRecordsetClipboard = DefaultProperties
        GoTo ExitProc
    End If
    
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsLeistungserfassungsblatt.SelectRecordset: " & strTableName & " does not exist. Procedure canceled."
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
            Debug.Print "clsLeistungserfassungsblatt.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
End Sub

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
        Debug.Print "execute clsLeistungserfassungsblatt.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsLeistungserfassungsblatt.CheckDataType: varInput ist null"
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
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt.CheckDataType executed"
    End If

End Function
Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "initialize clsLeistungserfassungsblatt"
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
    If gconVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt initialized"
    End If
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsLeistungserfassungsblatt terminated"
End Sub

Public Property Let Leistungserfassungsblatt(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "LeistungserfassungsblattID"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Let LeistungserfassungsblattID"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Let LeistungserfassungsblattID executed"
    End If
    
End Property

Public Property Get LeistungserfassungsblattID() As Variant

    Dim strPropertyName As String
    strPropertyName = "LeistungserfassungsblattID"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Get LeistungserfassungsblattID"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    LeistungserfassungsblattID = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Get LeistungserfassungsblattID executed"
    End If
    
End Property

Public Property Let RechnungNr(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "RechnungNr"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Let RechnungNr"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Let RechnungNr executed"
    End If
    
End Property

Public Property Get RechnungNr() As Variant

    Dim strPropertyName As String
    strPropertyName = "RechnungNr"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Get RechnungNr"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    RechnungNr = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Get RechnungNr executed"
    End If
    
End Property

Public Property Let Bemerkung(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Let Bemerkung"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Let Bemerkung executed"
    End If
    
End Property

Public Property Get Bemerkung() As Variant

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Get Bemerkung"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Get Bemerkung executed"
    End If
    
End Property

Public Property Let BelegID(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "BelegID"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Let BelegID"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Let BelegID executed"
    End If
    
End Property

Public Property Get BelegID() As Variant

    Dim strPropertyName As String
    strPropertyName = "BelegID"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Get BelegID"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    BelegID = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Get BelegID executed"
    End If
    
End Property

Public Property Let Brutto(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "Brutto"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Let Brutto"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Let Brutto executed"
    End If
    
End Property

Public Property Get Brutto() As Variant

    Dim strPropertyName As String
    strPropertyName = "Brutto"
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt Get Brutto"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    Brutto = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt Get Brutto executed"
    End If
    
End Property

Public Sub CreateRecordset()
    ' Error Code 1: "Leistungserfassungsblatt" input is missing
    ' Error Code 2: "Leistungserfassungsblatt" is taken, choose another value
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsLeistungserfassungsblatt.CreateRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "LeistungserfassungsblattID"
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        ' error message
        MsgBox strPrimaryKey & " ist ein Pflichtfeld, bitte Wert zuweisen!", vbCritical, "Fehlermeldung"
        Debug.Print "Error: clsLeistungserfassungsblatt.CreateRecordset aborted, Error Code 1"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    ' check if primary key value is taken
    If DCount(strPrimaryKey, strTableName, strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'") > 0 Then
        MsgBox mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & " wird bereits verwendet, bitte einen anderen Wert wählen.", vbCritical, "Fehlermeldung"
        Debug.Print "Error: clsLeistungserfassungsblatt.CreateRecordset aborted, Error Code 2"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    rstRecordset.AddNew
        rstRecordset!LeistungserfassungsblattID = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "LeistungserfassungsblattID"), intValuePosition)
        rstRecordset!RechnungNr = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "RechnungNr"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BelegID = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BelegID"), intValuePosition)
        rstRecordset!Brutto = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Brutto"), intValuePosition)
        ' class changed? Do not miss to update the class properties!
    rstRecordset.Update
    
    MsgBox "Datensatz erzeugt", vbOKOnly, "Leistungserfassungsblatt erstellen"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.CreateRecordset executed"
    End If
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Public Sub DeleteRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsLeistungserfassungsblatt.DeleteRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "LeistungserfassungsblattID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsLeistungserfassungsblatt.DeleteRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    ' delete recordset
    rstRecordset.Delete
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.DeleteRecordset executed"
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
        Debug.Print "execute clsLeistungserfassungsblatt.SaveRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "LeistungserfassungsblattID"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsLeistungserfassungsblatt.SaveRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    Dim strRecordsetName As String
    strRecordsetName = rstRecordset.Name
    
    rstRecordset.Edit
        rstRecordset!RechnungNr = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "RechnungNr"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BelegID = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BelegID"), intValuePosition)
        rstRecordset!Brutto = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Brutto"), intValuePosition)
    rstRecordset.Update
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsLeistungserfassungsblatt.SaveRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub
