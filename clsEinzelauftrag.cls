VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEinzelauftrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsEinzelauftrag

Option Compare Database
Option Explicit

'Hauptdatensatz der Klasse
Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher fuer die ID des
' mit dem Cursor ausgewaehlten Datensatzes im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' conainter for selected recordset, is also relevant for properties procedures
' lookup DefaultProperties for array composition
Private mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "clsEinzelauftrag initialized"
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsEinzelauftrag terminated"
End Sub

Private Function DefaultProperties()

    mbolVerbatim = False

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 15
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "EAkurzKey"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "MengengeruestLink"
        varDefaultProperties(2, 1) = "link"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "LeistungsbeschreibungLink"
        varDefaultProperties(3, 1) = "link"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "Bemerkung"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "BeauftragtDatum"
        varDefaultProperties(5, 1) = "date"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "AbgebrochenDatum"
        varDefaultProperties(6, 1) = "date"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "MitzeichnungI21Datum"
        varDefaultProperties(7, 1) = "date"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "MitzeichnungI25Datum"
        varDefaultProperties(8, 1) = "date"
        varDefaultProperties(8, 2) = Null
    varDefaultProperties(9, 0) = "AngebotDatum"
        varDefaultProperties(9, 1) = "date"
        varDefaultProperties(9, 2) = Null
    varDefaultProperties(10, 0) = "AbgenommenDatum"
        varDefaultProperties(10, 1) = "date"
        varDefaultProperties(10, 2) = Null
    varDefaultProperties(11, 0) = "StorniertDatum"
        varDefaultProperties(11, 1) = "date"
        varDefaultProperties(11, 2) = Null
    varDefaultProperties(12, 0) = "AngebotBrutto"
        varDefaultProperties(12, 1) = "currency"
        varDefaultProperties(12, 2) = Null
    varDefaultProperties(13, 0) = "BWIKey"
        varDefaultProperties(13, 1) = "string"
        varDefaultProperties(13, 2) = Null
    varDefaultProperties(14, 0) = "AftrBeginn"
        varDefaultProperties(14, 1) = "currency"
        varDefaultProperties(14, 2) = Null
    varDefaultProperties(15, 0) = "AftrEnde"
        varDefaultProperties(15, 1) = "string"
        varDefaultProperties(15, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)

    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag.DefaultProperties executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsEinzelauftrag.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsEinzelauftrag.CheckDataType: varInput ist null"
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
        Debug.Print "clsEinzelauftrag.CheckDataType executed"
    End If

End Function

' search varWanted in two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsEinzelauftrag.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsEinzelauftrag.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsEinzelauftrag.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsEinzelauftrag.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag.FindItemInArray executed"
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
        Debug.Print "execute clsEinzelauftrag.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsEinzelauftrag.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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

Public Property Let EAkurzKey(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "EAkurzKey"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let EAkurzKey"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let EAkurzKey executed"
    End If
    
End Property

Public Property Get EAkurzKey() As Variant

    Dim strPropertyName As String
    strPropertyName = "EAkurzKey"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get EAkurzKey"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    EAkurzKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get EAkurzKey executed"
    End If
    
End Property

Public Property Let MengengeruestLink(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "MengengeruestLink"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let MengengeruestLink"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let MengengeruestLink executed"
    End If
    
End Property

Public Property Get MengengeruestLink() As Variant

    Dim strPropertyName As String
    strPropertyName = "MengengeruestLink"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get MengengeruestLink"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    MengengeruestLink = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get MengengeruestLink executed"
    End If
    
End Property

Public Property Let LeistungsbeschreibungLink(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "LeistungsbeschreibungLink"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let LeistungsbeschreibungLink"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let LeistungsbeschreibungLink executed"
    End If
    
End Property

Public Property Get LeistungsbeschreibungLink() As Variant

    Dim strPropertyName As String
    strPropertyName = "LeistungsbeschreibungLink"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get LeistungsbeschreibungLink"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    LeistungsbeschreibungLink = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get LeistungsbeschreibungLink executed"
    End If
    
End Property

Public Property Let Bemerkung(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let Bemerkung"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let Bemerkung executed"
    End If
    
End Property

Public Property Get Bemerkung() As Variant

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get Bemerkung"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get Bemerkung executed"
    End If
    
End Property

Public Property Let BeauftragtDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "BeauftragtDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let BeauftragtDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let BeauftragtDatum executed"
    End If
    
End Property

Public Property Get BeauftragtDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "BeauftragtDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get BeauftragtDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    BeauftragtDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get BeauftragtDatum executed"
    End If
    
End Property

Public Property Let AbgebrochenDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AbgebrochenDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AbgebrochenDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AbgebrochenDatum executed"
    End If
    
End Property

Public Property Get AbgebrochenDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AbgebrochenDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AbgebrochenDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AbgebrochenDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AbgebrochenDatum executed"
    End If
    
End Property

Public Property Let MitzeichnungI21Datum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "MitzeichnungI21Datum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let MitzeichnungI21Datum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let MitzeichnungI21Datum executed"
    End If
    
End Property

Public Property Get MitzeichnungI21Datum() As Variant

    Dim strPropertyName As String
    strPropertyName = "MitzeichnungI21Datum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get MitzeichnungI21Datum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    MitzeichnungI21Datum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get MitzeichnungI21Datum executed"
    End If
    
End Property

Public Property Let MitzeichnungI25Datum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "MitzeichnungI25Datum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let MitzeichnungI25Datum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let MitzeichnungI25Datum executed"
    End If
    
End Property

Public Property Get MitzeichnungI25Datum() As Variant

    Dim strPropertyName As String
    strPropertyName = "MitzeichnungI25Datum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get MitzeichnungI25Datum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    MitzeichnungI25Datum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get MitzeichnungI25Datum executed"
    End If
    
End Property

Public Property Let AngebotDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AngebotDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AngebotDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AngebotDatum executed"
    End If
    
End Property

Public Property Get AngebotDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AngebotDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AngebotDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AngebotDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AngebotDatum executed"
    End If
    
End Property

Public Property Let AbgenommenDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AbgenommenDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AbgenommenDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AbgenommenDatum executed"
    End If
    
End Property

Public Property Get AbgenommenDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AbgenommenDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AbgenommenDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AbgenommenDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AbgenommenDatum executed"
    End If
    
End Property

Public Property Let StorniertDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "StorniertDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let StorniertDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let StorniertDatum executed"
    End If
    
End Property

Public Property Get StorniertDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "StorniertDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get StorniertDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    StorniertDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get StorniertDatum executed"
    End If
    
End Property

Public Property Let AngebotBrutto(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AngebotBrutto"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AngebotBrutto"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AngebotBrutto executed"
    End If
    
End Property

Public Property Get AngebotBrutto() As Variant

    Dim strPropertyName As String
    strPropertyName = "AngebotBrutto"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AngebotBrutto"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AngebotBrutto = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AngebotBrutto executed"
    End If
    
End Property

Public Property Let BWIKey(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "BWIKey"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let BWIKey"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let BWIKey executed"
    End If
    
End Property

Public Property Get BWIKey() As Variant

    Dim strPropertyName As String
    strPropertyName = "BWIKey"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get BWIKey"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    BWIKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get BWIKey executed"
    End If
    
End Property

Public Property Let AftrBeginn(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AftrBeginn"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AftrBeginn"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AftrBeginn executed"
    End If
    
End Property

Public Property Get AftrBeginn() As Variant

    Dim strPropertyName As String
    strPropertyName = "AftrBeginn"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AftrBeginn"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AftrBeginn = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AftrBeginn executed"
    End If
    
End Property

Public Property Let AftrEnde(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AftrEnde"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Let AftrEnde"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Let AftrEnde executed"
    End If
    
End Property

Public Property Get AftrEnde() As Variant

    Dim strPropertyName As String
    strPropertyName = "AftrEnde"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag Get AftrEnde"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AftrEnde = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag Get AftrEnde executed"
    End If
    
End Property
