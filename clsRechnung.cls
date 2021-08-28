VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRechnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsRechnung

Option Compare Database
Option Explicit

'Hauptdatensatz der Klasse
Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher fuer die ID des
' mit dem Cursor ausgewaehlten Datensatzes im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant
Private mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

Private Function DefaultProperties()

    mbolVerbatim = False

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnung.DefaultProperties"
    End If
    
    Dim varDefaultProperties(13, 2) As Variant
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "RechnungNr"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "Bemerkung"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "RechnungLink"
        varDefaultProperties(3, 1) = "link"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "TechnischRichtigDatum"
        varDefaultProperties(4, 1) = "date"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "IstTeilrechnung"
        varDefaultProperties(5, 1) = "boolean"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "IstSchlussrechnung"
        varDefaultProperties(6, 1) = "date"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "KalkulationLNWLink"
        varDefaultProperties(7, 1) = "link"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "RechnungBrutto"
        varDefaultProperties(8, 1) = "currency"
        varDefaultProperties(8, 2) = Null

    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnung.DefaultProperties executed"
    End If
    
End Function
Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "clsAngebot initialized"
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsRechnung terminated"
End Sub

Public Property Let RechnungNr(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "RechnungNr"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnung Let RechnungNr"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnung Let RechnungNr executed"
    End If
    
End Property

Public Property Get RechnungNr() As Variant

    Dim strPropertyName As String
    strPropertyName = "RechnungNr"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsrechnung Get RechnungNr"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    EndeSoll = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsrechnung Get RechnungNr executed"
    End If
    
End Property

Public Property Let Bemerkung(ByVal varPropertyInput As Variant)
    
    Dim strPropertyName As String
    strPropertyName = "Bemerkung"
    
    'command message
    If gconVerbatim = True Then
        Debug.Print "execute clsRechnung Let Bemerkung"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsRechnung Let Bemerkung executed"
    End If
    
End Property

Public Property Get Bemerkung() As Variant
    
    Dim strPropertyName As String
    strPropertyName = "Bemerkung"

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsrechnung Get Bemerkung"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    EndeSoll = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsRechnung Get Bemerkung executed"
        
    End If
End Property

Public Property Let RechnungLink(ByVal varPropertyInput As Variant)
    
    Dim strPropertyName As String
    strPropertyName = "RechnungLink"
    
    'command message
    If gconVerbatim = True Then
        Debug.Print "execute clsRechnung Let RechnungLink"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsRechnung Let RechnungLink executed"
    End If
    
End Property

Public Property Get RechnungLink() As Variant
    
    Dim strPropertyName As String
    strPropertyName = "RechnungLink"

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsrechnung Get RechnungLink"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    EndeSoll = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsRechnung Get RechnungLink executed"
    End If
    
End Property

Public Property Get TechnischRichtigDatum() As Variant
    TechnischRichtigDatum = basSupport.PruefeString(mvarTechnischRichtigDatum)
End Property

Public Property Let TechnischRichtigDatum(ByVal varTechnischRichtigDatum As Variant)
    mvarTechnischRichtigDatum = basSupport.PruefeString(varTechnischRichtigDatum)
End Property

Public Property Get IstTeilrechnung() As Variant
    IstTeilrechnung = basSupport.PruefeString(mvarIstTeilrechnung)
End Property

Public Property Let IstTeilrechnung(ByVal varIstTeilrechnung As Variant)
    mvarIstTeilrechnung = basSupport.PruefeString(varIstTeilrechnung)
End Property

Public Property Get IstSchlussrechnung() As Variant
    IstSchlussrechnung = basSupport.PruefeString(mvarIstSchlussrechnung)
End Property

Public Property Let IstSchlussrechnung(ByVal varIstSchlussrechnung As Variant)
    mvarIstSchlussrechnung = basSupport.PruefeString(varIstSchlussrechnung)
End Property

Public Property Get KalkulationLNWLink() As Variant
    KalkulationLNWLink = basSupport.PruefeString(mvarKalkulationLNWLink)
End Property

Public Property Let KalkulationLNWLink(ByVal varKalkulationLNWLink As Variant)
    mvarKalkulationLNWLink = basSupport.PruefeString(varKalkulationLNWLink)
End Property

Public Property Get RechnungBrutto() As Variant
    RechnungBrutto = basSupport.PruefeString(mvarRechnungBrutto)
End Property

Public Property Let RechnungBrutto(ByVal varRechnungBrutto As Variant)
    mvarRechnungBrutto = basSupport.PruefeString(varRechnungBrutto)
End Property

Public Sub AddRecordset()
       
mvarClipboard = basSupport.AddRecordsetMN( _
    "tblAngebot", "BWIKey", "Angebot Titel", "Titel des referenzierten Angebots eingeben", "Angebot erfassen", _
    "tblRechnung", "RechnungNr", "Rechnungsnummer", "Rechnungsnummer eingeben", "Rechnung erfassen", _
    "tblAngebotZuRechnung", _
    False)
    
End Sub

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblRechnung"
    
    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"
    
    ' name of the column that contains the recordset
    Dim strKeyColumn As String
    strKeyColumn = "RechnungNr"
    
    ' define database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' define recordset
    Dim rstRecordset As DAO.RecordSet
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' error handler if no recordset is selected
    ' end exitProc
    If IsNull(varRecordsetName) Then
        ' verbatim message
        If bolLocalVerbatim = True Then
            Debug.Print "clsRechnung.SelectRecordset: varRecordsetName ist leer, " _
                & "Defaultwert anzeigen"
        End If
        ' default Value
        mvarRecordsetClipboard = DefaultProperties
        GoTo ExitProc
    End If
    
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsRechnung.SelectRecordset: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
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
        Debug.Print "execute clsRechnung.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "basclsRechnung.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnung.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsRechnung.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsRechnung.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsRechnung.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsRechnung.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsRechnung.FindItemInArray executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsRechnung.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsRechnung.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsRechnung.CheckDataType: varInput ist null"
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
        Debug.Print "clsRechnung.CheckDataType execute"
    End If

End Function



