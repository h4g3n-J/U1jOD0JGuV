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

' dient als Zwischenspeicher fuer die ID des aktuell
' mit dem Cursor ausgewaehlten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

    Private mvarRechnungNr As Variant
    Private mvarBemerkung As Variant
    Private mvarRechnungLink As Variant
    Private mvarTechnischRichtigDatum As Variant
    Private mvarIstTeilrechnung As Variant
    Private mvarIstSchlussrechnung As Variant
    Private mvarKalkulationLNWLink As Variant
    Private mvarRechnungBrutto As Variant

Private Sub default()
    mvarRechnungNr = Null
    mvarBemerkung = Null
    mvarRechnungLink = Null
    mvarTechnischRichtigDatum = Null
    mvarIstTeilrechnung = Null
    mvarIstSchlussrechnung = Null
    mvarKalkulationLNWLink = Null
    mvarRechnungBrutto = Null
End Sub

Private Sub Class_Initialize()
    Debug.Print "Klasse clsRechnung instanziiert."
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsRechnung terminiert."
End Sub

Public Property Get RechnungNr() As Variant
    RechnungNr = basSupport.PruefeString(mvarRechnungNr)
End Property

Public Property Let RechnungNr(ByVal varRechnungNr As Variant)
    mvarRechnungNr = basSupport.PruefeString(varRechnungNr)
End Property

Public Property Get Bemerkung() As Variant
    Bemerkung = basSupport.PruefeString(mvarBemerkung)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    mvarBemerkung = basSupport.PruefeString(varBemerkung)
End Property

Public Property Get RechnungLink() As Variant
    RechnungLink = basSupport.PruefeString(mvarRechnungLink)
End Property

Public Property Let RechnungLink(ByVal varRechnungLink As Variant)
    mvarRechnungLink = basSupport.PruefeString(varRechnungLink)
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

Public Property Get Verbatim() As Variant
    Verbatim = basSupport.PruefeBoolean(mbolVerbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolVerbatim = basSupport.PruefeBoolean(varVerbatim)
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
        mvarRecordsetClipboard = AngebotDefaultProperties
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
    If gconVerbatim = True Then
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
    If gconVerbatim = True Then
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
    If gconVerbatim = True Then
        Debug.Print "clsRechnung.FindItemInArray executed"
    End If
    
End Function

' enables to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If gconVerbatim = True Then
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
    If gconVerbatim = True Then
        Debug.Print "clsRechnung.CheckDataType execute"
    End If

End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsRechnung.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsRechnung.CheckDataType: varInput ist null"
        End If
        
        ' return input
        CheckDataType = varInput
        
        ' event message
        Debug.Print "clsRechnung.CheckDatyType executed"
        Exit Function
    End If
    
    ' declare output
    Dim varOutput As Variant
    
    Select Case varMode
        Case "string"
            varOutput = (varInput)
        Case "link"
            ' check if varInput already is a link (#...#),
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
        Debug.Print "clsAngebot.CheckDataType execute"
    End If
    
End Function

' search varWanted in a two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant

    ' command message
    If gconVerbatim = True Then
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
                    Debug.Print "clsAngebot.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsAngebot.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsAngebot.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
        
         ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsRechnung.FindItemInArray executed"
    End If
    
End Function




