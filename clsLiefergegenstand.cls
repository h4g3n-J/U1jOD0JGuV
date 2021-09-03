VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLiefergegenstand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLiefergegenstand

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

Public Sub default()
    mvarLiefergegenstandID = Null
    mvarAftrID = Null
    mvarBezeichnung = Null
    mvarLieferscheinLink = Null
    mvarSeriennummer = Null
    mvarAnzahl = Null
    mvarHerstellerkennzeichen = Null
    mvarUnterauftragnehmerAngebotLink = Null
    mvarAngebotNetto = Null
    mvarPreisBrutto = Null
    mvarLieferdatum = Null
    mvarLieferschein = Null
    mvarBemerkung = Null
    mvarZielAftrID = Null
    mvarZielLieferschein = Null
    mvarZielLinklieferschein = Null
    mvarLieferdatum = Null
    mvarLiefergegenstandLink = Null
    mvarIstReserve = Null
    mvarAftrTitel = Null
    mvarBWIKey = Null
    mvarLeistungsbeschreibungLink = Null
    mvarParentKey = Null
End Sub

Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "clsLiefergegenstand initialized"
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsLiefergegenstand terminated"
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
        Debug.Print "execute clsLiefergegenstand.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsLiefergegenstand.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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
        Debug.Print "execute clsLiefergegenstand.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsLiefergegenstand.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsLiefergegenstand.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
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
                    Debug.Print "clsLiefergegenstand.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsLiefergegenstand.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLiefergegenstand.FindItemInArray executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLiefergegenstand.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsLiefergegenstand.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsLiefergegenstand.CheckDataType: varInput ist null"
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
        Debug.Print "clsLiefergegenstand.CheckDataType executed"
    End If

End Function

Private Function DefaultProperties()

    mbolVerbatim = False

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLiefergegenstand.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 20
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "LiefergegenstandID"
        varDefaultProperties(1, 1) = "integer"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "aftrIdKey"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "bezeichnung"
        varDefaultProperties(3, 1) = "link"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "linkLieferschein"
        varDefaultProperties(4, 1) = "link"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "seriennummer"
        varDefaultProperties(5, 1) = "string"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "anzahl"
        varDefaultProperties(6, 1) = "integer"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "herstellerkennzeichen"
        varDefaultProperties(7, 1) = "string"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "Uanangebot"
        varDefaultProperties(8, 1) = "link"
        varDefaultProperties(8, 2) = Null
    varDefaultProperties(9, 0) = "angebotNetto"
        varDefaultProperties(9, 1) = "currency"
        varDefaultProperties(9, 2) = Null
    varDefaultProperties(10, 0) = "preisBrutto"
        varDefaultProperties(10, 1) = "currency"
        varDefaultProperties(10, 2) = Null
    varDefaultProperties(11, 0) = "lieferdatum"
        varDefaultProperties(11, 1) = "date"
        varDefaultProperties(11, 2) = Null
    varDefaultProperties(12, 0) = "lieferschein"
        varDefaultProperties(12, 1) = "string"
        varDefaultProperties(12, 2) = Null
    varDefaultProperties(13, 0) = "bemerkung"
        varDefaultProperties(13, 1) = "sring"
        varDefaultProperties(13, 2) = Null
    varDefaultProperties(14, 0) = "zielAftrID"
        varDefaultProperties(14, 1) = "string"
        varDefaultProperties(14, 2) = Null
    varDefaultProperties(15, 0) = "zielLieferschein"
        varDefaultProperties(15, 1) = "string"
        varDefaultProperties(15, 2) = Null
    varDefaultProperties(16, 0) = "zielLinkLieferschein"
        varDefaultProperties(16, 1) = "link"
        varDefaultProperties(16, 2) = Null
    varDefaultProperties(17, 0) = "zielLieferdatum"
        varDefaultProperties(17, 1) = "date"
        varDefaultProperties(17, 2) = Null
    varDefaultProperties(18, 0) = "LiefergegenstandLink"
        varDefaultProperties(18, 1) = "link"
        varDefaultProperties(18, 2) = Null
    varDefaultProperties(19, 0) = "IstReserve"
        varDefaultProperties(19, 1) = "boolean"
        varDefaultProperties(19, 2) = Null
    varDefaultProperties(20, 0) = "seriennummer2"
        varDefaultProperties(20, 1) = "string"
        varDefaultProperties(20, 2) = Null
    
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)

    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLiefergegenstand.DefaultProperties executed"
    End If
    
End Function

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)

    ' command message
    If mbolVerbatim Then
        Debug.Print "execute clsLiefergegenstand.SelectRecordset"
    End If

    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblLiefergegenstand"
    
    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"
    
    ' name of the column that contains the primary key
    Dim strKeyColumn As String
    strKeyColumn = "LiefergegenstandID"
    
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
            Debug.Print "clsLiefergegenstand.SelectRecordset: varRecordsetName is empty, " _
                & "prompt default value"
        End If
        ' load default Value
        mvarRecordsetClipboard = DefaultProperties
        GoTo ExitProc
    End If
    
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsLiefergegenstand.SelectRecordset: " & strTableName & " does not exist. Procedure canceled."
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
            Debug.Print "clsLiefergegenstand.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
End Sub

Public Property Get LiefergegenstandID() As Variant
    LiefergegenstandID = PruefeLong(mvarLiefergegenstandID)
End Property

Public Property Let LiefergegenstandID(ByVal varLiefergegenstandID As Variant)
    mvarLiefergegenstandID = PruefeLong(varLiefergegenstandID)
End Property

Public Property Get AftrID() As Variant
    AftrID = PruefeString(mvarAftrID)
End Property

Public Property Let AftrID(ByVal varAftrID As Variant)
    mvarAftrID = PruefeString(varAftrID)
End Property

Public Property Get bezeichnung() As Variant
    bezeichnung = PruefeString(mvarBezeichnung)
End Property

Public Property Let bezeichnung(ByVal varBezeichnung As Variant)
    mvarBezeichnung = PruefeString(varBezeichnung)
End Property

Public Property Get LieferscheinLink() As Variant
    LieferscheinLink = PruefeString(mvarLieferscheinLink)
End Property

Public Property Let LieferscheinLink(ByVal varLieferscheinLink As Variant)
    mvarLieferscheinLink = PruefeLink(varLieferscheinLink)
End Property

Public Property Get seriennummer() As Variant
    seriennummer = PruefeString(mvarSeriennummer)
End Property

Public Property Let seriennummer(ByVal varSeriennummer As Variant)
    mvarSeriennummer = PruefeString(varSeriennummer)
End Property

Public Property Get anzahl() As Variant
    anzahl = PruefeInteger(mvarAnzahl)
End Property

Public Property Let anzahl(ByVal varAnzahl As Variant)
    mvarAnzahl = PruefeInteger(varAnzahl)
End Property

Public Property Get herstellerkennzeichen() As Variant
    herstellerkennzeichen = PruefeString(mvarHerstellerkennzeichen)
End Property

Public Property Let herstellerkennzeichen(ByVal varHerstellerkennzeichen As Variant)
    mvarHerstellerkennzeichen = PruefeString(varHerstellerkennzeichen)
End Property

Public Property Get UnterauftragnehmerAngebotLink() As Variant
    UnterauftragnehmerAngebotLink = PruefeLink(mvarUnterauftragnehmerAngebotLink)
End Property

Public Property Let UnterauftragnehmerAngebotLink(ByVal varUnterauftragnehmerAngebotLink As Variant)
    mvarUnterauftragnehmerAngebotLink = PruefeLink(varUnterauftragnehmerAngebotLink)
End Property

Public Property Get angebotNetto() As Variant
    angebotNetto = PruefeWaehrung(mvarAngebotNetto)
End Property

Public Property Let angebotNetto(ByVal varAngebotNetto As Variant)
    mvarAngebotNetto = PruefeWaehrung(varAngebotNetto)
End Property

Public Property Get preisBrutto() As Variant
    preisBrutto = PruefeWaehrung(mvarPreisBrutto)
End Property

Public Property Let preisBrutto(ByVal varPreisBrutto As Variant)
    mvarPreisBrutto = PruefeWaehrung(varPreisBrutto)
End Property

Public Property Get lieferdatum() As Variant
    lieferdatum = PruefeDatum(mvarLieferdatum)
End Property

Public Property Let lieferdatum(ByVal varLieferdatum As Variant)
    mvarLieferdatum = PruefeDatum(varLieferdatum)
End Property

Public Property Get lieferschein() As Variant
    lieferschein = PruefeString(mvarLieferschein)
End Property

Public Property Let lieferschein(ByVal varLieferschein As Variant)
    mvarLieferschein = PruefeString(varLieferschein)
End Property

Public Property Get Bemerkung() As Variant
    Bemerkung = PruefeString(mvarBemerkung)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    mvarBemerkung = PruefeString(varBemerkung)
End Property

Public Property Get zielAftrID() As Variant
    zielAftrID = PruefeString(mvarZielAftrID)
End Property

Public Property Let zielAftrID(ByVal varZielAftrID As Variant)
    mvarZielAftrID = PruefeString(varZielAftrID)
End Property

Public Property Get zielLieferschein() As Variant
    zielLieferschein = PruefeString(mvarZielLieferschein)
End Property

Public Property Let zielLieferschein(ByVal varZielLieferschein As Variant)
    mvarZielLieferschein = PruefeString(varZielLieferschein)
End Property

Public Property Get zielLinkLieferschein() As Variant
    zielLinkLieferschein = PruefeLink(mvarZielLinklieferschein)
End Property

Public Property Let zielLinkLieferschein(ByVal varZielLinkLieferschein As Variant)
    mvarZielLinklieferschein = PruefeLink(varZielLinkLieferschein)
End Property

Public Property Get zielLieferdatum() As Variant
    zielLieferdatum = PruefeDatum(mvarZielLieferdatum)
End Property

Public Property Let zielLieferdatum(ByVal varZielLieferdatum As Variant)
    mvarZielLieferdatum = PruefeDatum(varZielLieferdatum)
End Property

Public Property Get LiefergegenstandLink() As Variant
    LiefergegenstandLink = PruefeLink(mvarLiefergegenstandLink)
End Property

Public Property Let LiefergegenstandLink(ByVal varLiefergegenstandLink As Variant)
    mvarLiefergegenstandLink = PruefeLink(varLiefergegenstandLink)
End Property

Public Property Get IstReserve() As Variant
    IstReserve = PruefeBoolean(mvarIstReserve)
End Property

Public Property Let IstReserve(ByVal varIstReserve As Variant)
    mvarIstReserve = PruefeBoolean(varIstReserve)
End Property

Public Property Get AftrTitel() As Variant
    AftrTitel = PruefeString(mvarAftrTitel)
End Property

Public Property Get BWIKey() As Variant
    BWIKey = PruefeString(mvarBWIKey)
End Property

Public Property Get LeistungsbeschreibungLink() As Variant
    LeistungsbeschreibungLink = PruefeLink(mvarLeistungsbeschreibungLink)
End Property

Public Property Get ParentKey() As Variant
    ParentKey = PruefeString(mvarParentKey)
End Property




