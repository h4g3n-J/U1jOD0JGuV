VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAngebot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsAngebot

Option Compare Database
Option Explicit

' dient als Zwischenspeicher für die ID des
' mit dem Cursor ausgewählten Datensatzes im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant
Private mvarRecordsetClipboard As Variant
    
Private Function DefaultProperties()

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.DefaultProperties"
    End If
    
    Dim varAngebotSettings(13, 2) As Variant
    varAngebotSettings(0, 0) = "Attribut Name"
        varAngebotSettings(0, 1) = "data type"
        varAngebotSettings(0, 2) = "default value"
    varAngebotSettings(1, 0) = "BWIKey"
        varAngebotSettings(1, 1) = "string"
        varAngebotSettings(1, 2) = Null
    varAngebotSettings(2, 0) = "EAkurzKey"
        varAngebotSettings(2, 1) = "string"
        varAngebotSettings(2, 2) = Null
    varAngebotSettings(3, 0) = "MengengeruestLink"
        varAngebotSettings(3, 1) = "link"
        varAngebotSettings(3, 2) = Null
    varAngebotSettings(4, 0) = "LeistungsbeschreibungLink"
        varAngebotSettings(4, 1) = "link"
        varAngebotSettings(4, 2) = Null
    varAngebotSettings(5, 0) = "Bemerkung"
        varAngebotSettings(5, 1) = "string"
        varAngebotSettings(5, 2) = Null
    varAngebotSettings(6, 0) = "BeauftragtDatum"
        varAngebotSettings(6, 1) = "date"
        varAngebotSettings(6, 2) = Null
    varAngebotSettings(7, 0) = "AbgebrochenDatum"
        varAngebotSettings(7, 1) = "date"
        varAngebotSettings(7, 2) = Null
    varAngebotSettings(8, 0) = "AngebotDatum"
        varAngebotSettings(8, 1) = "date"
        varAngebotSettings(8, 2) = Null
    varAngebotSettings(9, 0) = "AbgenommenDatum"
        varAngebotSettings(9, 1) = "date"
        varAngebotSettings(9, 2) = Null
    varAngebotSettings(10, 0) = "AftrBeginn"
        varAngebotSettings(10, 1) = "date"
        varAngebotSettings(10, 2) = Null
    varAngebotSettings(11, 0) = "AftrEnde"
        varAngebotSettings(11, 1) = "date"
        varAngebotSettings(11, 2) = Null
    varAngebotSettings(12, 0) = "StorniertDatum"
        varAngebotSettings(12, 1) = "date"
        varAngebotSettings(12, 2) = Null
    varAngebotSettings(13, 0) = "AngebotBrutto"
        varAngebotSettings(13, 1) = "currency"
        varAngebotSettings(13, 2) = Null

    DefaultProperties = varAngebotSettings
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.DefaultProperties executed"
    End If
    
End Function
    
Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "Klasse clsAngebot instanziiert."
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    If gconVerbatim Then
        Debug.Print "Klasse clsAngebot terminiert."
    End If
End Sub

Public Property Get BWIKey() As Variant

    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.BWIKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    ' intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "BWIKey")
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "BWIKey")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intAttributPosition, 1)]
    BWIKey = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let BWIKey(ByVal varBWIKey As Variant)
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.BWIKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, "BWIKey", 0, "B")
    
    ' transform input [mavarColumnToVariable(1, intItemPosition)] then assign to attribut value [mavarColumnToVariable(2, intItemPosition)]
    mvarRecordsetClipboard(2, intItemPosition) = CheckDataType(varBWIKey, mvarRecordsetClipboard(1, intItemPosition))
    
    mvarBWIKey = basSupport.PruefeString(varBWIKey)
End Property

Public Property Get EAkurzKey() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.EAkurzKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "EakurzKey")
    
    ' get value [mavarColumnToVariable(2, intItemPosition)], transform data type corresponding to [mavarColumnToVariable(1, intItemPosition)]
    EAkurzKey = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let EAkurzKey(ByVal varEAkurzKey As Variant)
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.Kunde (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mavarColumnToVariable, "Kunde", 0, "B")
    
    ' transform input [mavarColumnToVariable(1, intItemPosition)], assign to attribut value [mavarColumnToVariable(2, intItemPosition)]
    mavarColumnToVariable(2, intItemPosition) = CheckDataType(varKunde, mavarColumnToVariable(1, intItemPosition))
    
End Property

Public Property Get MengengeruestLink() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.MengengeruestLink (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "MengengeruestLink")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    MengengeruestLink = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let MengengeruestLink(ByVal varMengengeruestLink As Variant)
    mvarMengengeruestLink = basSupport.PruefeLink(varMengengeruestLink)
End Property

Public Property Get LeistungsbeschreibungLink() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.LeistungsbeschreibungLink (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "LeistungsbeschreibungLink")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    LeistungsbeschreibungLink = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let LeistungsbeschreibungLink(ByVal varLeistungsbeschreibungLink As Variant)
    mvarLeistungsbeschreibungLink = basSupport.PruefeLink(varLeistungsbeschreibungLink)
End Property

Public Property Get Bemerkung() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.Bemerkung (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "Bemerkung")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    
End Property

Public Property Get BeauftragtDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.BeauftragtDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "BeauftragtDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    BeauftragtDatum = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let BeauftragtDatum(ByVal varBeauftragtDatum As Variant)
    
End Property

Public Property Get AbgebrochenDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AbgebrochenDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AbgebrochenDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AbgebrochenDatum = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AbgebrochenDatum(ByVal varAbgebrochenDatum As Variant)
    
End Property

Public Property Get AngebotDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AngebotDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AngebotDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AngebotDatum = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AngebotDatum(ByVal varAngebotDatum As Variant)
    mvarAngebotDatum = basSupport.PruefeDatum(varAngebotDatum)
End Property

Public Property Get AbgenommenDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AbgenommenDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AbgenommenDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AbgenommenDatum = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AbgenommenDatum(ByVal varAbgenommenDatum As Variant)
    mvarAbgenommenDatum = basSupport.PruefeDatum(varAbgenommenDatum)
End Property

Public Property Get AftrBeginn() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AftrBeginn (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AftrBeginn")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AftrBeginn = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AftrBeginn(ByVal varAftrBeginn As Variant)
    mvarAftrBeginn = basSupport.PruefeDatum(varAftrBeginn)
End Property

Public Property Get AftrEnde() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AftrEnde (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AftrEnde")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AftrEnde = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AftrEnde(ByVal varAftrEnde As Variant)
    mvarAftrEnde = basSupport.PruefeDatum(varAftrEnde)
End Property

Public Property Get StorniertDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.StorniertDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "StorniertDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    StorniertDatum = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let StorniertDatum(ByVal varStorniertDatum As Variant)
    mvarStorniertDatum = basSupport.PruefeDatum(varStorniertDatum)
End Property

Public Property Get AngebotBrutto() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AngebotBrutto (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = FindItemInArray(mvarRecordsetClipboard, "AngebotBrutto")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AngebotBrutto = CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AngebotBrutto(ByVal varAngebotBrutto As Variant)
    mvarAngebotBrutto = basSupport.PruefeWaehrung(varAngebotBrutto)
End Property

Public Sub AddRecordset()
    
    ' mvarClipboard = basSupport.AddRecordsetMN("tblAuftrag", "AftrID", "Auftrag", _
        "Titel des referenzierten Auftrags eingeben", "Auftrag zuordnen", _
        "tblAngebot", "BWIKey", "Angebot Titel", "Titel des Angebots eingeben", _
        "Angebot erfassen", "tblAuftragZuAngebot", False)
        
    MsgBox ("clsAngebot.AddRecordset Code überprüfen.")
    
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
    strTableName = "tblAngebot"

    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"

    ' name of the column that contains the recordset
    Dim strKeyColumn As String
    strKeyColumn = "BWIKey"

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
            Debug.Print "clsAngebot.SelectRecordset: varRecordsetName ist leer, " _
                & "Defaultwert anzeigen"
        End If
        ' default Value
        mvarRecordsetClipboard = DefaultProperties
        GoTo ExitProc
    End If
        
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsAngebot.SelectRecordset: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    End If
    
    ' check if recordset exists
    ' likely obsolete, since the recordset is chosen from a recordset selection
    If RecordsetExists(strTableName, strKeyColumn, varRecordsetName) <> True Then
        Debug.Print "clsAngebot.SelectRecordset: " & varRecordsetName & " existiert nicht. Prozedur abgebrochen."
        MsgBox varRecordsetName & " existiert nicht. Prozedur abgebrochen.", vbCritical, "Datensatz nicht gefunden"
        GoTo ExitProc
    End If

    ' select recordset from table or query
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = '" & CStr(varRecordsetName) & "'")

    ' assign recordset to properties
    ' mvarRecordsetClipboard(inti, 2): propertie value
    ' mvarRecordsetClipboard(inti, 0): propertie name
    Dim inti As Integer
    For inti = LBound(mvarRecordsetClipboard, 1) + 1 To UBound(mvarRecordsetClipboard, 1)
        mvarRecordsetClipboard(inti, 2) = rstRecordset.Fields(mvarRecordsetClipboard(inti, 0))
    Next
    
    ' verbatim message: return data content from array
    ' mvarRecordsetClipboard(inti, 2): propertie value
    ' mvarRecordsetClipboard(inti, 0): propertie name
    If bolLocalVerbatim = True Then
        For inti = LBound(mvarRecordsetClipboard, 1) To UBound(mvarRecordsetClipboard, 1)
            Debug.Print "clsAngebot.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
End Sub

' creates a new recordset in a m:n relationship
' checks table existence, input, reference recordset
' returns name of the new recordset
Private Function AddRecordsetMN(ByVal strTableAName, strTableAKeyColumn, strTableAArtifact, strTableAInputDialogMessage, _
    strTableAInputDialogTitle, strTableBName, strTableBKeyColumn, strTableBArtifact, strTableBInputDialogMessage, _
        strTableBInputDialogTitle, strTableAssistanceName As String, gconVerbatim As Boolean) As String
        
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.AddRecordsetMN ausfuehren"
    End If
    
    Dim astrConfig(6, 2) As String
    
        ' table A - the table on the m-side
        ' table name
            astrConfig(0, 0) = strTableAName
        ' key column name
            astrConfig(1, 0) = strTableAKeyColumn
        ' artifact name
            astrConfig(2, 0) = strTableAArtifact
        ' input dialog message
            astrConfig(3, 0) = strTableAInputDialogMessage
        ' input dialog title
            astrConfig(4, 0) = strTableAInputDialogTitle
        ' target state for RecordExists - supposed to be true
            astrConfig(5, 0) = CStr(True)
        ' recordset name - supposed to be empty
            astrConfig(6, 0) = ""
    
        ' table B - the table on the n-side
        ' table name
            astrConfig(0, 1) = strTableBName
        ' key column name
            astrConfig(1, 1) = strTableBKeyColumn
        ' artifact name
            astrConfig(2, 1) = strTableBArtifact
        ' input dialog message
            astrConfig(3, 1) = strTableBInputDialogMessage
        ' input dialog title
            astrConfig(4, 1) = strTableBInputDialogTitle
        ' target state for RecordExists - supposed to be false
            astrConfig(5, 1) = CStr(False)
        ' recordset name
            astrConfig(6, 1) = ""
    
        ' assistance table - the m:n-table
        ' table name
            astrConfig(0, 2) = strTableAssistanceName
            
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    
    Dim inti As Integer
    Dim intj As Integer
    
    If gconVerbatim = True Then
        For inti = LBound(astrConfig, 2) To UBound(astrConfig, 2)
            For intj = LBound(astrConfig, 1) To UBound(astrConfig, 1)
                Debug.Print "astrConfig(" & intj & ", " & inti & ") = " & astrConfig(intj, inti)
            Next
        Next
    End If

    ' check if tables exist
    Dim intError As Integer
    intError = 0
    
    For inti = LBound(astrConfig, 2) To UBound(astrConfig, 2)
        If ObjectExists(astrConfig(0, inti), "table", False) = False Then
            Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(0, inti) _
                & " existiert nicht."
            intError = intError + 1
        Else:
            If gconVerbatim = True Then
            Debug.Print "basSupport.AddRecordsetMN: " + astrConfig(0, inti) _
                + " existiert"
            End If
        End If
    Next
    
    If intError > 0 Then
        Debug.Print "basSupport.AddRecodset: Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    'ask for the names of the recordsets
    Dim lngi As Long
    For lngi = LBound(astrConfig, 2) To UBound(astrConfig, 2) - 1
        astrConfig(6, lngi) = InputBox(astrConfig(3, lngi), astrConfig(4, lngi))
        
            ' check if inputbox is empty, if true then messagebox + exit procedure
            If basSupport.PflichtfeldIstLeer(astrConfig(6, lngi)) = True Then
                Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(2, lngi) & " ist leer, " _
                    & "Prozedur abgebrochen"
                MsgBox astrConfig(2, lngi) & " ist leer. Prozedur wird abgebrochen.", vbCritical, "Fehler"
                GoTo ExitProc
            End If
        
            ' return input
            If gconVerbatim = True Then
                Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(2, lngi) _
                    & " = " & astrConfig(6, lngi)
            End If
            
            ' check if recordset exists
            If CStr(RecordsetExists(astrConfig(0, lngi), astrConfig(1, lngi), _
                astrConfig(6, lngi))) <> astrConfig(5, lngi) Then
                ' error message: messagebox + exit procedure
                    If astrConfig(5, lngi) = CStr(False) Then
                        Debug.Print "basSupport.AddRecordsetMN: '" & astrConfig(6, lngi) _
                            & "' existiert bereits. Prozedur abgebrochen."
                        MsgBox "'" & astrConfig(6, lngi) & "' existiert bereits.", _
                            vbCritical, "Doppelter Eintrag"
                    Else:
                        Debug.Print "basSupport.AddRecordsetMN: '" & astrConfig(6, lngi) _
                            & "' existiert nicht. Prozedur abgebrochen."
                        MsgBox "'" & astrConfig(6, lngi) & "' existiert nicht.", _
                            vbCritical, "Referenzdantensatz fehlt"
                    End If
                GoTo ExitProc
            End If
    Next
        
    ' create recordset in assistance table
    Set rstRecordset = dbsCurrentDB.OpenRecordset(astrConfig(0, 2), dbOpenDynaset)
    
        rstRecordset.AddNew
            rstRecordset.Fields(astrConfig(1, 0)) = astrConfig(6, 0)
            rstRecordset.Fields(astrConfig(1, 1)) = astrConfig(6, 1)
        rstRecordset.Update
        
        ' close rstRecordset
        rstRecordset.Close
        Set rstRecordset = Nothing
        
    ' create recordset in table B
    Set rstRecordset = dbsCurrentDB.OpenRecordset(astrConfig(0, 1), dbOpenDynaset)
    
        rstRecordset.AddNew
            rstRecordset.Fields(astrConfig(1, 1)) = astrConfig(6, 1)
        rstRecordset.Update
        
    ' confirmation message
    MsgBox astrConfig(6, 1) & " erzeugt.", vbOKOnly, "Datensatz erstellen"

        ' return table B recordset name
        AddRecordsetMN = astrConfig(6, 1)
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
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
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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
        Debug.Print "execute clsAngebot.RecordsetExists"
    End If

    Dim bolStatus As Boolean
    bolStatus = False
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.RecordsetExists: varTblName = " & varTblName & " , varFieldName = " & " , varRecordsetName = " & varRecordsetName
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
        Debug.Print "clsAngebot.RecordsetExists executed"
    End If
    
End Function

' search varWanted in a two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.FindItemInArray: varWanted = " & varWanted
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
        Debug.Print "clsAngebot.FindItemInArray executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.CheckDataType"
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
            Debug.Print "clsAngebot.CheckDataType: varInput ist null"
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
