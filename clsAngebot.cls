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

' dient als Zwischenspeicher für die ID des aktuell
' mit dem Cursor ausgewählten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant
Private mvarRecordsetClipboard As Variant
    
Private Function AngebotDefaultProperties()

    If gconVerbatim = True Then
        Debug.Print "clsAngebot.AngebotDefaultProperties ausfuehren"
    End If
    
    Dim varAngebotSettings(16, 2) As Variant
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
    varAngebotSettings(5, 0) = "Verfuegung"
        varAngebotSettings(5, 1) = "link"
        varAngebotSettings(5, 2) = Null
    varAngebotSettings(6, 0) = "Bemerkung"
        varAngebotSettings(6, 1) = "string"
        varAngebotSettings(6, 2) = Null
    varAngebotSettings(7, 0) = "BeauftragtDatum"
        varAngebotSettings(7, 1) = "date"
        varAngebotSettings(7, 2) = Null
    varAngebotSettings(8, 0) = "AbgebrochenDatum"
        varAngebotSettings(8, 1) = "date"
        varAngebotSettings(8, 2) = Null
    varAngebotSettings(9, 0) = "MitzeichnungI21Datum"
        varAngebotSettings(9, 1) = "date"
        varAngebotSettings(9, 2) = Null
    varAngebotSettings(10, 0) = "MitzeichnungI25Datum"
        varAngebotSettings(10, 1) = "date"
        varAngebotSettings(10, 2) = Null
    varAngebotSettings(11, 0) = "AngebotDatum"
        varAngebotSettings(11, 1) = "date"
        varAngebotSettings(11, 2) = Null
    varAngebotSettings(12, 0) = "AbgenommenDatum"
        varAngebotSettings(12, 1) = "date"
        varAngebotSettings(12, 2) = Null
    varAngebotSettings(13, 0) = "AftrBeginn"
        varAngebotSettings(13, 1) = "date"
        varAngebotSettings(13, 2) = Null
    varAngebotSettings(14, 0) = "AftrEnde"
        varAngebotSettings(14, 1) = "date"
        varAngebotSettings(14, 2) = Null
    varAngebotSettings(15, 0) = "StorniertDatum"
        varAngebotSettings(15, 1) = "date"
        varAngebotSettings(15, 2) = Null
    varAngebotSettings(16, 0) = "AngebotBrutto"
        varAngebotSettings(16, 1) = "currency"
        varAngebotSettings(16, 2) = Null
    
    AngebotDefaultProperties = varAngebotSettings
End Function
    
Private Sub Class_Initialize()
    If gconVerbatim Then
        Debug.Print "Klasse clsAngebot instanziiert."
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = AngebotDefaultProperties
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "BWIKey")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform corresponding to data type [mvarRecordsetClipboard(intAttributPosition, 1)]
    BWIKey = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let BWIKey(ByVal varBWIKey As Variant)
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.BWIKey (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "BWIKey", 0, "B")
    
    ' transform input [mavarColumnToVariable(1, intItemPosition)] then assign to attribut value [mavarColumnToVariable(2, intItemPosition)]
    mvarRecordsetClipboard(2, intItemPosition) = basSupport.CheckDataType(varBWIKey, mvarRecordsetClipboard(1, intItemPosition))
    
    mvarBWIKey = basSupport.PruefeString(varBWIKey)
End Property

Public Property Get EAkurzKey() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.EAkurzKey (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "EakurzKey")
    
    ' get value [mavarColumnToVariable(2, intItemPosition)], transform data type corresponding to [mavarColumnToVariable(1, intItemPosition)]
    EAkurzKey = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let EAkurzKey(ByVal varEAkurzKey As Variant)
    
    'verbatim message
    If gconVerbatim = True Then
        Debug.Print "clsAuftrag.Kunde (Let) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intItemPosition As Integer
    intItemPosition = basSupport.FindItemInArray(mavarColumnToVariable, "Kunde", 0, "B")
    
    ' transform input [mavarColumnToVariable(1, intItemPosition)], assign to attribut value [mavarColumnToVariable(2, intItemPosition)]
    mavarColumnToVariable(2, intItemPosition) = basSupport.CheckDataType(varKunde, mavarColumnToVariable(1, intItemPosition))
    
End Property

Public Property Get MengengeruestLink() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.MengengeruestLink (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "MengengeruestLink")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    MengengeruestLink = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "LeistungsbeschreibungLink")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    LeistungsbeschreibungLink = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let LeistungsbeschreibungLink(ByVal varLeistungsbeschreibungLink As Variant)
    mvarLeistungsbeschreibungLink = basSupport.PruefeLink(varLeistungsbeschreibungLink)
End Property

Public Property Get Verfuegung() As Variant

    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.Verfuegung (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "Verfuegung")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    Verfuegung = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let Verfuegung(ByVal varVerfuegung As Variant)

End Property

Public Property Get Bemerkung() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.Bemerkung (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "Bemerkung")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    Bemerkung = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "BeauftragtDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    BeauftragtDatum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AbgebrochenDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AbgebrochenDatum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AbgebrochenDatum(ByVal varAbgebrochenDatum As Variant)
    
End Property

Public Property Get MitzeichnungI21Datum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.MitzeichnungI21Datum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "MitzeichnungI21Datum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    MitzeichnungI21Datum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let MitzeichnungI21Datum(ByVal varMitzeichnungI21Datum As Variant)
    
End Property

Public Property Get MitzeichnungI25Datum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.MitzeichnungI25Datum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "MitzeichnungI25Datum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    MitzeichnungI25Datum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let MitzeichnungI25Datum(ByVal varMitzeichnungI25Datum As Variant)
    
End Property

Public Property Get AngebotDatum() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "clsAngebot.AngebotDatum (Get) ausfuehren"
    End If
    
    ' scan attribut array for attribut name
    Dim intAttributPosition As Integer
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AngebotDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AngebotDatum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AbgenommenDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AbgenommenDatum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AftrBeginn")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AftrBeginn = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AftrEnde")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AftrEnde = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "StorniertDatum")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    StorniertDatum = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
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
    intAttributPosition = basSupport.FindItemInArray(mvarRecordsetClipboard, "AngebotBrutto")
    
    ' get value [mvarRecordsetClipboard(intAttributPosition, 2)], transform data type corresponding [mvarRecordsetClipboard(intAttributPosition, 1)]
    AngebotBrutto = basSupport.CheckDataType(mvarRecordsetClipboard(intAttributPosition, 2), mvarRecordsetClipboard(intAttributPosition, 1))
    
End Property

Public Property Let AngebotBrutto(ByVal varAngebotBrutto As Variant)
    mvarAngebotBrutto = basSupport.PruefeWaehrung(varAngebotBrutto)
End Property

Public Sub AddRecordset()
    
    mvarClipboard = basSupport.AddRecordsetMN("tblAuftrag", "AftrID", "Auftrag", _
        "Titel des referenzierten Auftrags eingeben", "Auftrag zuordnen", _
        "tblAngebot", "BWIKey", "Angebot Titel", "Titel des Angebots eingeben", _
        "Angebot erfassen", "tblAuftragZuAngebot", False)
    
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
        mvarRecordsetClipboard = AngebotDefaultProperties
        GoTo ExitProc
    End If
        
    ' check if target table exists, when negative _
    ' return error + exitProc
    If basSupport.ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsAngebot.SelectRecordset: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    End If
    
    ' check if recordset exists
    ' likely obsolete, since the recordset is chosen from a recordset selection
    If basSupport.RecordsetExists(strTableName, strKeyColumn, varRecordsetName) <> True Then
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
