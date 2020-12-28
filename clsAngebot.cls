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

' Hauptdatensatz der Klasse
Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher für die ID des aktuell
' mit dem Cursor ausgewählten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' steuert Debug-Meldungen
Private mbolVerbatim As Boolean

    Private mvarBWIKey As Variant
    Private mvarEAkurzKey As Variant
    Private mvarMengengeruestLink As Variant
    Private mvarLeistungsbeschreibungLink As Variant
    Private mvarBemerkung As Variant
    Private mvarBeauftragtDatum As Variant
    Private mvarAbgebrochenDatum As Variant
    Private mvarMitzeichnungI21Datum As Variant
    Private mvarMitzeichnungI25Datum As Variant
    Private mvarAngebotDatum As Variant
    Private mvarAbgenommenDatum As Variant
    Private mvarAftrBeginn As Variant
    Private mvarAftrEnde As Variant
    Private mvarStorniertDatum As Variant
    Private mvarAngebotBrutto As Variant
    
Private Sub Default()
    mvarBWIKey = Null
    mvarMengengeruestLink = Null
    mvarLeistungsbeschreibungLink = Null
    mvarBemerkung = Null
    mvarBeauftragtDatum = Null
    mvarAbgebrochenDatum = Null
    mvarMitzeichnungI21Datum = Null
    mvarMitzeichnungI25Datum = Null
    mvarAngebotDatum = Null
    mvarAbgenommenDatum = Null
    mvarAftrBeginn = Null
    mvarAftrEnde = Null
    mvarStorniertDatum = Null
    mvarAngebotBrutto = Null
End Sub
    
Private Sub Class_Initialize()
    Debug.Print "Klasse clsAngebot instanziiert."
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsAngebot terminiert."
End Sub

Public Property Get BWIKey() As Variant
    BWIKey = basSupport.PruefeString(mvarBWIKey)
End Property

Public Property Let BWIKey(ByVal varBWIKey As Variant)
    mvarBWIKey = basSupport.PruefeString(varBWIKey)
End Property

Public Property Get MengengeruestLink() As Variant
    MengengeruestLink = basSupport.PruefeLink(mvarMengengeruestLink)
End Property

Public Property Let MengengeruestLink(ByVal varMengengeruestLink As Variant)
    mvarMengengeruestLink = basSupport.PruefeLink(varMengengeruestLink)
End Property

Public Property Get LeistungsbeschreibungLink() As Variant
    LeistungsbeschreibungLink = basSupport.PruefeLink(mvarLeistungsbeschreibungLink)
End Property

Public Property Let LeistungsbeschreibungLink(ByVal varLeistungsbeschreibungLink As Variant)
    mvarLeistungsbeschreibungLink = basSupport.PruefeLink(varLeistungsbeschreibungLink)
End Property

Public Property Get Bemerkung() As Variant
    Bemerkung = basSupport.PruefeString(mvarBemerkung)
End Property

Public Property Let Bemerkung(ByVal varBemerkung As Variant)
    mvarBemerkung = basSupport.PruefeString(varBemerkung)
End Property

Public Property Get BeauftragtDatum() As Variant
    BeauftragtDatum = basSupport.PruefeDatum(mvarBeauftragtDatum)
End Property

Public Property Let BeauftragtDatum(ByVal varBeauftragtDatum As Variant)
    mvarBeauftragtDatum = basSupport.PruefeDatum(varBeauftragtDatum)
End Property

Public Property Get AbgebrochenDatum() As Variant
    AbgebrochenDatum = basSupport.PruefeDatum(mvarAbgebrochenDatum)
End Property

Public Property Let AbgebrochenDatum(ByVal varAbgebrochenDatum As Variant)
    mvarAbgebrochenDatum = basSupport.PruefeDatum(varAbgebrochenDatum)
End Property

Public Property Get MitzeichnungI21Datum() As Variant
    MitzeichnungI21Datum = basSupport.PruefeDatum(mvarMitzeichnungI21Datum)
End Property

Public Property Let MitzeichnungI21Datum(ByVal varMitzeichnungI21Datum As Variant)
    mvarMitzeichnungI21Datum = basSupport.PruefeDatum(varMitzeichnungI21Datum)
End Property

Public Property Get MitzeichnungI25Datum() As Variant
    MitzeichnungI25Datum = basSupport.PruefeDatum(mvarMitzeichnungI25Datum)
End Property

Public Property Let MitzeichnungI25Datum(ByVal varMitzeichnungI25Datum As Variant)
    mvarMitzeichnungI25Datum = basSupport.PruefeDatum(varMitzeichnungI25Datum)
End Property

Public Property Get AngebotDatum() As Variant
    AngebotDatum = basSupport.PruefeDatum(mvarAngebotDatum)
End Property

Public Property Let AngebotDatum(ByVal varAngebotDatum As Variant)
    mvarAngebotDatum = basSupport.PruefeDatum(varAngebotDatum)
End Property

Public Property Get AbgenommenDatum() As Variant
    AbgenommenDatum = basSupport.PruefeDatum(mvarAbgenommenDatum)
End Property

Public Property Let AbgenommenDatum(ByVal varAbgenommenDatum As Variant)
    mvarAbgenommenDatum = basSupport.PruefeDatum(varAbgenommenDatum)
End Property

Public Property Get AftrBeginn() As Variant
    AftrBeginn = basSupport.PruefeDatum(mvarAftrBeginn)
End Property

Public Property Let AftrBeginn(ByVal varAftrBeginn As Variant)
    mvarAftrBeginn = basSupport.PruefeDatum(varAftrBeginn)
End Property

Public Property Get AftrEnde() As Variant
    AftrEnde = basSupport.PruefeDatum(mvarAftrEnde)
End Property

Public Property Let AftrEnde(ByVal varAftrEnde As Variant)
    mvarAftrEnde = basSupport.PruefeDatum(varAftrEnde)
End Property

Public Property Get StorniertDatum() As Variant
    StorniertDatum = basSupport.PruefeDatum(mvarStorniertDatum)
End Property

Public Property Let StorniertDatum(ByVal varStorniertDatum As Variant)
    mvarStorniertDatum = basSupport.PruefeDatum(varStorniertDatum)
End Property

Public Property Get AngebotBrutto() As Variant
    AngebotBrutto = basSupport.PruefeWaehrung(mvarAngebotBrutto)
End Property

Public Property Let AngebotBrutto(ByVal varAngebotBrutto As Variant)
    mvarAngebotBrutto = basSupport.PruefeWaehrung(varAngebotBrutto)
End Property

Public Property Let Verbatim(ByVal bolInput As Boolean)
    mbolVerbatim = basSupport.PruefeBoolean(bolInput)
End Property

Public Property Get Verbatim() As Boolean
    Verbatim = basSupport.PruefeBoolean(mbolVerbatim)
End Property

Public Sub AddRecordset()
    Dim dbsCurrentDb As DAO.Database
    Set dbsCurrentDb = CurrentDb
    
    ' name main table
    Dim strTableMain As String
    strTableMain = "tblAngebot"
    
    ' name assistance table
    Dim strTableAssistance As String
    strTableAssistance = "tblAuftragZuAngebot"
    
    ' create table array
    Dim astrTables(1) As String
    Dim varTable As Variant
        astrTables(0) = strTableMain
        astrTables(1) = strTableAssistance
            
    ' check variable
    Dim intStatus As Integer
    intStatus = 0
    
    ' check if tables exist
    For Each varTable In astrTables
        If basSupport.TabelleExistiert(varTable) = False Then
            Debug.Print "clsAngebot.AddRecordset: " & varTable _
                & " existiert nicht."
            intStatus = intStatus + 1
        Else:
            If mbolVerbatim = True Then
            Debug.Print "clsAngebot.AddRecordset: " + varTable _
                + " existiert"
            End If
        End If
    Next
    
    If intStatus > 0 Then
        Debug.Print "clsAngebot.AddRecodset: Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    ' create array for order and offer title
    Dim astrRecordset(3, 1) As String
        astrRecordset(0, 0) = "Order Title"
        astrRecordset(1, 0) = "Titel des referenzierten Auftrags eingeben"
        astrRecordset(2, 0) = "Angebot erfassen"
        astrRecordset(3, 0) = ""
        
        astrRecordset(0, 1) = "Offer Title"
        astrRecordset(1, 1) = "BWI Alias des Angebots eingeben"
        astrRecordset(2, 1) = "Angebot erfassen"
        astrRecordset(3, 1) = ""
    
    ' create increment
    Dim lngi As Long
    
    'ask for order title and offer title
    For lngi = LBound(astrRecordset, 2) To UBound(astrRecordset, 2)
        astrRecordset(3, lngi) = InputBox(astrRecordset(1, lngi), astrRecordset(2, lngi))
        
            ' check if inputbox is empty, if so prompt message and exit procedure
            If basSupport.PflichtfeldIstLeer(astrRecordset(3, lngi)) = True Then
                Debug.Print "clsAngebot.AddRecordset: " & astrRecordset(0, lngi) & " ist leer, " _
                    & "Prozedur abgebrochen"
                GoTo ExitProc
            End If
        
            ' prompt input
            If mbolVerbatim = True Then
                Debug.Print "clsAngebot.AddRecordset: " & astrRecordset(0, lngi) _
                    & " = " & astrRecordset(3, lngi)
            End If
    Next
        
    ' create the new recordset in the assistance table
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(strTableAssistance, dbOpenDynaset)
    
        mrstRecordset.AddNew
            mrstRecordset!AftrID = astrRecordset(3, 0)
            mrstRecordset!BWIKey = astrRecordset(3, 1)
        mrstRecordset.Update
        
        ' close mrstRecordset
        mrstRecordset.Close
        Set mrstRecordset = Nothing
        
    ' create the new offer recordset in the main Tabelle
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(strTableMain, dbOpenDynaset)
    
        mrstRecordset.AddNew
            mrstRecordset!BWIKey = astrRecordset(3, 1)
        mrstRecordset.Update
        
        ' copy offer name to clipboard
        mvarClipboard = astrRecordset(3, 1)
        
        If mbolVerbatim = True Then
            Debug.Print "clsAngebot.AddRecordset: clsAngebot!mvarClipboard = " + mvarClipboard
        End If
    
    MsgBox "'" & mvarClipboard & "' erzeugt", vbOKOnly, "Angebot hinzufuegen"
    
    ' clean up
    mrstRecordset.Close
    Set mrstRecordset = Nothing
    
ExitProc:
    dbsCurrentDb.Close
    Set dbsCurrentDb = Nothing
End Sub
