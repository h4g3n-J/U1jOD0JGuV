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
    Debug.Print "Klasse clsAngebotgNew terminiert."
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
    
    ' name relevant tables
    
    Dim strTabelleTitel As String
    strTabelleTitel = "tblAngebot"
    
    Dim strHilfstabelleTitel As String
    strHilfstabelleTitel = "tblAuftragZuAngebot"
    
    ' check if tables exist, if not prompt error
    If basSupport.TabelleExistiert(strTabelleTitel) = False Then
        Debug.Print "clsAngebot.AddRecordset: " & strTabelleTitel _
            & " existiert nicht."
    End If
    
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot.AddRecordset: " + strTabelleTitel _
            + " existiert"
    End If
    
    If basSupport.TabelleExistiert(strHilfstabelleTitel) = False Then
        Debug.Print "clsAngebot.AddRecordset: " & strHilfstabelleTitel _
            & " existiert nicht."
    End If
    
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot.AddRecordset: " + strHilfstabelleTitel _
            + " existiert"
    End If
    
    If basSupport.TabelleExistiert(strTabelleTitel) = False Or basSupport.TabelleExistiert(strHilfstabelleTitel) = False Then
        Debug.Print "clsAngebot.AddRecodset: Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    ' name the referenced order
    Dim varOrderTitle As Variant
    varOrderTitle = InputBox("Titel des referenzierten Auftrags eingeben", "Angebot erfassen")
    
    If basSupport.PflichtfeldIstLeer(varOrderTitle) = True Then
        Debug.Print "clsAngebot.AddRecordset: varOrderTitle ist leer, " _
            & "Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    varOrderTitle = CStr(varOrderTitle)
    
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot.AddRecordset!varOrderTitle: " + varOrderTitle
    End If
    
    ' name the offer
    Dim varOfferTitle As Variant
    varOfferTitle = InputBox("BWI Alias des Angebots eingeben.", "Angebot erfassen")
    
    If basSupport.PflichtfeldIstLeer(varOfferTitle) = True Then
        Debug.Print "clsAngebot.AddRecordset: varOfferTitle ist leer, " _
            & "Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    varOfferTitle = CStr(varOfferTitle)
    
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot.AddRecordset!varOfferTitle: " + varOfferTitle
    End If
    
    ' Create RecordSet in Hilfstabelle
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(strHilfstabelleTitel, dbOpenDynaset)
    
    mrstRecordset.AddNew
        mrstRecordset!AftrID = varOrderTitle
        mrstRecordset!BWIKey = varOfferTitle
    mrstRecordset.Update
    
    mrstRecordset.Close
    Set mrstRecordset = Nothing
    
    ' Create Recordset in main Tabelle
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(strTabelleTitel, dbOpenDynaset)
    
    mrstRecordset.AddNew
        mrstRecordset!BWIKey = varOfferTitle
    mrstRecordset.Update
    
    ' copy recordset title to clipboard
    mvarClipboard = varOfferTitle
    
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot.AddRecordset: clsAngebot!mvarClipboard = " + mvarClipboard
    End If
    
    ' Abschlussmeldung
    MsgBox "'" & mvarClipboard & "' erzeugt", vbOKOnly, "Angebot hinzufuegen"
    
ExitProc:
    mrstRecordset.Close
    Set mrstRecordset = Nothing
    dbsCurrentDb.Close
    Set dbsCurrentDb = Nothing
End Sub
