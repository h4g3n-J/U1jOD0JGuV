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

' dient als Zwischenspeicher f�r die ID des aktuell
' mit dem Cursor ausgew�hlten Datensatz im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' steuert Debug-Meldungen
Private mbolVerbatim As Boolean

    Private mvarBWIKey As Variant
    Private EAkurzKey As Variant
    Private MengengeruestLink As Variant
    Private LeistungsbeschreibungLink As Variant
    Private Bemerkung As Variant
    Private BeauftragtDatum As Variant
    Private AbgebrochenDatum As Variant
    Private MitzeichnungI21Datum As Variant
    Private MitzeichnungI25Datum As Variant
    Private AngebotDatum As Variant
    Private AbgenommenDatum As Variant
    Private AftrBeginn As Variant
    Private AftrEnde As Variant
    Private StorniertDatum As Variant
    Private AngebotBrutto As Variant
    
Private Sub Default()
    mvarBWIKey = Null
    MengengeruestLink = Null
    LeistungsbeschreibungLink = Null
    Bemerkung = Null
    BeauftragtDatum = Null
    AbgebrochenDatum = Null
    MitzeichnungI21Datum = Null
    MitzeichnungI25Datum = Null
    AngebotDatum = Null
    AbgenommenDatum = Null
    AftrBeginn = Null
    AftrEnde = Null
    StorniertDatum = Null
    AngebotBrutto = Null
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

' das ist ein test