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

Private Sub Default()
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
       
mvarClipboard = basSupport.AddRecordsetMN("tblAngebot", "BWIKey", "Angebot Titel", _
    "Titel des referenzierten Angebots eingeben", "Angebot erfassen", "tblRechnung", _
    "RechnungNr", "Rechnungsnummer", "Rechnungsnummer eingeben", "Rechnung erfassen", _
    "tblAngebotZuRechnung", False)
    
End Sub

