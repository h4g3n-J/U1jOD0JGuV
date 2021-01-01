VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLeistungserfassungungsblatt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsLeistungserfassungsblatt

Option Compare Database
Option Explicit

Private mvarClipboard As Variant

' verbatim setting
Private mbolverbatim As Boolean

    Private mvarLeistungserfassungsblatt As Variant
    Private mvarRechnungNr As Variant
    Private mvarBemerkung As Variant
    Private mvarBelegID As Variant
    Private mvarBrutto As Variant

Private Sub default()
    mvarLeistungserfassungsblatt = Null
    mvarRechnungNr = Null
    mvarBemerkung = Null
    mvarBelegID = Null
    mvarBrutto = Null
End Sub

Private Sub Class_Initialize()
    Debug.Print "Klasse clsLeistungserfassungsblatt instanziiert."
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsLeistungserfassungsblatt terminiert."
End Sub

Public Property Get Leistungserfassungsblatt() As Variant
    Leistungserfassungsblatt = basSupport.PruefeString(mvarLeistungserfassungsblatt)
End Property

Public Property Let Leistungserfassungsblatt(ByVal varLeistungserfassungsblatt As Variant)
    mvarLeistungserfassungsblatt = basSupport.PruefeString(varLeistungserfassungsblatt)
End Property

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

Public Property Get BelegID() As Variant
    BelegID = basSupport.PruefeString(mvarBelegID)
End Property

Public Property Let BelegID(ByVal varBelegID As Variant)
    mvarBelegID = basSupport.PruefeString(varBelegID)
End Property

Public Property Get Brutto() As Variant
    Brutto = basSupport.PruefeString(mvarBrutto)
End Property

Public Property Let Brutto(ByVal varBrutto As Variant)
    mvarBrutto = basSupport.PruefeString(varBrutto)
End Property

Public Property Get Verbatim() As Variant
    Verbatim = basSupport.PruefeString(mbolverbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolverbatim = basSupport.PruefeString(varVerbatim)
End Property

Public Sub AddRecordset()
    mvarClipboard = basSupport.AddRecordsetMN("tblRechnung", "RechnungNr", "Rechnungsnummer", "Referenzierte Rechnungsnummer eingeben", "Rechnung erfassen", _
    "tblLeistungserfassungsblatt", "Leistungserfassungsblatt", "Leistungserfassungsblatt", "Leistungserfassungsblattnummer eingeben", "Leistungserfassungsblatt erfassen", _
    "tblRechnungZuLeistungserfassung", _
    False)
End Sub

