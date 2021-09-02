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

'Hauptdatensatz der Klasse
Private mrstRecordset As DAO.RecordSet

' dient als Zwischenspeicher fuer die ID des
' mit dem Cursor ausgewaehlten Datensatzes im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean

Private Sub DefaultProperties()

    mbolVerbatim = False
    
    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsLeistungserfassungsblatt.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 5
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "Leistungserfassungsblatt"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "RechnungNr"
        varDefaultProperties(2, 1) = "string"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "Bemerkung"
        varDefaultProperties(3, 1) = "string"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "BelegID"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "Brutto"
        varDefaultProperties(5, 1) = "currency"
        varDefaultProperties(5, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)
        
    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsLeistungserfassungsblatt.DefaultProperties executed"
    End If

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
    Verbatim = basSupport.PruefeString(mbolVerbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolVerbatim = basSupport.PruefeString(varVerbatim)
End Property

Public Sub AddRecordset()
    mvarClipboard = basSupport.AddRecordsetMN("tblRechnung", "RechnungNr", "Rechnungsnummer", "Referenzierte Rechnungsnummer eingeben", "Rechnung erfassen", _
    "tblLeistungserfassungsblatt", "Leistungserfassungsblatt", "Leistungserfassungsblatt", "Leistungserfassungsblattnummer eingeben", "Leistungserfassungsblatt erfassen", _
    "tblRechnungZuLeistungserfassung", _
    False)
End Sub

