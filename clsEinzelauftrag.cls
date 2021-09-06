VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEinzelauftrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsEinzelauftrag

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

Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "clsEinzelauftrag initialized"
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    Debug.Print "clsEinzelauftrag terminated"
End Sub

Private Function DefaultProperties()

    mbolVerbatim = False

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsEinzelauftrag.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 15
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "EAkurzKey"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "MengengeruestLink"
        varDefaultProperties(2, 1) = "link"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "LeistungsbeschreibungLink"
        varDefaultProperties(3, 1) = "link"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "Bemerkung"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "BeauftragtDatum"
        varDefaultProperties(5, 1) = "date"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "AbgebrochenDatum"
        varDefaultProperties(6, 1) = "date"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "MitzeichnungI21Datum"
        varDefaultProperties(7, 1) = "date"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "MitzeichnungI25Datum"
        varDefaultProperties(8, 1) = "date"
        varDefaultProperties(8, 2) = Null
    varDefaultProperties(9, 0) = "AngebotDatum"
        varDefaultProperties(9, 1) = "date"
        varDefaultProperties(9, 2) = Null
    varDefaultProperties(10, 0) = "AbgenommenDatum"
        varDefaultProperties(10, 1) = "date"
        varDefaultProperties(10, 2) = Null
    varDefaultProperties(11, 0) = "StorniertDatum"
        varDefaultProperties(11, 1) = "date"
        varDefaultProperties(11, 2) = Null
    varDefaultProperties(12, 0) = "AngebotBrutto"
        varDefaultProperties(12, 1) = "currency"
        varDefaultProperties(12, 2) = Null
    varDefaultProperties(13, 0) = "BWIKey"
        varDefaultProperties(13, 1) = "string"
        varDefaultProperties(13, 2) = Null
    varDefaultProperties(14, 0) = "AftrBeginn"
        varDefaultProperties(14, 1) = "currency"
        varDefaultProperties(14, 2) = Null
    varDefaultProperties(15, 0) = "AftrEnde"
        varDefaultProperties(15, 1) = "string"
        varDefaultProperties(15, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)

    DefaultProperties = varDefaultProperties
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsEinzelauftrag.DefaultProperties executed"
    End If
    
End Function


