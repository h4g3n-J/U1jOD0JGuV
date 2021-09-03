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




