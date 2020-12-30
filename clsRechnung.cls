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

    Private mvarRechnungNr
    Private mvarBemerkung
    Private mvarRechnungLink
    Private mvarTechnischRichtigDatum
    Private mvarIstTeilrechnung
    Private mvarIstSchlussrechnung
    Private mvarKalkulationLNWLink
    Private mvarRechnungBrutto

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
    Dim dbsCurrentDb As DAO.Database
    Set dbsCurrentDb = CurrentDb

    Dim inti As Integer
        
    ' create table array
    Dim astrRecordset(6, 2) As String
        ' table A - the table on the m-side
        ' table name
        astrRecordset(0, 0) = "tblAngebot"
        ' key column name
        astrRecordset(1, 0) = "BWIKey"
        ' title
        astrRecordset(2, 0) = "Angebot Titel"
        ' input dialog message
        astrRecordset(3, 0) = "Titel des referenzierten Angebots eingeben"
        ' input dialog title
        astrRecordset(4, 0) = "Angebot erfassen"
        ' target state for RecordExists - supposed to be true
        astrRecordset(5, 0) = CStr(True)
        ' recordset name - supposed to be empty
        astrRecordset(6, 0) = ""
    
        ' table B - the table on the n-side
        ' table name
        astrRecordset(0, 1) = "tblRechnung"
        ' key column name
        astrRecordset(1, 1) = "RechnungNr"
        ' title
        astrRecordset(2, 1) = "Rechnungsnummer"
        ' input dialog message
        astrRecordset(3, 1) = "Rechnungsnummer eingeben"
        ' input dialog title
        astrRecordset(4, 1) = "Rechnung erfassen"
        ' target state for RecordExists - supposed to be false
        astrRecordset(5, 1) = CStr(False)
        ' recordset name
        astrRecordset(6, 1) = ""
    
        ' assistance table - the m:n-table
        ' table name
        astrRecordset(0, 2) = "tblAngebotZuRechnung"

    ' check if tables exist
    Dim intError As Integer
    intError = 0
    
    For inti = LBound(astrRecordset, 2) To UBound(astrRecordset, 2)
        If basSupport.TabelleExistiert(astrRecordset(0, inti)) = False Then
            Debug.Print "clsAngebot.AddRecordset: " & astrRecordset(0, inti) _
                & " existiert nicht."
            intError = intError + 1
        Else:
            If mbolVerbatim = True Then
            Debug.Print "clsAngebot.AddRecordset: " + astrRecordset(0, inti) _
                + " existiert"
            End If
        End If
    Next
    
    If intError > 0 Then
        Debug.Print "clsAngebot.AddRecodset: Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    'ask for the name of the recordsets
    Dim lngi As Long
    For lngi = LBound(astrRecordset, 2) To UBound(astrRecordset, 2) - 1
        astrRecordset(6, lngi) = InputBox(astrRecordset(3, lngi), astrRecordset(4, lngi))
        
            ' check if inputbox is empty, if so send messagebox and exit procedure
            If basSupport.PflichtfeldIstLeer(astrRecordset(6, lngi)) = True Then
                Debug.Print "clsAngebot.AddRecordset: " & astrRecordset(2, lngi) & " ist leer, " _
                    & "Prozedur abgebrochen"
                GoTo ExitProc
            End If
        
            ' return input
            If mbolVerbatim = True Then
                Debug.Print "clsAngebot.AddRecordset: " & astrRecordset(2, lngi) _
                    & " = " & astrRecordset(4, lngi)
            End If
            
            ' check if recordset exists
            If CStr(basSupport.RecordsetExists(astrRecordset(0, lngi), astrRecordset(1, lngi), _
                astrRecordset(6, lngi))) <> astrRecordset(5, lngi) Then
                ' error message: send messagebox and exit procedure
                    If astrRecordset(5, lngi) = CStr(False) Then
                        Debug.Print "clsAngebot.AddRecordset: '" & astrRecordset(6, lngi) _
                            & "' existiert bereits. Prozedur abgebrochen."
                        MsgBox "'" & astrRecordset(6, lngi) & "' existiert bereits.", _
                            vbCritical, "Doppelter Eintrag"
                    Else:
                        Debug.Print "clsAngebot.AddRecordset: '" & astrRecordset(6, lngi) _
                            & "' existiert nicht. Prozedur abgebrochen."
                        MsgBox "'" & astrRecordset(6, lngi) & "' existiert nicht.", _
                            vbCritical, "Referenzdantensatz fehlt"
                    End If
                GoTo ExitProc
            End If
    Next
        
    ' create recordset in assistance table
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(astrRecordset(0, 2), dbOpenDynaset)
    
        mrstRecordset.AddNew
            mrstRecordset.Fields(astrRecordset(1, 0)) = astrRecordset(6, 0)
            mrstRecordset.Fields(astrRecordset(1, 1)) = astrRecordset(6, 1)
        mrstRecordset.Update
        
        ' close mrstRecordset
        mrstRecordset.Close
        Set mrstRecordset = Nothing
        
    ' create recordset in table B
    Set mrstRecordset = dbsCurrentDb.OpenRecordset(astrRecordset(0, 1), dbOpenDynaset)
    
        mrstRecordset.AddNew
            mrstRecordset!RechnungNr = astrRecordset(6, 1)
            mrstRecordset.Fields(astrRecordset(1, 1)) = astrRecordset(6, 1)
        mrstRecordset.Update
        
        ' copy offer name to clipboard
        mvarClipboard = astrRecordset(6, 1)
        
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

