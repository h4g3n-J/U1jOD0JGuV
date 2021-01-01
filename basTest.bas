Attribute VB_Name = "basTest"
' basTest

Option Compare Database
Option Explicit

Private Sub TestAuftragHinzufuegen()
    Dim TestKlasse As clsAuftrag
    Set TestKlasse = New clsAuftrag
    
    ' TestKlasse.Verbatim = False
    TestKlasse.AddRecordset
End Sub

Private Sub TestAngebotHinzufuegen()
    Dim TestKlasse As clsAngebot
    Set TestKlasse = New clsAngebot
    
    TestKlasse.Verbatim = True
    TestKlasse.AddRecordset
End Sub

Private Sub TestRecordsetExists()
    Debug.Print "TestRecordsetExists: " _
        & basSupport.RecordsetExists("tblAuftrag", "AftrID", "345")
End Sub

Private Sub TestForEach()
    Dim astrTable(1) As String
    Dim varTableName As Variant
    
    astrTable(0) = "tblAngebot"
    astrTable(1) = "tblAuftrag"
    
    For Each varTableName In astrTable
        Debug.Print varTableName
    Next
End Sub

Private Sub TestForNext()
    Dim astrTable(1, 2) As String
    Dim lngi As Long
    Dim lngj As Long
    
    Dim bolTest As Boolean
    bolTest = True
    
    astrTable(0, 0) = "test 0, 0"
    astrTable(0, 1) = "test 0, 1"
    astrTable(0, 2) = "test 0, 2"
    astrTable(1, 0) = "test 1, 0"
    astrTable(1, 1) = "test 1, 1"
    astrTable(1, 2) = "test 1, 2"
    
    astrTable(0, 0) = "test 0, 0"
    astrTable(0, 1) = "test 0, 1"
    astrTable(0, 2) = "test 0, 2"
    astrTable(1, 0) = "test 1, 0"
    astrTable(1, 1) = CStr(bolTest)
    astrTable(1, 2) = "test 1, 2"
    
    For lngi = LBound(astrTable, 1) To UBound(astrTable, 1)
        For lngj = LBound(astrTable, 2) To UBound(astrTable, 2)
            Debug.Print astrTable(lngi, lngj)
        Next
    Next
    
    Debug.Print "CBool(astrTable(1, 1)): "; CBool(astrTable(1, 1))
End Sub

Private Sub TestRechnungHinzufuegen()
    Dim TestKlasse As clsRechnung
    Set TestKlasse = New clsRechnung
    
    TestKlasse.Verbatim = True
    TestKlasse.AddRecordset
End Sub

Private Sub testLeistungserfassungsblattHinzufuegen()
    Dim TestKlasse As clsLeistungserfassungungsblatt
    Set TestKlasse = New clsLeistungserfassungungsblatt
    
    TestKlasse.Verbatim = True
    TestKlasse.AddRecordset
End Sub
