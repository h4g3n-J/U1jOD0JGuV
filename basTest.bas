Attribute VB_Name = "basTest"
' basTest

Option Compare Database
Option Explicit

Public Sub TestAuftragHinzufuegen()
    Dim TestKlasse As clsAuftrag
    Set TestKlasse = New clsAuftrag
    
    ' TestKlasse.Verbatim = False
    TestKlasse.AddRecordset
End Sub

Public Sub TestAngebotHinzufuegen()
    Dim TestKlasse As clsAngebot
    Set TestKlasse = New clsAngebot
    
    TestKlasse.Verbatim = True
    TestKlasse.AddRecordset
End Sub

Public Sub TestRecordsetExists()
    Debug.Print "TestRecordsetExists: " _
        & basSupport.RecordsetExists("tblAuftrag", "AftrID", "345")
End Sub
