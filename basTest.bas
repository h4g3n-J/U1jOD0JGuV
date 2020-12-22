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
