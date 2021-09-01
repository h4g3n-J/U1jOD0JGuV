Attribute VB_Name = "basTestRechnung"
Option Compare Database
Option Explicit

Private Sub RechnungNr()

    ' command message
    Debug.Print "execute basTestRechnung.RechnungNr"
    
    Dim rstRecordset As clsRechnung
    Set rstRecordset = New clsRechnung
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.RechnungNr = strTestValue
    
    If Not rstRecordset.RechnungNr = strTestValue Then
        MsgBox "basTestRechnung.RechnungNr failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestRechnung.RechnungNr: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestRechnung.RechnunugNr executed"
    
End Sub

Private Sub Bemerkung()

    ' command message
    Debug.Print "execute basTestRechnung.Bemerkung"
    
    Dim rstRecordset As clsRechnung
    Set rstRecordset = New clsRechnung
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.Bemerkung = strTestValue
    
    If Not rstRecordset.Bemerkung = strTestValue Then
        MsgBox "basTestRechnung.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestRechnung.Bemerkung: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestRechnung.Bemerkung executed"

End Sub

Private Sub RechnungLink()

    ' command message
    Debug.Print "execute basTestRechnung.RechnungLink"
    
    Dim rstRecordset As clsRechnung
    Set rstRecordset = New clsRechnung
    
    Dim strTestValue As String
    strTestValue = "#Test#"
    
    rstRecordset.RechnungLink = strTestValue
    
    If Not rstRecordset.RechnungLink = strTestValue Then
        MsgBox "basTestRechnung.RechnungLink failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestRechnung.RechnungLink: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestRechnung.RechnungLink executed"
    
End Sub
