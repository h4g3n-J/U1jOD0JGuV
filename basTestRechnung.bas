Attribute VB_Name = "basTestRechnung"
Option Compare Database
Option Explicit

Private Sub RechnungNr()

    ' command message
    Debug.Print "execute RechnungNr"
    
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
    Debug.Print "RechnunugNr executed"
    
End Sub
