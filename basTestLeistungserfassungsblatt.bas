Attribute VB_Name = "basTestLeistungserfassungsblatt"
Option Compare Database
Option Explicit

Private Sub Leistungserfassungsblatt()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Leistungserfassungsblatt"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.Leistungserfassungsblatt = strTestValue
    
    If Not rstRecordset.Leistungserfassungsblatt = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Leistungserfassungsblatt failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Leistungserfassungsblatt: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.RechnunugNr executed"
    
End Sub
