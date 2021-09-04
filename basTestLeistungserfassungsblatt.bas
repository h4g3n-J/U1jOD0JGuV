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

Private Sub RechnungNr()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.RechnungNr"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.RechnungNr = strTestValue
    
    If Not rstRecordset.RechnungNr = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.RechnungNr failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.RechnungNr: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.RechnunugNr executed"
    
End Sub

Private Sub Bemerkung()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Bemerkung"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.Bemerkung = strTestValue
    
    If Not rstRecordset.Bemerkung = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Bemerkung: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.Bemerkung executed"
    
End Sub

Private Sub BelegID()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.BelegID"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    rstRecordset.BelegID = strTestValue
    
    If Not rstRecordset.BelegID = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.BelegID failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.BelegID: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.BelegID executed"
    
End Sub

Private Sub Brutto()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Brutto"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim curTestValue As Currency
    curTestValue = 123.45
    
    rstRecordset.Brutto = curTestValue
    
    If Not rstRecordset.Brutto = curTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Brutto: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.Brutto executed"
    
End Sub
