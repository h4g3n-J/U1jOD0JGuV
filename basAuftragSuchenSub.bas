Attribute VB_Name = "basAuftragSuchenSub"
Option Compare Database
Option Explicit

Public Sub SelectRecordset()
    ' initiate form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basAuftragSuchenSub.SelectRecordset ausfuehren"
    End If
    
    ' error handler, case strFormName is not loaded
    If CurrentProject.AllForms(strFormName).IsLoaded Then
        basSearchMain.ShowRecordset Forms.Item(strFormName).Controls("frb1").Controls("AftrID")
    End If
End Sub

