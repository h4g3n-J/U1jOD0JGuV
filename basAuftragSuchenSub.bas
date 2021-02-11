Attribute VB_Name = "basAuftragSuchenSub"
Option Compare Database
Option Explicit

Public Sub SelectRecordset()
    If gconVerbatim = True Then
        Debug.Print "basAuftrag.SuchenSub.SelectRecordset ausfuehren"
    End If
    
    basSearchMain.ShowRecordset Forms.Item("frmSearchMain").Controls("frb1").Controls("AftrID")
End Sub

