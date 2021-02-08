Attribute VB_Name = "basAuftragSuchenSub"
Option Compare Database
Option Explicit

Public Sub Test()
    If gconVerbatim = True Then
        Debug.Print "basAuftrag.SuchenSub.Test ausfuehren"
    End If
    
    basSearchMain.ShowRecordset Forms.Item("frmSearchMain").Controls("frb1").Controls("AftrID")
End Sub

