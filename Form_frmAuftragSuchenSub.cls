VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAuftragSuchenSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form_frmAuftragSuchenSub

Option Compare Database
Option Explicit

Private bolVerbatim As Boolean

Private Sub Form_Current()
    
    If gconVerbatim Then
        Debug.Print "Form_frmAuftragSuchenSub.Form_Current ausfuehren"
    End If
    
    basAuftragSuchenSub.SelectRecordsetAuftrag
End Sub
