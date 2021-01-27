VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMenueExtraSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmAuftragEdit

Option Compare Database
Option Explicit

Private Sub cmd0_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmd1_Click()
    If gobjMenueExtraSmall Is Nothing Then
        Debug.Print "Form_fmrMenueExtraSmall: " _
            & "gobjMenueExtraSmall nicht instanziiert, " _
            & "speichern nicht erfolgreich"
        Exit Sub
    End If
        
    gobjMenueExtraSmall.Speichern
    
End Sub

Private Sub cmd2_Click()

End Sub
