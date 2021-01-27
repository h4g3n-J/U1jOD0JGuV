Attribute VB_Name = "basMenueExtraSmall"
' basMenueExtraSmall

Option Compare Database
Option Explicit

Public Sub AuftragErstellen()
    DoCmd.OpenForm "frmMenueExtraSmall", acNormal
    
    Dim avarLayoutMenueExtraSmall(5, 5) As Variant
    
    ' 0 = label name
    ' 1 = label caption
    ' 2 = label isVisible
    ' 3 = textbox name
    ' 4 = textbox caption
    ' 5 = textbox visible
    avarLayoutMenueExtraSmall(0, 0) = Null
        avarLayoutMenueExtraSmall(1, 0) = Null
        avarLayoutMenueExtraSmall(2, 0) = False
        avarLayoutMenueExtraSmall(3, 0) = "txt0"
        avarLayoutMenueExtraSmall(4, 0) = Null
        avarLayoutMenueExtraSmall(5, 0) = False
    avarLayoutMenueExtraSmall(0, 1) = Null
        avarLayoutMenueExtraSmall(1, 1) = Null
        avarLayoutMenueExtraSmall(2, 1) = False
        avarLayoutMenueExtraSmall(3, 1) = "txt1"
        avarLayoutMenueExtraSmall(4, 1) = Null
        avarLayoutMenueExtraSmall(5, 1) = False
    avarLayoutMenueExtraSmall(0, 2) = "lbl2"
        avarLayoutMenueExtraSmall(1, 2) = "ID"
        avarLayoutMenueExtraSmall(2, 2) = True
        avarLayoutMenueExtraSmall(3, 2) = "txt2"
        avarLayoutMenueExtraSmall(4, 2) = Null
        avarLayoutMenueExtraSmall(5, 2) = True
    avarLayoutMenueExtraSmall(0, 3) = "lbl3"
        avarLayoutMenueExtraSmall(1, 3) = "Titel"
        avarLayoutMenueExtraSmall(2, 3) = True
        avarLayoutMenueExtraSmall(3, 3) = "txt3"
        avarLayoutMenueExtraSmall(4, 3) = Null
        avarLayoutMenueExtraSmall(5, 3) = True
    avarLayoutMenueExtraSmall(0, 4) = "lbl4"
        avarLayoutMenueExtraSmall(1, 4) = Null
        avarLayoutMenueExtraSmall(2, 4) = True
        avarLayoutMenueExtraSmall(3, 4) = "txt4"
        avarLayoutMenueExtraSmall(4, 4) = Null
        avarLayoutMenueExtraSmall(5, 4) = False
    avarLayoutMenueExtraSmall(0, 5) = "lbl5"
        avarLayoutMenueExtraSmall(1, 5) = Null
        avarLayoutMenueExtraSmall(2, 5) = True
        avarLayoutMenueExtraSmall(3, 5) = "txt5"
        avarLayoutMenueExtraSmall(4, 5) = Null
        avarLayoutMenueExtraSmall(5, 5) = False
        
        Forms.Item("frmMenueExtraSmall").Controls("txt2").SetFocus
        
    Dim inti As Integer
    For inti = LBound(avarLayoutMenueExtraSmall, 2) To UBound(avarLayoutMenueExtraSmall, 2)
        
        ' set label
        If Not (IsNull(avarLayoutMenueExtraSmall(0, inti)) Or IsNull(avarLayoutMenueExtraSmall(1, inti))) Then
            ' set label caption
            Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(0, inti)).Caption = avarLayoutMenueExtraSmall(1, inti)
            ' set label visibility
            Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(0, inti)).Visible = avarLayoutMenueExtraSmall(2, inti)
        End If
        
        ' set textbox
        ' set textbox visibility
        Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(3, inti)).Visible = avarLayoutMenueExtraSmall(5, inti)
    Next
    
End Sub

