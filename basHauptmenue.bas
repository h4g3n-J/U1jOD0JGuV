Attribute VB_Name = "basHauptmenue"
' basHautpmenue

Option Compare Database
Option Explicit

Private avarLayout(2, 1) As Variant

Private Sub LayoutConfig()
    ' 0 = object name
    ' 1 = object caption
    ' 2 = object visible
    avarLayout(0, 0) = "cmd0"
        avarLayout(1, 0) = "Ticket suchen"
        avarLayout(2, 0) = True
    avarLayout(0, 1) = "cmd1"
        avarLayout(1, 1) = "Liefergegenstand suchen"
        avarLayout(2, 1) = True
End Sub

' open frmSearchMain and set textboxes and labels
Public Sub OpenFormHauptmenue()
    ' DoCmd.OpenForm "frmHauptmenue", acNormal
    
    If gconVerbatim = True Then
        Debug.Print "basHauptmenue.OpenFormHauptmenue ausfuehren"
    End If
    
    ' Set Form name
    Dim strFormName As String
    strFormName = "frmHauptmenue"
    
    ' initialize LayoutConfig
    LayoutConfig
    
    ' set labels and textboxes
    Dim inti As Integer
        
    ' set command buttons
    For inti = LBound(avarLayout, 2) To UBound(avarLayout, 2)
        ' set caption
        Forms.Item(strFormName).Controls.Item(avarLayout(0, inti)).Caption = avarLayout(1, inti)
        ' set visibility
        Forms.Item(strFormName).Controls.Item(avarLayout(0, inti)).Visible = avarLayout(2, inti)
    Next
    
End Sub

