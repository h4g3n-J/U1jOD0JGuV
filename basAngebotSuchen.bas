Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

Public Sub BuildAngebotSuchen()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen"
    End If
    
    ' define form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' declare temporary form name
    Dim strTempFormName As String

     ' clear form
    basSupport.ClearForm strFormName

    ' declare form
    Dim frm As Form
    Set frm = CreateForm
    
    ' write temporary form name to strFormName
    strTempFormName = frm.Name
    
    ' set form caption
    frm.Caption = strFormName
    
    ' get label settings
    Dim avarLabelSettings As Variant
    avarLabelSettings = basAngebotSuchen.LabelSettings
    Dim inti As Integer
    
    ' set labels
    Dim lbl As Label
    For inti = LBound(avarLabelSettings, 1) + 1 To UBound(avarLabelSettings, 1)
        Set lbl = CreateControl(strTempFormName, acLabel, acDetail)
        lbl.Name = avarLabelSettings(inti, 0)
        lbl.Left = avarLabelSettings(inti, 1)
        lbl.Top = avarLabelSettings(inti, 2)
        lbl.Width = avarLabelSettings(inti, 3)
        lbl.Height = avarLabelSettings(inti, 4)
        lbl.Caption = avarLabelSettings(inti, 5)
        Set lbl = Nothing
    Next
    
    ' set command buttons
    Dim avarCommandButtonSettings As Variant
    avarCommandButtonSettings = basAngebotSuchen.CommandButtonSettings
    Dim cmd As CommandButton
    
    For inti = LBound(avarCommandButtonSettings, 1) + 1 To UBound(avarCommandButtonSettings, 1)
        Set cmd = CreateControl(strTempFormName, acCommandButton, acDetail)
        cmd.Name = avarCommandButtonSettings(inti, 0)
        cmd.Left = avarCommandButtonSettings(inti, 1)
        cmd.Top = avarCommandButtonSettings(inti, 2)
        cmd.Width = avarCommandButtonSettings(inti, 3)
        cmd.Height = avarCommandButtonSettings(inti, 4)
        cmd.Caption = avarCommandButtonSettings(inti, 5)
        cmd.OnClick = avarCommandButtonSettings(inti, 6)
    Next
    
    ' set subform
    Dim avarSubFormSettings As Variant
    avarSubFormSettings = basAngebotSuchen.SubFormSettings
    Dim subFrm As SubForm
    
    For inti = LBound(avarSubFormSettings, 1) + 1 To UBound(avarSubFormSettings, 1)
        Set subFrm = CreateControl(strTempFormName, acSubform, acDetail)
        subFrm.Name = avarSubFormSettings(inti, 0)
        subFrm.Left = avarSubFormSettings(inti, 1)
        subFrm.Top = avarSubFormSettings(inti, 2)
        subFrm.Width = avarSubFormSettings(inti, 3)
        subFrm.Height = avarSubFormSettings(inti, 4)
        subFrm.SourceObject = avarSubFormSettings(inti, 5)
    Next
    
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngbotSuchen: " & strFormName & " erstellt"
    End If

End Sub

Private Function LabelSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.LabelSettings ausfuehren"
    End If
    
    Dim avarSettings(1, 5) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
        avarSettings(0, 5) = "Caption"
    avarSettings(1, 0) = "lblTitle"
        avarSettings(1, 1) = 566
        avarSettings(1, 2) = 227
        avarSettings(1, 3) = 9210
        avarSettings(1, 4) = 507
        avarSettings(1, 5) = "Angebot Suchen"
        
    LabelSettings = avarSettings
End Function

Private Function CommandButtonSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CommandButtonSettings ausfuehren"
    End If

    Dim avarSettings(1, 6) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
        avarSettings(0, 5) = "Caption"
        avarSettings(0, 6) = "OnClick"
    avarSettings(1, 0) = "cmdExit"
        avarSettings(1, 1) = 12585
        avarSettings(1, 2) = 960
        avarSettings(1, 3) = 3120
        avarSettings(1, 4) = 330
        avarSettings(1, 5) = "Schlieﬂen"
        avarSettings(1, 6) = "=CloseFrmAngebotSuchen()"
        
    CommandButtonSettings = avarSettings
End Function

Private Function SubFormSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SubFormSettings ausfuehren"
    End If
    
    Dim avarSubFormSettings(1, 5) As Variant
    
    avarSubFormSettings(0, 0) = "Name"
        avarSubFormSettings(0, 1) = "Left"
        avarSubFormSettings(0, 2) = "Top"
        avarSubFormSettings(0, 3) = "Width"
        avarSubFormSettings(0, 4) = "Height"
        avarSubFormSettings(0, 5) = "SourceObject"
    avarSubFormSettings(1, 0) = "subFrm"
        avarSubFormSettings(1, 1) = 510
        avarSubFormSettings(1, 2) = 2453
        avarSubFormSettings(1, 3) = 9218
        avarSubFormSettings(1, 4) = 5055
        avarSubFormSettings(1, 5) = "frmAngebotSuchenSub"
        
    SubFormSettings = avarSubFormSettings
End Function

Public Function CloseFrmAngebotSuchen()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function
