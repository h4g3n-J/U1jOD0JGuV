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
    
    ' get textbox settings
    Dim avarTextboxSettings As Variant
    avarTextboxSettings = basAngebotSuchen.TextBoxSettings
    Dim txt As TextBox
    
    ' set textbox
    Dim inti As Integer
    For inti = LBound(avarTextboxSettings, 1) + 1 To UBound(avarTextboxSettings, 1)
        Set txt = CreateControl(strTempFormName, acTextBox, acDetail)
        txt.Name = avarTextboxSettings(inti, 0)
        txt.Left = avarTextboxSettings(inti, 1)
        txt.Top = avarTextboxSettings(inti, 2)
        txt.Width = avarTextboxSettings(inti, 3)
        txt.Height = avarTextboxSettings(inti, 4)
        Set txt = Nothing
    Next
    
    ' get label settings
    Dim avarLabelSettings As Variant
    avarLabelSettings = basAngebotSuchen.LabelSettings
    Dim lbl As Label
    
    ' set labels
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
    
    ' get commandbutton settings
    Dim avarCommandButtonSettings As Variant
    avarCommandButtonSettings = basAngebotSuchen.CommandButtonSettings
    Dim cmd As CommandButton
    
    ' set command button
    For inti = LBound(avarCommandButtonSettings, 1) + 1 To UBound(avarCommandButtonSettings, 1)
        Set cmd = CreateControl(strTempFormName, acCommandButton, acDetail)
        cmd.Name = avarCommandButtonSettings(inti, 0)
        cmd.Left = avarCommandButtonSettings(inti, 1)
        cmd.Top = avarCommandButtonSettings(inti, 2)
        cmd.Width = avarCommandButtonSettings(inti, 3)
        cmd.Height = avarCommandButtonSettings(inti, 4)
        cmd.Caption = avarCommandButtonSettings(inti, 5)
        cmd.OnClick = avarCommandButtonSettings(inti, 6)
        Set cmd = Nothing
    Next
    
    ' get subform settings
    Dim avarSubFormSettings As Variant
    avarSubFormSettings = basAngebotSuchen.SubFormSettings
    Dim subFrm As SubForm
    
    ' set subform
    For inti = LBound(avarSubFormSettings, 1) + 1 To UBound(avarSubFormSettings, 1)
        Set subFrm = CreateControl(strTempFormName, acSubform, acDetail)
        subFrm.Name = avarSubFormSettings(inti, 0)
        subFrm.Left = avarSubFormSettings(inti, 1)
        subFrm.Top = avarSubFormSettings(inti, 2)
        subFrm.Width = avarSubFormSettings(inti, 3)
        subFrm.Height = avarSubFormSettings(inti, 4)
        subFrm.SourceObject = avarSubFormSettings(inti, 5)
        subFrm.Locked = avarSubFormSettings(inti, 6)
        Set subFrm = Nothing
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

    Dim avarSettings(2, 6) As Variant
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
    avarSettings(2, 0) = "cmdSearch"
        avarSettings(2, 1) = 6975
        avarSettings(2, 2) = 960
        avarSettings(2, 3) = 2730
        avarSettings(2, 4) = 330
        avarSettings(2, 5) = "Suchen"
        avarSettings(2, 6) = "=CloseFrmAngebotSuchen()"
        
    CommandButtonSettings = avarSettings
End Function

Private Function TextBoxSettings() As Variant

    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CommandButtonSettings ausfuehren"
    End If

    Dim intLeft As Integer
    intLeft = 12585
    
    Dim intWidth As Integer
    intWidth = 3120
    
    Dim intHeight As Integer
    intHeight = 330
    
    Dim avarSettings(4, 5) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
    avarSettings(1, 0) = "txtSearch"
        avarSettings(1, 1) = 510
        avarSettings(1, 2) = 960
        avarSettings(1, 3) = 6405
        avarSettings(1, 4) = intHeight
    avarSettings(2, 0) = "txt0"
        avarSettings(2, 1) = intLeft
        avarSettings(2, 2) = 2430
        avarSettings(2, 3) = intWidth
        avarSettings(2, 4) = intHeight
    avarSettings(3, 0) = "txt1"
        avarSettings(3, 1) = intLeft
        avarSettings(3, 2) = 2820
        avarSettings(3, 3) = intWidth
        avarSettings(3, 4) = intHeight
    avarSettings(4, 0) = "txt2"
        avarSettings(4, 1) = intLeft
        avarSettings(4, 2) = 3210
        avarSettings(4, 3) = intWidth
        avarSettings(4, 4) = intHeight
    
    TextBoxSettings = avarSettings
End Function

Private Function SubFormSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SubFormSettings ausfuehren"
    End If
    
    Dim avarSubFormSettings(1, 6) As Variant
    
    avarSubFormSettings(0, 0) = "Name"
        avarSubFormSettings(0, 1) = "Left"
        avarSubFormSettings(0, 2) = "Top"
        avarSubFormSettings(0, 3) = "Width"
        avarSubFormSettings(0, 4) = "Height"
        avarSubFormSettings(0, 5) = "SourceObject"
        avarSubFormSettings(0, 6) = "Locked"
    avarSubFormSettings(1, 0) = "subFrm"
        avarSubFormSettings(1, 1) = 510
        avarSubFormSettings(1, 2) = 2453
        avarSubFormSettings(1, 3) = 9218
        avarSubFormSettings(1, 4) = 5055
        avarSubFormSettings(1, 5) = "frmAngebotSuchenSub"
        avarSubFormSettings(1, 6) = True
        
    SubFormSettings = avarSubFormSettings
End Function

Public Function CloseFrmAngebotSuchen()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function
