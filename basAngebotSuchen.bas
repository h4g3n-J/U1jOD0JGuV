Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

Private Function tableSettings(ByVal intColumn As Integer, intRow As Integer, intProperty As Integer) As Integer
    
    Const cintCol00Left As Integer = 10000
    Const cintCol00Width As Integer = 2540
    
    Const cintHorizontalSpacing As Integer = 60
    
    Const cintCol01Left As Integer = cintCol00Left + cintCol00Width + cintHorizontalSpacing
    Const cintCol01Width As Integer = 3120
    
    Const cintHeight As Integer = 330
    
    Const cintVerticalSpacing As Integer = 60
    
    Const cintRow00Top As Integer = 2430
    Const cintRow01Top As Integer = cintRow00Top + cintHeight + cintVerticalSpacing
    Const cintRow02Top As Integer = cintRow01Top + cintHeight + cintVerticalSpacing
    
    Dim aintTableSettings(1, 2, 3) As Integer
    ' avarSettings(0, 1) = "Left"
    ' avarSettings(0, 2) = "Top"
    ' avarSettings(0, 3) = "Width"
    ' avarSettings(0, 4) = "Height"
    
    ' column0, row 0
    aintTableSettings(0, 0, 0) = cintCol00Left
    aintTableSettings(0, 0, 1) = cintRow00Top
    aintTableSettings(0, 0, 2) = cintCol00Width
    aintTableSettings(0, 0, 3) = cintHeight
    
    ' column1, row 0
    aintTableSettings(1, 0, 0) = cintCol01Left
    aintTableSettings(1, 0, 1) = cintRow00Top
    aintTableSettings(1, 0, 2) = cintCol01Width
    aintTableSettings(1, 0, 3) = cintHeight
    
    ' column0, row 1
    aintTableSettings(0, 1, 0) = cintCol00Left
    aintTableSettings(0, 1, 1) = cintRow01Top
    aintTableSettings(0, 1, 2) = cintCol00Width
    aintTableSettings(0, 1, 3) = cintHeight
    
    ' column1, row 1
    aintTableSettings(1, 1, 0) = cintCol01Left
    aintTableSettings(1, 1, 1) = cintRow01Top
    aintTableSettings(1, 1, 2) = cintCol01Width
    aintTableSettings(1, 1, 3) = cintHeight
    
    ' column0, row2
    aintTableSettings(0, 2, 0) = cintCol00Left
    aintTableSettings(0, 2, 1) = cintRow02Top
    aintTableSettings(0, 2, 2) = cintCol00Width
    aintTableSettings(0, 2, 3) = cintHeight
    
    ' column1, row 2
    aintTableSettings(1, 2, 0) = cintCol01Left
    aintTableSettings(1, 2, 1) = cintRow02Top
    aintTableSettings(1, 2, 2) = cintCol01Width
    aintTableSettings(1, 2, 3) = cintHeight

    tableSettings = aintTableSettings(intColumn, intRow, intProperty)
    
End Function
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
    
    ' textboxes
    ' get textbox settings
    Dim avarTextboxSettings As Variant
    avarTextboxSettings = basAngebotSuchen.TextBoxSettings
    Dim txt As TextBox
    
    ' create textboxes
    Dim inti As Integer
    Dim intj As Integer
    ' skip propertie name and datatype => + 2
    For inti = LBound(avarTextboxSettings, 1) + 2 To UBound(avarTextboxSettings, 1)
        Set txt = CreateControl(strTempFormName, acTextBox, acDetail)
        ' set textbox properties
        ' get propertie name: avarTextboxSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarTextboxSettings(inti, intj)
        ' propertie value: avarTextboxSettings(1, intj)
        For intj = LBound(avarTextboxSettings, 2) To UBound(avarTextboxSettings, 2)
            CallByName txt, avarTextboxSettings(0, intj), VbLet, basSupport.CheckDataType(avarTextboxSettings(inti, intj), avarTextboxSettings(1, intj))
        Next
    Next
    
    'labels
    ' get label settings
    Dim avarLabelSettings As Variant
    avarLabelSettings = basAngebotSuchen.LabelSettings
    Dim lbl As Label
    
    ' create labels
    ' skip propertie name and datatype => + 2
    For inti = LBound(avarLabelSettings, 1) + 2 To UBound(avarLabelSettings, 1)
        ' create label
        Set lbl = CreateControl(strTempFormName, acLabel, acDetail)
        ' set label properties
        ' get propertie name: avarLabelSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarLabelSettings(inti, intj)
        ' propertie value: avarLabelSettings(1, intj)
        For intj = LBound(avarLabelSettings, 2) To UBound(avarLabelSettings, 2)
            CallByName lbl, avarLabelSettings(0, intj), VbLet, basSupport.CheckDataType(avarLabelSettings(inti, intj), avarLabelSettings(1, intj))
        Next
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
    
    Dim avarSettings(5, 6) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
        avarSettings(0, 5) = "Caption"
        avarSettings(0, 6) = "Visible"
    avarSettings(1, 0) = "string"
        avarSettings(1, 1) = "integer"
        avarSettings(1, 2) = "integer"
        avarSettings(1, 3) = "integer"
        avarSettings(1, 4) = "integer"
        avarSettings(1, 5) = "string"
        avarSettings(1, 6) = "boolean"
    avarSettings(2, 0) = "lblTitle"
        avarSettings(2, 1) = 566
        avarSettings(2, 2) = 227
        avarSettings(2, 3) = 9210
        avarSettings(2, 4) = 507
        avarSettings(2, 5) = "Angebot Suchen"
        avarSettings(2, 6) = True
    avarSettings(3, 0) = "lbl00"
        avarSettings(3, 1) = basAngebotSuchen.tableSettings(0, 0, 0)
        avarSettings(3, 2) = basAngebotSuchen.tableSettings(0, 0, 1)
        avarSettings(3, 3) = basAngebotSuchen.tableSettings(0, 0, 2)
        avarSettings(3, 4) = basAngebotSuchen.tableSettings(0, 0, 3)
        avarSettings(3, 5) = "Angebot"
        avarSettings(3, 6) = True
    avarSettings(4, 0) = "lbl01"
        avarSettings(4, 1) = basAngebotSuchen.tableSettings(0, 1, 0)
        avarSettings(4, 2) = basAngebotSuchen.tableSettings(0, 1, 1)
        avarSettings(4, 3) = basAngebotSuchen.tableSettings(0, 1, 2)
        avarSettings(4, 4) = basAngebotSuchen.tableSettings(0, 1, 3)
        avarSettings(4, 5) = "Einzelauftrag"
        avarSettings(4, 6) = True
    avarSettings(5, 0) = "lbl02"
        avarSettings(5, 1) = basAngebotSuchen.tableSettings(0, 2, 0)
        avarSettings(5, 2) = basAngebotSuchen.tableSettings(0, 2, 1)
        avarSettings(5, 3) = basAngebotSuchen.tableSettings(0, 2, 2)
        avarSettings(5, 4) = basAngebotSuchen.tableSettings(0, 2, 3)
        avarSettings(5, 5) = "test"
        avarSettings(5, 6) = False
        
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
    
    Dim avarSettings(5, 5) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
        avarSettings(0, 5) = "Visible"
    avarSettings(1, 0) = "string"
        avarSettings(1, 1) = "integer"
        avarSettings(1, 2) = "integer"
        avarSettings(1, 3) = "integer"
        avarSettings(1, 4) = "integer"
        avarSettings(1, 5) = "boolean"
    avarSettings(2, 0) = "txtSearch"
        avarSettings(2, 1) = 510
        avarSettings(2, 2) = 960
        avarSettings(2, 3) = 6405
        avarSettings(2, 4) = 330
        avarSettings(2, 5) = True
    avarSettings(3, 0) = "txt0"
        avarSettings(3, 1) = basAngebotSuchen.tableSettings(1, 0, 0)
        avarSettings(3, 2) = basAngebotSuchen.tableSettings(1, 0, 1)
        avarSettings(3, 3) = basAngebotSuchen.tableSettings(1, 0, 2)
        avarSettings(3, 4) = basAngebotSuchen.tableSettings(1, 0, 3)
        avarSettings(3, 5) = True
    avarSettings(4, 0) = "txt1"
        avarSettings(4, 1) = basAngebotSuchen.tableSettings(1, 1, 0)
        avarSettings(4, 2) = basAngebotSuchen.tableSettings(1, 1, 1)
        avarSettings(4, 3) = basAngebotSuchen.tableSettings(1, 1, 2)
        avarSettings(4, 4) = basAngebotSuchen.tableSettings(1, 1, 3)
        avarSettings(4, 5) = True
    avarSettings(5, 0) = "txt2"
        avarSettings(5, 1) = basAngebotSuchen.tableSettings(1, 2, 0)
        avarSettings(5, 2) = basAngebotSuchen.tableSettings(1, 2, 1)
        avarSettings(5, 3) = basAngebotSuchen.tableSettings(1, 2, 2)
        avarSettings(5, 4) = basAngebotSuchen.tableSettings(1, 2, 3)
        avarSettings(5, 5) = False
    
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
    avarSubFormSettings(1, 0) = "frbSubFrm"
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
