Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

' returns array
' calculates left, top, width and height parameters
Private Function CreateTable(Optional ByVal intNumberOfRows As Integer = 6)
    
    Const cintNumberOfColumns As Integer = 1
    ' Const cintNumberOfRows As Integer = intNumberOfRows - 1
    ' Const cintNumberOfRows As Integer = 5
    intNumberOfRows = intNumberOfRows - 1
    
    Const cintCol00Left As Integer = 10000
    Const cintCol00Width As Integer = 2540
    Const cintHorizontalSpacingCol00Col01 As Integer = 60
    Const cintCol01Width As Integer = 3120
        
    Const cintRow00Top As Integer = 2430
    Const cintRowHeight As Integer = 330
    Const cintVerticalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    ' avarSettings(0, 0) = "Left"
    ' avarSettings(0, 1) = "Top"
    ' avarSettings(0, 2) = "Width"
    ' avarSettings(0, 3) = "Height"
        
    Dim inti As Integer
    Dim intj As Integer
    
    Dim aintTableSettings() As Integer
    ReDim aintTableSettings(cintNumberOfColumns, intNumberOfRows, cintNumberOfProperties)
    
    For inti = 0 To cintNumberOfColumns
        For intj = 0 To intNumberOfRows
            ' set column left
            aintTableSettings(inti, intj, 0) = cintCol00Left + inti * (cintCol00Width + cintHorizontalSpacingCol00Col01)
            ' aintTableSettings(inti, intj, 0) = cintCol00Left
            ' aintTableSettings(0, 0, 0) = CStr(cintCol00Left)
            ' set row top
            aintTableSettings(inti, intj, 1) = cintRow00Top + intj * (cintRowHeight + cintVerticalSpacing)
            ' set column width
            Select Case inti
                ' column 0
                Case 0
                    aintTableSettings(inti, intj, 2) = cintCol00Width
                ' column 1
                Case 1
                    aintTableSettings(inti, intj, 2) = cintCol01Width
            End Select
            ' set row height
            aintTableSettings(inti, intj, 3) = cintRowHeight
        Next
    Next
        
    ' ReDim Preserve aintTableSettings(cintNumberOfColumns, intNumberOfRows, cintNumberOfProperties)

    Debug.Print "basAngebotSuchen.CreateTable ausgefuehrt"
    CreateTable = aintTableSettings

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
    Dim txtTextbox As TextBox
    
    ' create textboxes
    Dim inti As Integer
    Dim intj As Integer
    ' skip propertie name and datatype => + 2
    For inti = LBound(avarTextboxSettings, 1) + 2 To UBound(avarTextboxSettings, 1)
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        ' set textbox properties
        ' get property name: avarTextboxSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarTextboxSettings(inti, intj)
        ' propertie value: avarTextboxSettings(1, intj)
        For intj = LBound(avarTextboxSettings, 2) To UBound(avarTextboxSettings, 2)
            CallByName txtTextbox, avarTextboxSettings(0, intj), VbLet, basSupport.CheckDataType(avarTextboxSettings(inti, intj), avarTextboxSettings(1, intj))
        Next
    Next
    
    ' labels
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
        ' get property name: avarLabelSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarLabelSettings(inti, intj)
        ' propertie value: avarLabelSettings(1, intj)
        For intj = LBound(avarLabelSettings, 2) To UBound(avarLabelSettings, 2)
            CallByName lbl, avarLabelSettings(0, intj), VbLet, basSupport.CheckDataType(avarLabelSettings(inti, intj), avarLabelSettings(1, intj))
        Next
    Next
    
    ' command buttons
    ' get commandbutton settings
    Dim avarCommandButtonSettings As Variant
    avarCommandButtonSettings = basAngebotSuchen.CommandButtonSettings
    Dim cmd As CommandButton
    
    ' create command buttons
    ' skip propertie name and datatype => + 2
    For inti = LBound(avarCommandButtonSettings, 1) + 2 To UBound(avarCommandButtonSettings, 1)
        ' create command button
        Set cmd = CreateControl(strTempFormName, acCommandButton, acDetail)
        ' set command button properties
        ' get property name: avarCommandButtonSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarCommandButtonSettings(inti, intj)
        ' propertie value: avarCommandButtonSettings(1, intj)
        For intj = LBound(avarCommandButtonSettings, 2) To UBound(avarCommandButtonSettings, 2)
            CallByName cmd, avarCommandButtonSettings(0, intj), VbLet, basSupport.CheckDataType(avarCommandButtonSettings(inti, intj), avarCommandButtonSettings(1, intj))
        Next
    Next
    
    ' subform
    ' get subform settings
    Dim avarSubFormSettings As Variant
    avarSubFormSettings = basAngebotSuchen.SubFormSettings
    Dim subFrm As SubForm
    
    ' create subform
    ' skip propertie name and datatype => + 2
    For inti = LBound(avarSubFormSettings, 1) + 2 To UBound(avarSubFormSettings, 1)
        ' create subform
        Set subFrm = CreateControl(strTempFormName, acSubform, acDetail)
        ' set subform properties
        ' get property name: avarSubFormSettings(0, intj)
        ' transform datatype: basSupport.CheckDataType
        ' with destination datatype: avarSubFormSettings(inti, intj)
        ' propertie value: avarSubFormSettings(1, intj)
        For intj = LBound(avarSubFormSettings, 2) To UBound(avarSubFormSettings, 2)
            CallByName subFrm, avarSubFormSettings(0, intj), VbLet, basSupport.CheckDataType(avarSubFormSettings(inti, intj), avarSubFormSettings(1, intj))
        Next
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
    
    ' get table parameters
    Dim intTableSettings() As Integer
    intTableSettings = basAngebotSuchen.CreateTable(5)
    
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
        avarSettings(3, 1) = intTableSettings(0, 0, 0)
        avarSettings(3, 2) = intTableSettings(0, 0, 1)
        avarSettings(3, 3) = intTableSettings(0, 0, 2)
        avarSettings(3, 4) = intTableSettings(0, 0, 3)
        avarSettings(3, 5) = "Angebot"
        avarSettings(3, 6) = True
    avarSettings(4, 0) = "lbl01"
        avarSettings(4, 1) = intTableSettings(0, 1, 0)
        avarSettings(4, 2) = intTableSettings(0, 1, 1)
        avarSettings(4, 3) = intTableSettings(0, 1, 2)
        avarSettings(4, 4) = intTableSettings(0, 1, 3)
        avarSettings(4, 5) = "Einzelauftrag"
        avarSettings(4, 6) = True
    avarSettings(5, 0) = "lbl02"
        avarSettings(5, 1) = intTableSettings(0, 2, 0)
        avarSettings(5, 2) = intTableSettings(0, 2, 1)
        avarSettings(5, 3) = intTableSettings(0, 2, 2)
        avarSettings(5, 4) = intTableSettings(0, 2, 3)
        avarSettings(5, 5) = "test"
        avarSettings(5, 6) = False
        
    LabelSettings = avarSettings
End Function

Private Function CommandButtonSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CommandButtonSettings ausfuehren"
    End If

    Dim avarSettings(3, 6) As Variant
    avarSettings(0, 0) = "Name"
        avarSettings(0, 1) = "Left"
        avarSettings(0, 2) = "Top"
        avarSettings(0, 3) = "Width"
        avarSettings(0, 4) = "Height"
        avarSettings(0, 5) = "Caption"
        avarSettings(0, 6) = "OnClick"
    avarSettings(1, 0) = "string"
        avarSettings(1, 1) = "integer"
        avarSettings(1, 2) = "integer"
        avarSettings(1, 3) = "integer"
        avarSettings(1, 4) = "integer"
        avarSettings(1, 5) = "string"
        avarSettings(1, 6) = "string"
    avarSettings(2, 0) = "cmdExit"
        avarSettings(2, 1) = 12585
        avarSettings(2, 2) = 960
        avarSettings(2, 3) = 3120
        avarSettings(2, 4) = 330
        avarSettings(2, 5) = "Schlieﬂen"
        avarSettings(2, 6) = "=CloseFrmAngebotSuchen()"
    avarSettings(3, 0) = "cmdSearch"
        avarSettings(3, 1) = 6975
        avarSettings(3, 2) = 960
        avarSettings(3, 3) = 2730
        avarSettings(3, 4) = 330
        avarSettings(3, 5) = "Suchen"
        avarSettings(3, 6) = "=CloseFrmAngebotSuchen()"
        
    CommandButtonSettings = avarSettings
End Function

Private Function TextBoxSettings() As Variant

    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.TextBoxSettingsNeu ausfuehren"
    End If
    
    ' get table parameters
    Dim intTableSettings() As Integer
    intTableSettings = basAngebotSuchen.CreateTable(5)
    
    Dim avarSettings(7, 5) As Variant
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
        avarSettings(3, 1) = intTableSettings(1, 0, 0)
        avarSettings(3, 2) = intTableSettings(1, 0, 1)
        avarSettings(3, 3) = intTableSettings(1, 0, 2)
        avarSettings(3, 4) = intTableSettings(1, 0, 3)
        avarSettings(3, 5) = True
    avarSettings(4, 0) = "txt1"
        avarSettings(4, 1) = intTableSettings(1, 1, 0)
        avarSettings(4, 2) = intTableSettings(1, 1, 1)
        avarSettings(4, 3) = intTableSettings(1, 1, 2)
        avarSettings(4, 4) = intTableSettings(1, 1, 3)
        avarSettings(4, 5) = True
    avarSettings(5, 0) = "txt2"
        avarSettings(5, 1) = intTableSettings(1, 2, 0)
        avarSettings(5, 2) = intTableSettings(1, 2, 1)
        avarSettings(5, 3) = intTableSettings(1, 2, 2)
        avarSettings(5, 4) = intTableSettings(1, 2, 3)
        avarSettings(5, 5) = True
    avarSettings(6, 0) = "txt3"
        avarSettings(6, 1) = intTableSettings(1, 3, 0)
        avarSettings(6, 2) = intTableSettings(1, 3, 1)
        avarSettings(6, 3) = intTableSettings(1, 3, 2)
        avarSettings(6, 4) = intTableSettings(1, 3, 3)
        avarSettings(6, 5) = True
    avarSettings(7, 0) = "txt4"
        avarSettings(7, 1) = intTableSettings(1, 4, 0)
        avarSettings(7, 2) = intTableSettings(1, 4, 1)
        avarSettings(7, 3) = intTableSettings(1, 4, 2)
        avarSettings(7, 4) = intTableSettings(1, 4, 3)
        avarSettings(7, 5) = True
    
    TextBoxSettings = avarSettings
End Function

Private Function SubFormSettings() As Variant
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SubFormSettings ausfuehren"
    End If
    
    Dim avarSubFormSettings(2, 6) As Variant
    
    avarSubFormSettings(0, 0) = "Name"
        avarSubFormSettings(0, 1) = "Left"
        avarSubFormSettings(0, 2) = "Top"
        avarSubFormSettings(0, 3) = "Width"
        avarSubFormSettings(0, 4) = "Height"
        avarSubFormSettings(0, 5) = "SourceObject"
        avarSubFormSettings(0, 6) = "Locked"
    avarSubFormSettings(1, 0) = "string"
        avarSubFormSettings(1, 1) = "integer"
        avarSubFormSettings(1, 2) = "integer"
        avarSubFormSettings(1, 3) = "integer"
        avarSubFormSettings(1, 4) = "integer"
        avarSubFormSettings(1, 5) = "string"
        avarSubFormSettings(1, 6) = "booelan"
    avarSubFormSettings(2, 0) = "frbSubFrm"
        avarSubFormSettings(2, 1) = 510
        avarSubFormSettings(2, 2) = 2453
        avarSubFormSettings(2, 3) = 9218
        avarSubFormSettings(2, 4) = 5055
        avarSubFormSettings(2, 5) = "frmAngebotSuchenSub"
        avarSubFormSettings(2, 6) = True
        
    SubFormSettings = avarSubFormSettings
End Function

Public Function CloseFrmAngebotSuchen()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function
