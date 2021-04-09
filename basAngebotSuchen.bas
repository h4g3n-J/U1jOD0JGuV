Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

' returns array
' (column, row, property)
' properties: 0 - Left, 1 - Top, 2 - Width, 3 - Height
' calculates left, top, width and height parameters
Private Function CreateTable(ByVal intNumberOfColumns As Integer, ByRef intColumnWidth() As Integer, ByVal intNumberOfRows As Integer)
    
    intNumberOfColumns = intNumberOfColumns - 1
    intNumberOfRows = intNumberOfRows - 1
    
    ' column dimension
    Const cintCol00Left As Integer = 10000
    Const cintHorizontalSpacing As Integer = 60
            
    ' row dimension
    Const cintRow00Top As Integer = 2430
    Dim intRowHeight As Integer
    intRowHeight = 330
    
    Const cintVerticalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    ' avarSettings(0, 0) = "Left"
    ' avarSettings(0, 1) = "Top"
    ' avarSettings(0, 2) = "Width"
    ' avarSettings(0, 3) = "Height"
        
    Dim inti As Integer
    Dim intj As Integer
    
    Dim aintTableSettings() As Integer
    ReDim aintTableSettings(intNumberOfColumns, intNumberOfRows, cintNumberOfProperties)
    
    For inti = 0 To intNumberOfColumns
        ' For intr = 0 To cintNumberOfRows
        For intj = 0 To intNumberOfRows
            ' set column left
            aintTableSettings(inti, intj, 0) = cintCol00Left + inti * (intColumnWidth(inti) + cintHorizontalSpacing)
            ' set row top
            aintTableSettings(inti, intj, 1) = cintRow00Top + intj * (intRowHeight + cintVerticalSpacing)
            ' set column width
            aintTableSettings(inti, intj, 2) = intColumnWidth(inti)
            ' set row height
            aintTableSettings(inti, intj, 3) = intRowHeight
        Next
    Next
        
    ReDim Preserve aintTableSettings(intNumberOfColumns, intNumberOfRows, cintNumberOfProperties)

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
    
    ' create table
    Dim intColumnWidth(1) As Integer
    intColumnWidth(0) = 2540
    intColumnWidth(1) = 3120
    
    Dim intTableSettings() As Integer
    intTableSettings = CreateTable(2, intColumnWidth, 6)
    
    ' create textboxes
    basAngebotSuchen.CreateTextbox strTempFormName, intTableSettings
    
    ' create labels
    basAngebotSuchen.CreateLabel strTempFormName, intTableSettings
    
    ' create command buttons
    basAngebotSuchen.CreateCommandButton strTempFormName, intTableSettings
    
    ' create subform
    basAngebotSuchen.CreateSubForm strTempFormName, intTableSettings
        
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngbotSuchen: " & strFormName & " erstellt"
    End If

End Sub

Private Sub CreateTextbox(ByVal strFormName As String, intTableSettings() As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateTextbox ausfuehren"
    End If
    
    ' declare textbox
    Dim txtTextbox As Textbox
    
    ' create search input box
    Set txtTextbox = CreateControl(strFormName, acTextBox, acDetail)
        txtTextbox.Name = "txtSearch"
        txtTextbox.Left = 510
        txtTextbox.Top = 960
        txtTextbox.Width = 6405
        txtTextbox.Height = 330
        txtTextbox.Visible = True
    
    Dim avarSettingsTable(4, 3) As Variant
    
    avarSettingsTable(0, 0) = "txt0"
        avarSettingsTable(0, 1) = 1
        avarSettingsTable(0, 2) = 0
        avarSettingsTable(0, 3) = True
    avarSettingsTable(1, 0) = "txt1"
        avarSettingsTable(1, 1) = 1
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
    avarSettingsTable(2, 0) = "txt2"
        avarSettingsTable(2, 1) = 1
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
    avarSettingsTable(3, 0) = "txt3"
        avarSettingsTable(3, 1) = 1
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = True
    avarSettingsTable(4, 0) = "txt4"
        avarSettingsTable(4, 1) = 1
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = True
        
        ' intTableSettings = basAngebotSuchen.CreateTable(5) ' create / calculate table parameters
        Dim intColumn As Integer
        Dim intRow As Integer
        
        Dim inti As Integer   ' column
        For inti = LBound(avarSettingsTable, 1) To UBound(avarSettingsTable, 1)
            Set txtTextbox = CreateControl(strFormName, acTextBox, acDetail)
            txtTextbox.Name = avarSettingsTable(inti, 0) ' set name
            txtTextbox.Visible = avarSettingsTable(inti, 3) ' set visibility
            
            intColumn = avarSettingsTable(inti, 1)
            intRow = avarSettingsTable(inti, 2)
            ' txtTextbox = TextboxPosition(txtTextbox, intTableSettings, avarSettingsTable(inti, 1), avarSettingsTable(inti, 2)) ' set position
            Set txtTextbox = PositionObjectInTable(txtTextbox, intTableSettings, intColumn, intRow) ' set position
        Next

End Sub

Private Sub CreateLabel(ByVal strFormName As String, intTableSettings() As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateLabel ausfuehren"
    End If
    
    ' declare label
    Dim lblLabel As Label
    
    ' create form title
    Set lblLabel = CreateControl(strFormName, acLabel, acDetail)
        lblLabel.Name = "lblTitle"
        lblLabel.Visible = True
        lblLabel.Left = 566
        lblLabel.Top = 227
        lblLabel.Width = 9210
        lblLabel.Height = 507
        lblLabel.Caption = "Angebot Suchen"
    
    Dim avarSettingsTable(4, 3) As Variant
    avarSettingsTable(0, 0) = "lbl00"
        avarSettingsTable(0, 1) = 0
        avarSettingsTable(0, 2) = 0
        avarSettingsTable(0, 3) = True
        ' avarSettings(3, 5) = "Angebot"
    avarSettingsTable(1, 0) = "lbl01"
        avarSettingsTable(1, 1) = 0
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
        ' avarSettings(4, 5) = "Einzelauftrag"
    avarSettingsTable(2, 0) = "lbl02"
        avarSettingsTable(2, 1) = 0
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
    avarSettingsTable(3, 0) = "lbl03"
        avarSettingsTable(3, 1) = 0
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = True
    avarSettingsTable(4, 0) = "lbl04"
        avarSettingsTable(4, 1) = 0
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = True
    
    Dim intColumn As Integer
    Dim intRow As Integer
        
    Dim inti As Integer   ' column
    For inti = LBound(avarSettingsTable, 1) To UBound(avarSettingsTable, 1)
        Set lblLabel = CreateControl(strFormName, acLabel, acDetail)
        lblLabel.Name = avarSettingsTable(inti, 0) ' set name
        lblLabel.Visible = avarSettingsTable(inti, 3) ' set visibility
        
        intColumn = avarSettingsTable(inti, 1)
        intRow = avarSettingsTable(inti, 2)
        Set lblLabel = PositionObjectInTable(lblLabel, intTableSettings, intColumn, intRow) ' set position
    Next
    
End Sub

Private Sub CreateCommandButton(ByVal strFormName As String, intTableSettings() As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateCommandButton ausfuehren"
    End If
    
    Dim btnButton As CommandButton
    Set btnButton = CreateControl(strFormName, acCommandButton, acDetail)
    
    btnButton.Name = "cmdExit"
        btnButton.Left = 12585
        btnButton.Top = 960
        btnButton.Width = 3120
        btnButton.Height = 330
        btnButton.Caption = "Schlieﬂen"
        btnButton.OnClick = "=CloseFrmAngebotSuchen()"
        
    Set btnButton = CreateControl(strFormName, acCommandButton, acDetail)
    btnButton.Name = "cmdSearch"
        btnButton.Left = 6975
        btnButton.Top = 960
        btnButton.Width = 2730
        btnButton.Height = 330
        btnButton.Caption = "Suchen"
        btnButton.OnClick = "=CloseFrmAngebotSuchen()"
        
End Sub

Private Sub CreateSubForm(ByVal strFormName As String, intTableSettings() As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateSubForm ausfuehren"
    End If
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SubFormSettings ausfuehren"
    End If
    
    Dim frmSubForm As SubForm
    Set frmSubForm = CreateControl(strFormName, acSubform, acDetail)
    
    frmSubForm.Name = "frbSubFrm"
        frmSubForm.Left = 510
        frmSubForm.Top = 2453
        frmSubForm.Width = 9218
        frmSubForm.Height = 5055
        frmSubForm.SourceObject = "frmAngebotSuchenSub"
        frmSubForm.Locked = True

End Sub

Private Function PositionObjectInTable(ByVal objObject As Object, aintTableSetting() As Integer, intColumn As Integer, intRow As Integer) As Object
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.TextboxPosition ausfuehren"
    End If
    
    If Not (TypeOf objObject Is Textbox Or TypeOf objObject Is Label Or TypeOf objObject Is CommandButton) Then
        Debug.Print "basAngebotSuchen.TextboxPosition: falscher Objekttyp uebergeben, Funktion abgebrochen"
        Exit Function
    End If
    
    objObject.Left = aintTableSetting(intColumn, intRow, 0)
    objObject.Top = aintTableSetting(intColumn, intRow, 1)
    objObject.Width = aintTableSetting(intColumn, intRow, 2)
    objObject.Height = aintTableSetting(intColumn, intRow, 3)
    
    Set PositionObjectInTable = objObject
End Function

Public Function CloseFrmAngebotSuchen()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function
