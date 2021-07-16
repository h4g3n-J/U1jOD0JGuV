Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

' build form AngebotSuchen
Public Sub BuildAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen ausführen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' declare temporary form name
    Dim strTempFormName As String

    ' clear form
    basAngebotSuchen.ClearForm strFormName

    ' declare form
    Dim frm As Form
    Set frm = CreateForm
    
    ' write temporary form name to strFormName
    strTempFormName = frm.Name
    
    ' set form caption
    frm.Caption = strFormName
    
    ' create information grid
        ' set top left position
        Dim intLeft As Integer
        intLeft = 10000
        
        ' set top position
        Dim intTop As Integer
        intTop = 2430
        
        ' set column width
        Dim intColumnWidth(1) As Integer
        intColumnWidth(0) = 2540
        intColumnWidth(1) = 3120
    
        ' set number of rows
        Dim intNumberOfRows As Integer
        intNumberOfRows = 6
        
        Dim aintInformationGrid() As Integer
        aintInformationGrid = basAngebotSuchen.CalculateInformationGrid(2, intColumnWidth, intNumberOfRows, intLeft, intTop)
    
        ' create textboxes
        basAngebotSuchen.CreateTextbox strTempFormName, aintInformationGrid, intNumberOfRows
        
        ' create labels
        basAngebotSuchen.CreateLabel strTempFormName, aintInformationGrid, intNumberOfRows
    
        ' create captions
        Dim astrCaptionSettings() As String
        astrCaptionSettings = basAngebotSuchen.CaptionAndValueSettings(intNumberOfRows) ' get caption settings
        basAngebotSuchen.SetLabelCaption strTempFormName, astrCaptionSettings, intNumberOfRows ' set caption
    
    ' create command buttons
    basAngebotSuchen.CreateCommandButton strTempFormName, aintInformationGrid
    
    ' create subform
    basAngebotSuchen.CreateSubForm strTempFormName, aintInformationGrid
        
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen: " & strFormName & " erstellt"
    End If

End Sub

' create textbox
Private Sub CreateTextbox(ByVal strFormName As String, aintTableSettings() As Integer, ByVal intNumberOfRows As Integer)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateTextbox ausfuehren"
    End If
    
    ' declare textbox
    Dim txtTextbox As Textbox
    
    ' create search box
    Set txtTextbox = CreateControl(strFormName, acTextBox, acDetail)
        txtTextbox.Name = "txtSearchBox"
        txtTextbox.Left = 510
        txtTextbox.Top = 960
        txtTextbox.Width = 6405
        txtTextbox.Height = 330
        txtTextbox.Visible = True
    
    intNumberOfRows = intNumberOfRows - 1
    
    Dim avarSettingsTable() As Variant
    ReDim avarSettingsTable(intNumberOfRows, 4)
    
    ' set default values
    avarSettingsTable(0, 0) = "txt00" ' name
        avarSettingsTable(0, 1) = 1 ' column
        avarSettingsTable(0, 2) = 0 ' row
        avarSettingsTable(0, 3) = True ' visibility
        avarSettingsTable(0, 4) = False ' isHyperlink
    avarSettingsTable(1, 0) = "txt01"
        avarSettingsTable(1, 1) = 1
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
        avarSettingsTable(1, 4) = False
    avarSettingsTable(2, 0) = "txt02"
        avarSettingsTable(2, 1) = 1
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
        avarSettingsTable(2, 4) = True
    avarSettingsTable(3, 0) = "txt03"
        avarSettingsTable(3, 1) = 1
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = True
        avarSettingsTable(3, 4) = True
    avarSettingsTable(4, 0) = "txt04"
        avarSettingsTable(4, 1) = 1
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = True
        avarSettingsTable(4, 4) = False
    avarSettingsTable(5, 0) = "txt05"
        avarSettingsTable(5, 1) = 1
        avarSettingsTable(5, 2) = 5
        avarSettingsTable(5, 3) = True
        avarSettingsTable(5, 4) = False
        
        Dim intColumn As Integer
        Dim intRow As Integer
        
        Dim inti As Integer
        For inti = LBound(avarSettingsTable, 1) To intNumberOfRows
        
            ' create textbox
            Set txtTextbox = CreateControl(strFormName, acTextBox, acDetail)
            txtTextbox.Name = avarSettingsTable(inti, 0) ' set name
            txtTextbox.Visible = avarSettingsTable(inti, 3) ' set visibility
            txtTextbox.IsHyperlink = avarSettingsTable(inti, 4) ' set IsHyperlink
            
            ' set position
            intColumn = avarSettingsTable(inti, 1)
            intRow = avarSettingsTable(inti, 2)
            Set txtTextbox = basSupport.PositionObjectInTable(txtTextbox, aintTableSettings, intColumn, intRow)
            
        Next

End Sub

' create label
Private Sub CreateLabel(ByVal strFormName As String, ByRef intTableSettings() As Integer, ByVal intNumberOfRows As Integer)
    
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
    
    intNumberOfRows = intNumberOfRows - 1
    
    Dim avarSettingsTable() As Variant
    ReDim avarSettingsTable(intNumberOfRows, 4)
    
    ' set default values
    avarSettingsTable(0, 0) = "lbl00"
        avarSettingsTable(0, 1) = 0
        avarSettingsTable(0, 2) = 0
        avarSettingsTable(0, 3) = True
        avarSettingsTable(0, 4) = "txt00"
    avarSettingsTable(1, 0) = "lbl01"
        avarSettingsTable(1, 1) = 0
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
        avarSettingsTable(1, 4) = "txt01"
    avarSettingsTable(2, 0) = "lbl02"
        avarSettingsTable(2, 1) = 0
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
        avarSettingsTable(2, 4) = "txt02"
    avarSettingsTable(3, 0) = "lbl03"
        avarSettingsTable(3, 1) = 0
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = True
        avarSettingsTable(3, 4) = "txt03"
    avarSettingsTable(4, 0) = "lbl04"
        avarSettingsTable(4, 1) = 0
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = True
        avarSettingsTable(4, 4) = "txt04"
    avarSettingsTable(5, 0) = "lbl05"
        avarSettingsTable(5, 1) = 0
        avarSettingsTable(5, 2) = 5
        avarSettingsTable(5, 3) = True
        avarSettingsTable(5, 4) = "txt05"
    
    Dim intColumn As Integer
    Dim intRow As Integer
        
    Dim inti As Integer   ' column
    For inti = LBound(avarSettingsTable, 1) To intNumberOfRows
        Set lblLabel = CreateControl(strFormName, acLabel, acDetail, avarSettingsTable(inti, 4))
        lblLabel.Name = avarSettingsTable(inti, 0) ' set name
        lblLabel.Visible = avarSettingsTable(inti, 3) ' set visibility
        
        intColumn = avarSettingsTable(inti, 1)
        intRow = avarSettingsTable(inti, 2)
        Set lblLabel = basSupport.PositionObjectInTable(lblLabel, intTableSettings, intColumn, intRow) ' set position
    Next
    
End Sub

' create command buttons
Private Sub CreateCommandButton(ByVal strFormName As String, ByRef intTableSettings() As Integer)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateCommandButton"
    End If
    
    ' declare CommandButton
    Dim btnButton As CommandButton
    
    ' create exit button
    Set btnButton = CreateControl(strFormName, acCommandButton, acDetail)
    btnButton.Name = "cmdExit"
        btnButton.Left = 13180
        btnButton.Top = 960
        btnButton.Width = 3120
        btnButton.Height = 330
        btnButton.Caption = "Schließen"
        btnButton.OnClick = "=CloseFrmAngebotSuchen()"
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateCommandButton: exit button created"
    End If
        
    ' create search button
    Set btnButton = CreateControl(strFormName, acCommandButton, acDetail)
    btnButton.Name = "cmdSearch"
        btnButton.Left = 6975
        btnButton.Top = 960
        btnButton.Width = 2730
        btnButton.Height = 330
        btnButton.Caption = "Suchen"
        btnButton.OnClick = "=SearchAngebot()"
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateCommandButton: search button created"
    End If
        
    ' create control grid
        ' calculate positions
        Dim intNumberOfColumns As Integer
        intNumberOfColumns = 1
        
        Dim intColumnWidth As Integer
        intColumnWidth = 2730
        
        Dim intLeft As Integer
        intLeft = 510
        
        Dim intTop As Integer
        intTop = 1700
        
        Dim intRowHeight As Integer
        intRowHeight = 330
        
        Dim aintPositions As Integer
        ' replace with 'LifecycleGrid' function
        aintPositions = basSupport.CalculateLifecycleBar(intNumberOfColumns, intColumnWidth, intLeft, intTop, intRowHeight)
        
        ' create CreateAngebot button
            Set btnButton = CreateControl(strFormName, acCommandButton, acDetail)
            btnButton.Name = "cmdCreateOffer"
                btnButton.Left = 6975
                btnButton.Top = 960
                btnButton.Width = 2730
                btnButton.Height = 330
                btnButton.Caption = "Angebot erstellen"
                btnButton.OnClick = "=OpenCreateOffer()"
                
            ' set createAngebot button
                ' insert code here
        
        ' event message
        If gconVerbatim Then
            Debug.Print "basAngebotSuchen.CreateCommandButton successful"
        End If
        
End Sub

' create subform
Private Sub CreateSubForm(ByVal strFormName As String, ByRef intTableSettings() As Integer)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateSubForm ausfuehren"
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

Public Function CloseFrmAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function

Public Function SearchAngebot()

    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SearchAngebot ausfuehren"
    End If
    
    ' search term
    basBuild.BuildQryAngebotAuswahl Application.Forms.Item("frmAngebotSuchen").Controls("txtSearchBox")
    
    ' close form
    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
    
    ' open form
    DoCmd.OpenForm "frmAngebotSuchen", acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SearchAngebot executed"
    End If
    
End Function

Private Sub SetLabelCaption(ByVal strFormName As String, ByRef astrCaptionSettings() As String, ByVal intNumberOfRows As Integer)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SetLabelCaption ausfuehren"
    End If
        
    Dim inti As Integer
    For inti = LBound(astrCaptionSettings, 1) + 1 To intNumberOfRows
        Forms(strFormName).Controls(astrCaptionSettings(inti, 0)).Caption = astrCaptionSettings(inti, 1)
    Next
    
End Sub

' set captions and values
Public Function CaptionAndValueSettings(ByVal intNumberOfRows As Integer) As String()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CaptionAndValue ausfuehren"
    End If
    
    Dim astrSettings() As String
    ReDim astrSettings(intNumberOfRows, 3)

    astrSettings(0, 0) = "label.Name"
        astrSettings(0, 1) = "label.Caption"
        astrSettings(0, 2) = "txtbox.Name"
        astrSettings(0, 3) = "txtbox"
    astrSettings(1, 0) = "lbl00"
        astrSettings(1, 1) = "Angebot"
        astrSettings(1, 2) = "txt00"
        astrSettings(1, 3) = "BWIKey"
    astrSettings(2, 0) = "lbl01"
        astrSettings(2, 1) = "Einzelauftrag"
        astrSettings(2, 2) = "txt01"
        astrSettings(2, 3) = "EAkurzKey"
    astrSettings(3, 0) = "lbl02"
        astrSettings(3, 1) = "Mengengerüst"
        astrSettings(3, 2) = "txt02"
        astrSettings(3, 3) = "MengengeruestLink"
    astrSettings(4, 0) = "lbl03"
        astrSettings(4, 1) = "Leistungsbeschreibung"
        astrSettings(4, 2) = "txt03"
        astrSettings(4, 3) = "LeistungsbeschreibungLink"
    astrSettings(5, 0) = "lbl04"
        astrSettings(5, 1) = "Bemerkung"
        astrSettings(5, 2) = "txt04"
        astrSettings(5, 3) = "Bemerkung"
    astrSettings(6, 0) = "lbl05"
        astrSettings(6, 1) = "wildcard"
        astrSettings(6, 2) = "txt05"
        astrSettings(6, 3) = "Bemerkung"
    
    CaptionAndValueSettings = astrSettings
End Function

' returns array
' (column, row, property)
' properties: 0 - Left, 1 - Top, 2 - Width, 3 - Height
' calculates left, top, width and height parameters
Private Function CalculateInformationGrid(ByVal intNumberOfColumns As Integer, ByRef aintColumnWidth() As Integer, ByVal intNumberOfRows As Integer, Optional ByVal intLeft As Integer = 10000, Optional ByVal intTop As Integer = 2430)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CalculateTableSetting ausfuehren"
    End If
    
    intNumberOfColumns = intNumberOfColumns - 1
    intNumberOfRows = intNumberOfRows - 1
    
    ' column dimension
    Const cintHorizontalSpacing As Integer = 60
            
    ' row dimension
    Dim intRowHeight As Integer
    intRowHeight = 330
    
    Const cintVerticalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    Dim aintGridSettings() As Integer
    ReDim aintGridSettings(intNumberOfColumns, intNumberOfRows, cintNumberOfProperties)
    
    ' compute cell position properties
    Dim inti As Integer
    Dim intj As Integer
    For inti = 0 To intNumberOfColumns
        ' For intr = 0 To cintNumberOfRows
        For intj = 0 To intNumberOfRows
            ' set column left
            aintGridSettings(inti, intj, 0) = intLeft + inti * (aintColumnWidth(inti) + cintHorizontalSpacing)
            ' set row top
            aintGridSettings(inti, intj, 1) = intTop + intj * (intRowHeight + cintVerticalSpacing)
            ' set column width
            aintGridSettings(inti, intj, 2) = aintColumnWidth(inti)
            ' set row height
            aintGridSettings(inti, intj, 3) = intRowHeight
        Next
    Next

    CalculateInformationGrid = aintGridSettings
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CalculateInformationGrid ausgefuehrt"
    End If

End Function

Public Function OpenCreateOffer()

    Dim strFormName As String
    strFormName = "frmCreateOffer"

    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.OpenCreateOffer ausfuehren"
    End If
    
    ' close form
    ' DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.OpenCreateOffer ausgefuehrt"
    End If
    
End Function

' delete form
' 1. check if form exists
' 2. close if form is loaded
' 3. delete form
Public Sub ClearForm(ByVal strFormName As String)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClearForm ausfuehren"
    End If
    
    Dim objDummy As Object
    For Each objDummy In Application.CurrentProject.AllForms
        If objDummy.Name = strFormName Then
            
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
                
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basAngebotSuchen.ClearForm: " & strFormName & " ist geoeffnet, Formular schließen"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "basAngebotSuchen.ClearForm: " & strFormName & " existiert bereits, Formular loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
End Sub

' intNumberOfColumns: defines the number of columns
' aintColumnWidth: array, defines the width of each column
' intLeft: top left position
' intTop: top position
' intRowHeight: row height
' returns array: (i, 0) Left, (i, 1) Top, (i, 2) Width, (i, 3) Height
Public Function CalculateLifecycleGrid()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.CalculateLifecycleGrid"
    End If

    Dim intNumberOfColumns As Integer
    intNumberOfColumns = 2
    Const cintColumnWidth As Integer = 2730
    Const cintLeft As Integer = 100
    Const cintTop As Integer = 2430
    Const cintRowHeight As Integer = 330
    Const cintHorizontalSpacing As Integer = 60
    
    ' compute cell position properties
    intNumberOfColumns = intNumberOfColumns - 1 ' adjusted for counting
    Const cintNumberOfProperties = 3
    Dim aintBarSettings() As Integer
    ReDim aintBarSettings(intNumberOfColumns, cintNumberOfProperties)
    
    Dim inti As Integer
    For inti = 0 To intNumberOfColumns
            ' set column left
            aintBarSettings(inti, 0) = cintLeft + inti * (cintColumnWidth + cintHorizontalSpacing)
            ' set row top
            aintBarSettings(inti, 1) = cintTop
            ' set column width
            aintBarSettings(inti, 2) = cintColumnWidth
            ' set row height
            aintBarSettings(inti, 3) = cintRowHeight
    Next

    CalculateLifecycleGrid = aintBarSettings
    
End Function

' get left from grid
Public Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuche.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuche.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    intColumn = intColumn - 1
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetLeft executed"
    End If
    
    GetLeft = aintGrid(intColumn, 0)
End Function

' get top from grid
Public Function GetTop(ByRef aintGrid As Variant, ByVal intColumn As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuche.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuche.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    intColumn = intColumn - 1
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetTop executed"
    End If
    
    GetTop = aintGrid(intColumn, 1)
End Function

' get width from grid
Public Function GetWidth(ByRef aintGrid As Variant, ByVal intColumn As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuche.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuche.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    intColumn = intColumn - 1
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetWidth executed"
    End If
    
    GetWidth = aintGrid(intColumn, 2)
End Function

' get height from grid
Public Function GetHeight(ByRef aintGrid As Variant, ByVal intColumn As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuche.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuche.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    intColumn = intColumn - 1
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetHeigth executed"
    End If
    
    GetHeight = aintGrid(intColumn, 3)
End Function

