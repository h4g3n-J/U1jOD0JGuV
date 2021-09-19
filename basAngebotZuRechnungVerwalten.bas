Attribute VB_Name = "basAngebotZuRechnungVerwalten"
Option Compare Database
Option Explicit

Public Sub BuildAngebotZuRechnungVerwalten()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute BuildAngebotZuRechnungVerwalten.BuildAngebotZuRechnungVerwalten"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    ' clear form
     basAngebotZuRechnungVerwalten.ClearForm strFormName
     
     ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' set form caption
    objForm.Caption = strFormName
    
    ' declare command button
    Dim btnButton As CommandButton
    
    ' declare label
    Dim lblLabel As Label
    
    ' declare textbox
    Dim txtTextbox As TextBox
    
    ' declare combobox
    Dim cboCombobox As ComboBox
    
    ' declare subform
    Dim frmSubForm As SubForm
    
    ' declare grid variables
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intWidth As Integer
        Dim intHeight As Integer
        
        Dim intColumn As Integer
        Dim intRow As Integer
        Dim strParent As String
        
    ' create information grid
    Dim aintInformationGrid() As Integer
            
        ' grid settings
        intNumberOfColumns = 2
        intNumberOfRows = 4
        intLeft = 10000
        intTop = 2820
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "ID"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'cbo01
    intColumn = 2
    intRow = 2
    Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
        With cboCombobox
            .Name = "cbo01"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .RowSource = "tblAngebot"
            .AllowValueListEdits = False
            .ListItemsEditForm = "frmAngebotErstellen"
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Angebot ID"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    
    'cbo02
    intColumn = 2
    intRow = 3
    Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
        With cboCombobox
            .Name = "cbo02"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .RowSource = "tblRechnung"
            .AllowValueListEdits = False
            .ListItemsEditForm = "frmRechnungErstellen"
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Rechnung ID"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Bemerkung"
            .Left = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' create lifecycle grid
    Dim aintLifecycleGrid() As Integer
    
        ' grid settings
        intNumberOfColumns = 2
        intNumberOfRows = 1
        intLeft = 510
        intTop = 1700
        intWidth = 2730
        intHeight = 330
        
        ReDim aintLifecycleGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        aintLifecycleGrid = basAngebotZuRechnungVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Rechnung erstellen" button
        intColumn = 1
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateRechnung"
                .Left = basAngebotZuRechnungVerwalten.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotZuRechnungVerwalten.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotZuRechnungVerwalten.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotZuRechnungVerwalten.GetHeight(aintLifecycleGrid, intColumn, intRow)
                ' .Caption = "Rechnung erstellen"
                .Caption = "Angebot erstellen"
                ' .OnClick = "=OpenFormAuftragErstellen_AngebotZuRechnung()"
                .OnClick = "=OpenFormAngebotErstellen_AngebotZuRechnungVerwalten()"
                .Visible = True
            End With
            
        ' create "Angebot erstellen" button
        intColumn = 2
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateAngebot"
                .Left = basAngebotZuRechnungVerwalten.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotZuRechnungVerwalten.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotZuRechnungVerwalten.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotZuRechnungVerwalten.GetHeight(aintLifecycleGrid, intColumn, intRow)
                ' .Caption = "Angebot erstellen"
                .Caption = "Rechnung erstellen"
                ' .OnClick = "=OpenFormAngebotErstellen_AngebotZuRechnung()"
                .OnClick = "=OpenFormRechnungErstellen_AngebotZuRechnungVerwalten()"
                .Visible = True
            End With
            
        ' create form title
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
            lblLabel.Name = "lblTitle"
            lblLabel.Visible = True
            lblLabel.Left = 566
            lblLabel.Top = 227
            lblLabel.Width = 9210
            lblLabel.Height = 507
            lblLabel.Caption = "Beziehung Angebot - Rechnung verwalten"
            
        ' create search box
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            txtTextbox.Name = "txtSearchBox"
            txtTextbox.Left = 510
            txtTextbox.Top = 960
            txtTextbox.Width = 6405
            txtTextbox.Height = 330
            txtTextbox.Visible = True
            
        ' create search button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSearch"
            btnButton.Left = 6975
            btnButton.Top = 960
            btnButton.Width = 2730
            btnButton.Height = 330
            btnButton.Caption = "Suchen"
            btnButton.OnClick = "=SearchAndReloadAngebotZuRechnungVerwalten()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormAngebotZuRechnungVerwalten()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 13180
            btnButton.Top = 1425
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AngebotZuRechnungVerwaltenSaveRecordset()"
            
        ' create createRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdCreateRecordset"
            btnButton.Left = 13180
            btnButton.Top = 1875
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Beziehung erstellen"
            btnButton.OnClick = "=CreateRecordsetAngebotZuRechnung()"
            
        ' create deleteRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdDeleteRecordset"
            btnButton.Left = 13180
            btnButton.Top = 2325
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Datensatz löschen"
            btnButton.OnClick = "=AngebotZuRechnungVerwaltenDeleteRecordset()"

        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            ' .Top = 2453
            .Top = 2820
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmAngebotZuRechnungVerwaltenSub"
            .Locked = True
        End With
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.BuildAngebotZuRechnungVerwalten executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.ClearForm"
    End If
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
            End If
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            Exit For
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    ' delete form
    basAngebotZuRechnungVerwalten.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAngebotZuRechnungVerwalten.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAngebotZuRechnungVerwalten.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAngebotZuRechnungVerwalten.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.CalculateGrid"
    End If
    
    Const cintHorizontalSpacing As Integer = 60
    Const cintVerticalSpacing As Integer = 60
    
    Dim aintGrid() As Integer
    ReDim aintGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    Dim intColumn As Integer
    Dim intRow As Integer
    
    For intColumn = 0 To intNumberOfColumns - 1
        For intRow = 0 To intNumberOfRows - 1
            ' left
            aintGrid(intColumn, intRow, 0) = intLeft + intColumn * (intColumnWidth + cintHorizontalSpacing)
            ' top
            aintGrid(intColumn, intRow, 1) = intTop + intRow * (intRowHeight + cintVerticalSpacing)
            ' width
            aintGrid(intColumn, intRow, 2) = intColumnWidth
            ' height
            aintGrid(intColumn, intRow, 3) = intRowHeight
        Next
    Next
    
    CalculateGrid = aintGrid
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestCalculateGrid"
    End If
    
    Dim intNumberOfRows As Integer
    Dim intNumberOfColumns As Integer
    Dim intRowHeight As Integer
    Dim intColumnWidth As Integer
    Dim intLeft As Integer
    Dim intTop As Integer
    
    intLeft = 50
    intTop = 50
    intNumberOfColumns = 2
    intNumberOfRows = 3
    intRowHeight = 100
    intColumnWidth = 50
    
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

    Const cintHorizontalSpacing As Integer = 60
    Const cintVerticalSpacing As Integer = 60
    
    Dim intErrorState As Integer
    intErrorState = 0
    
    Dim intBottom As Integer
    Dim intRight As Integer
    
    intBottom = intTop + (intNumberOfRows - 1) * (intRowHeight + cintVerticalSpacing)
    intRight = intLeft + (intNumberOfColumns - 1) * (intColumnWidth + cintHorizontalSpacing)
    
    If intRight <> aintInformationGrid(intNumberOfColumns - 1, 0, 0) Then
        intErrorState = intErrorState + 1
    End If
    
    If intBottom <> aintInformationGrid(0, intNumberOfRows - 1, 1) Then
        intErrorState = intErrorState + 2
    End If
    
    Select Case intErrorState
        Case 0
            MsgBox "basAngebotZuRechnungVerwalten.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAngebotZuRechnungVerwalten.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAngebotZuRechnungVerwalten.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAngebotZuRechnungVerwalten.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuRechnungVerwalten.GetLeft: column 0 is not available"
        MsgBox "basAngebotZuRechnungVerwalten.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Const cintHorizontalSpacing As Integer = 60
    Dim intLeftExpected As Integer
    intLeftExpected = cintLeft + (cintTestColumn - 1) * (cintHorizontalSpacing + cintColumnWidth)
    
    ' test run
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    Dim intLeftResult As Integer
    intLeftResult = basAngebotZuRechnungVerwalten.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAngebotZuRechnungVerwalten.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuRechnungVerwalten.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuRechnungVerwalten.GetTop: column 0 is not available"
        MsgBox "basAngebotZuRechnungVerwalten.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Const cintVerticalSpacing As Integer = 60
    Dim intTopExpected As Integer
    intTopExpected = cintTop + (cintTestRow - 1) * (cintVerticalSpacing + cintRowHeight)
    
    ' test run
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    Dim intTopResult As Integer
    intTopResult = basAngebotZuRechnungVerwalten.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAngebotZuRechnungVerwalten.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuRechnungVerwalten.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetHeight: column 0 is not available"
        MsgBox "basAngebotZuRechnungVerwalten.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAngebotZuRechnungVerwalten.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAngebotZuRechnungVerwalten.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuRechnungVerwalten.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetWidth: column 0 is not available"
        MsgBox "basAngebotZuRechnungVerwalten.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuRechnungVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAngebotZuRechnungVerwalten.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAngebotZuRechnungVerwalten.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuRechnungVerwalten.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadAngebotZuRechnungVerwalten()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.SearchAndReloadAngebotZuRechnungVerwalten"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search Rechnung
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    ' build query
    basAngebotZuRechnungVerwaltenSub.SearchAngebotZuRechnung varSearchTerm
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.SearchAndReloadAngebotZuRechnungVerwalten executed"
    End If
    
End Function

Public Function CloseFormAngebotZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.CloseForm executed"
    End If
    
End Function

Public Function OpenFormRechnungErstellen_AngebotZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.OpenFormRechnungErstellen_AngebotZuRechnungVerwalten"
    End If
    
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.OpenFormRechnungErstellen_AngebotZuRechnungVerwalten executed"
    End If
    
End Function

Public Function OpenFormAngebotErstellen_AngebotZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.OpenFormAngebotErstellen_AngebotZuRechnungVerwalten"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.OpenFormAngebotErstellen_AngebotZuRechnungVerwalten"
    End If
    
End Function

Public Function AngebotZuRechnungVerwaltenSaveRecordset()
    ' Error Code 1: user canceled function

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.AngebotZuRechnungVerwaltenSaveRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    ' declare subform name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strFieldName01 As String
    strFieldName01 = "ID"
    
    ' set recordset origin
    Dim varRecordsetID As Variant
    varRecordsetID = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strFieldName01)
    
    ' check if a record was selected
    If NotARecordSelected(varRecordsetID) Then
        Exit Function
    End If
    
    ' check for forbidden values
    Dim varValue01 As Variant
    varValue01 = Forms.Item(strFormName)!cbo01
    
    Dim varValue02 As Variant
    varValue02 = Forms.Item(strFormName)!cbo02
    
    If IsForbiddenValue(varValue01, varValue02) Then
        Exit Function
    End If
    
    ' declare class
    Dim AngebotZuRechnung As clsAngebotZuRechnung
    Set AngebotZuRechnung = New clsAngebotZuRechnung
    
    ' select recordset
    AngebotZuRechnung.SelectRecordset varRecordsetID
    
    ' allocate values to recordset properties
    With AngebotZuRechnung
        .RefBWIkey = Forms.Item(strFormName).Controls("cbo01")
        .RefRechnungNr = Forms.Item(strFormName).Controls("cbo02")
        .Bemerkung = Forms.Item(strFormName).Controls("txt03")
    End With
    
    ' consent request
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Datensatz '" & varRecordsetID & "' wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
        If varUserInput = 1 Then
            ' save changes
            AngebotZuRechnung.SaveRecordset
            MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
        Else
            Debug.Print "Error: basAngebotZuRechnungVerwalten.AngebotZuRechnungVerwaltenSaveRecordset aborted, Error Code 1"
            MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
        End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.AngebotZuRechnungVerwaltenSaveRecordset executed"
    End If
    
End Function

Public Function AngebotZuRechnungVerwaltenDeleteRecordset()

    ' Error Code 1: user aborted function
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.AngebotZuRechnungVerwaltenDeleteRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare field name
    Dim strFieldName01 As String
    strFieldName01 = "ID"
    
    Dim varRecordsetID As Variant
    varRecordsetID = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strFieldName01)
    
    ' check if a record was selected
    If NotARecordSelected(varRecordsetID) Then
        Exit Function
    End If
        
    ' initiate class AuftragZuAngebot
    Dim AngebotZuRechnung As clsAngebotZuRechnung
    Set AngebotZuRechnung = New clsAngebotZuRechnung
    
    ' select recordset
    AngebotZuRechnung.SelectRecordset varRecordsetID
    
    ' consent request
    Dim varUserInput As Variant
    varUserInput = MsgBox("Soll der Datensatz " & varRecordsetID & " wirklich gelöscht werden?", vbOKCancel, "Datensatz löschen")
    
        If varUserInput = 1 Then
            ' delete recordset
            AngebotZuRechnung.DeleteRecordset
            MsgBox "Datensatz gelöscht", vbInformation, "Datensatz löschen"
        Else
            Debug.Print "Error: basAngebotZuRechnungVerwalten.AuftragSuchenDeleteRecordset aborted, Error Code 2"
            MsgBox "löschen abgebrochen", vbInformation, "Datensatz löschen"
        End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.AngebotZuRechnungVerwaltenSaveRecordset execute"
    End If
    
End Function

Public Function CreateRecordsetAngebotZuRechnung()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.RecordsetAuftragZuAngebotErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
       
    Dim varInput01 As Variant
    varInput01 = Forms.Item(strFormName)!cbo01
    
    Dim varInput02 As Variant
    varInput02 = Forms.Item(strFormName)!cbo02
    
    ' check for missing input
    If InputIsMissing(varInput01, varInput02) Then
        Exit Function
    End If
    
    ' check for forbidden values
    If IsForbiddenValue(varInput01, varInput02) Then
        Exit Function
    End If
    
    ' create recordset
    Dim rstRecordset01 As clsAngebotZuRechnung
    Set rstRecordset01 = New clsAngebotZuRechnung
    
    ' transfer values
    With Forms.Item(strFormName)
        rstRecordset01.RefBWIkey = !cbo01
        rstRecordset01.RefRechnungNr = !cbo02
        rstRecordset01.Bemerkung = !txt03
    End With
    
    ' create Recordset
    rstRecordset01.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.RecordsetAuftragZuAngebotErstellen execute"
    End If

End Function

Private Function IsForbiddenValue(ByVal varInput01 As Variant, ByVal varInput02 As Variant)
' Error Code 1: input is not on the value list

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.IsForbiddenValue"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False
    
    ' name of table01
    Dim strDomainName01 As String
    strDomainName01 = "tblAngebot"
    
    ' name of field01 in table01
    Dim strFieldName01 As String
    strFieldName01 = "BWIKey"
    
    ' alias for field01 in error prompt
    Dim strFieldAlias01 As String
    strFieldAlias01 = "Angebot ID"
    
    ' name of table02
    Dim strDomainName02 As String
    strDomainName02 = "tblRechnung"
    
    ' name of field02
    Dim strFieldName02 As String
    strFieldName02 = "RechnungNr"
    
    ' alias for field02 in error prompt
    Dim strFieldAlias02 As String
    strFieldAlias02 = "Rechnung Nr"

    ' declare error variable
    Dim strErrorMessage As String

    ' check table01
    If DCount("[" & strFieldName01 & "]", strDomainName01, "[" & strFieldName01 & "] Like '" & varInput01 & "'") = 0 Then
        strErrorMessage = "Bitte wählen Sie im Feld " & strFieldAlias01 & "' ausschließlich Werte aus der Drop-Down-Liste." & vbCrLf
    End If
    
    ' check table02
    If DCount("[" & strFieldName02 & "]", strDomainName02, "[" & strFieldName02 & "] Like '" & varInput02 & "'") = 0 Then
        strErrorMessage = strErrorMessage & "Bitte wählen Sie im Feld " & strFieldAlias02 & "' ausschließlich Werte aus der Drop-Down-Liste." & vbCrLf
    End If
    
    ' error prompt
    If strErrorMessage <> "" Then
        MsgBox strErrorMessage, vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotZuRechnungVerwalten.IsForbiddenValue, Error Code 1"
        bolStatus = True
    End If
    
    IsForbiddenValue = bolStatus

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.IsForbiddenValue executed"
    End If
    
End Function

Private Function InputIsMissing(ByVal varInput01 As Variant, ByVal varInput02 As Variant)
' Error Code 1: mandatory value is missing

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.IsForbiddenValue"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False

    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuRechnungVerwalten"
    
    ' declare name of field01
    Dim strFieldAlias01 As String
    strFieldAlias01 = "Rechnung Nr"
    
    ' declare name of field02
    Dim strFieldAlias02 As String
    strFieldAlias02 = "Angebot ID"
    
    ' declare error variable
    Dim strErrorMessage As String
    
    ' check values
    If IsNull(varInput01) Then
        strErrorMessage = "Sie haben im Pflichtfeld '" & strFieldAlias01 & "' keinen Wert eingegeben." & vbCrLf
    End If
    
    If IsNull(varInput02) Then
        strErrorMessage = strErrorMessage & "Sie haben im Pflichtfeld '" & strFieldAlias02 & "' keinen Wert eingegeben."
    End If
    
    If strErrorMessage <> "" Then
        MsgBox strErrorMessage, vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        bolStatus = True
    End If
    
    InputIsMissing = bolStatus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.IsForbiddenValue executed"
    End If
    
End Function

Private Function NotARecordSelected(ByVal varInput As Variant) As Boolean
' Error Code 1: no recordset selected

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuRechnungVerwalten.NotARecordSelected"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False

    ' check primary key value
    If IsNull(varInput) Then
        Debug.Print "Error: basAngebotZuRechnungVerwalten.NotARecordSelected aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt!", vbCritical, "Fehler"
        bolStatus = True
    End If
    
    NotARecordSelected = bolStatus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuRechnungVerwalten.NotARecordSelected executed"
    End If
    
End Function


