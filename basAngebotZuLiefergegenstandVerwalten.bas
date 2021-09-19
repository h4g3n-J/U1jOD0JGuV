Attribute VB_Name = "basAngebotZuLiefergegenstandVerwalten"
Option Compare Database
Option Explicit

Public Sub BuildAngebotZuLiefergegenstandVerwalten()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.BuildAngebotZuLiefergegenstandVerwalten"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
    ' clear form
     basAngebotZuLiefergegenstandVerwalten.ClearForm strFormName
     
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
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'cbo01
    intColumn = 2
    intRow = 2
    Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
        With cboCombobox
            .Name = "cbo01"
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    
    'cbo02
    intColumn = 2
    intRow = 3
    Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
        With cboCombobox
            .Name = "cbo02"
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .RowSource = "tblLiefergegenstand"
            .AllowValueListEdits = False
            .ListItemsEditForm = "frmLiefergegenstandErstellen"
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Liefergegenstand ID"
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, intColumn, intRow)
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
        
        aintLifecycleGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Angebot erstellen" button
        intColumn = 1
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateAngebot"
                .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Angebot erstellen"
                .OnClick = "=OpendFormAngebotErstellen_AngebotZuLiefergegenstandVerwalten()"
                .Visible = True
            End With
            
        ' create "Liefergegenstand erstellen" button
        intColumn = 2
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateLiefergegenstand"
                .Left = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotZuLiefergegenstandVerwalten.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Liefergegenstand erstellen"
                .OnClick = "=OpenFormLiefergegenstandErstellen_AngebotZuLiefergegenstandVerwalten()"
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
            lblLabel.Caption = "Beziehung Angebot - Liefergegenstand verwalten"
            
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
            btnButton.OnClick = "=SearchAndReloadAngebotZuLiefergegenstandVerwalten()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormAngebotZuLiefergegenstandVerwalten()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 13180
            btnButton.Top = 1425
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AngebotZuLiefergegenstandVerwaltenSaveRecordset()"
            
        ' create createRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdCreateRecordset"
            btnButton.Left = 13180
            btnButton.Top = 1875
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Beziehung erstellen"
            btnButton.OnClick = "=CreateRecordsetAngebotZuLiefergegenstand()"
            
        ' create deleteRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdDeleteRecordset"
            btnButton.Left = 13180
            btnButton.Top = 2325
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Datensatz löschen"
            btnButton.OnClick = "=AngebotZuLiefergegenstandVerwaltenDeleteRecordset()"

        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            ' .Top = 2453
            .Top = 2820
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmAngebotZuLiefergegenstandVerwaltenSub"
            .Locked = True
        End With
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.BuildAngebotZuLiefergegenstandVerwalten executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.ClearForm"
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
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
    ' delete form
    basAngebotZuLiefergegenstandVerwalten.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAngebotZuLiefergegenstandVerwalten.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAngebotZuLiefergegenstandVerwalten.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAngebotZuLiefergegenstandVerwalten.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.CalculateGrid"
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
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid"
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
    
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.GetLeft: column 0 is not available"
        MsgBox "basAngebotZuLiefergegenstandVerwalten.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basAngebotZuLiefergegenstandVerwalten.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.GetTop: column 0 is not available"
        MsgBox "basAngebotZuLiefergegenstandVerwalten.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basAngebotZuLiefergegenstandVerwalten.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetHeight: column 0 is not available"
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAngebotZuLiefergegenstandVerwalten.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetWidth: column 0 is not available"
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotZuLiefergegenstandVerwalten.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAngebotZuLiefergegenstandVerwalten.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotZuLiefergegenstandVerwalten.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadAngebotZuLiefergegenstandVerwalten()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.SearchAndReloadAngebotZuLiefergegenstandVerwalten"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search Rechnung
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    ' build query
    basAngebotZuLiefergegenstandVerwaltenSub.SearchAngebotZuLiefergegenstand varSearchTerm
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.SearchAndReloadAngebotZuLiefergegenstandVerwalten executed"
    End If
    
End Function

Public Function CloseFormAngebotZuLiefergegenstandVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.CloseForm executed"
    End If
    
End Function

Public Function OpendFormAngebotErstellen_AngebotZuLiefergegenstandVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.OpendFormAngebotErstellen_AngebotZuLiefergegenstand"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.OpendFormAngebotErstellen_AngebotZuLiefergegenstand executed"
    End If
    
End Function

Public Function OpenFormLiefergegenstandErstellen_AngebotZuLiefergegenstandVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.OpenFormLiefergegenstandErstellen_AuftragZuAngebot"
    End If
    
    Dim strFormName As String
    strFormName = "frmLiefergegenstandErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.OpenFormLiefergegenstandErstellen_AuftragZuAngebot executed"
    End If
    
End Function

Public Function AngebotZuLiefergegenstandVerwaltenSaveRecordset()
    ' Error Code 1: user canceled function

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.AngebotZuLiefergegenstandVerwaltenSaveRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
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
    Dim AngebotZuLiefergegenstand As clsAngebotZuLiefergegenstand
    Set AngebotZuLiefergegenstand = New clsAngebotZuLiefergegenstand
    
    ' select recordset
    AngebotZuLiefergegenstand.SelectRecordset varRecordsetID
    
    ' allocate values to recordset properties
    With AngebotZuLiefergegenstand
        .RefBWIkey = Forms.Item(strFormName).Controls("cbo01")
        .RefLiefergegenstandID = Forms.Item(strFormName).Controls("cbo02")
        .Bemerkung = Forms.Item(strFormName).Controls("txt03")
    End With
    
    ' consent request
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Datensatz '" & varRecordsetID & "' wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
        If varUserInput = 1 Then
            ' save changes
            AngebotZuLiefergegenstand.SaveRecordset
            MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
        Else
            Debug.Print "Error: basAngebotZuLiefergegenstandVerwalten.AngebotZuLiefergegenstandVerwaltenSaveRecordset aborted, Error Code 1"
            MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
        End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.AngebotZuLiefergegenstandVerwaltenSaveRecordset executed"
    End If
    
End Function

Public Function AngebotZuLiefergegenstandVerwaltenDeleteRecordset()

    ' Error Code 1: user aborted function
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.AngebotZuLiefergegenstandVerwaltenDeleteRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
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
    Dim AngebotZuLiefergegenstand As clsAngebotZuLiefergegenstand
    Set AngebotZuLiefergegenstand = New clsAngebotZuLiefergegenstand
    
    ' select recordset
    AngebotZuLiefergegenstand.SelectRecordset varRecordsetID
    
    ' consent request
    Dim varUserInput As Variant
    varUserInput = MsgBox("Soll der Datensatz " & varRecordsetID & " wirklich gelöscht werden?", vbOKCancel, "Datensatz löschen")
    
        If varUserInput = 1 Then
            ' delete recordset
            AngebotZuLiefergegenstand.DeleteRecordset
            MsgBox "Datensatz gelöscht", vbInformation, "Datensatz löschen"
        Else
            Debug.Print "Error: basAngebotZuLiefergegenstandVerwalten.AuftragSuchenDeleteRecordset aborted, Error Code 2"
            MsgBox "löschen abgebrochen", vbInformation, "Datensatz löschen"
        End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.AngebotZuLiefergegenstandVerwaltenSaveRecordset execute"
    End If
    
End Function

Public Function CreateRecordsetAngebotZuLiefergegenstand()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.RecordsetAuftragZuAngebotErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
       
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
    Dim rstRecordset01 As clsAngebotZuLiefergegenstand
    Set rstRecordset01 = New clsAngebotZuLiefergegenstand
    
    ' transfer values
    With Forms.Item(strFormName)
        rstRecordset01.RefBWIkey = !cbo01
        rstRecordset01.RefLiefergegenstandID = !cbo02
        rstRecordset01.Bemerkung = !txt03
    End With
    
    ' create Recordset
    rstRecordset01.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.RecordsetAuftragZuAngebotErstellen execute"
    End If

End Function

Private Function IsForbiddenValue(ByVal varInput01 As Variant, ByVal varInput02 As Variant)
' Error Code 1: input is not on the value list

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.IsForbiddenValue"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False
    
    ' name of table01
    Dim strDomainName01 As String
    strDomainName01 = "tblAngebot"
    
    ' name of field01 in table01
    Dim strFieldName01 As String
    strFieldName01 = "BWIKey"
    
    ' field01 alias
    Dim strFieldAlias01 As String
    strFieldAlias01 = "Angebot ID"
    
    ' name of table02
    Dim strDomainName02 As String
    strDomainName02 = "tblLiefergegenstand"
    
    ' name of field02
    Dim strFieldName02 As String
    strFieldName02 = "LiefergegenstandID"
    
    ' field02 alias in error prompt
    Dim strFieldAlias02 As String
    strFieldAlias02 = "Liefergegenstand ID"

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
        Debug.Print "Error: basAngebotZuLiefergegenstandVerwalten.IsForbiddenValue, Error Code 1"
        bolStatus = True
    End If
    
    IsForbiddenValue = bolStatus

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.IsForbiddenValue executed"
    End If
    
End Function

Private Function InputIsMissing(ByVal varInput01 As Variant, ByVal varInput02 As Variant)
' Error Code 1: mandatory value is missing

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.IsForbiddenValue"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False

    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotZuLiefergegenstandVerwalten"
    
    ' declare alias of field01
    Dim strFieldAlias01 As String
    strFieldAlias01 = "Angebot ID"
    
    ' declare alias of field02
    Dim strFieldAlias02 As String
    strFieldAlias02 = "Liefergegenstand ID"
    
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
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.IsForbiddenValue executed"
    End If
    
End Function

Private Function NotARecordSelected(ByVal varInput As Variant) As Boolean
' Error Code 1: no recordset selected

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotZuLiefergegenstandVerwalten.NotARecordSelected"
    End If
    
    Dim bolStatus As Boolean
    bolStatus = False

    ' check primary key value
    If IsNull(varInput) Then
        Debug.Print "Error: basAngebotZuLiefergegenstandVerwalten.NotARecordSelected aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt!", vbCritical, "Fehler"
        bolStatus = True
    End If
    
    NotARecordSelected = bolStatus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotZuLiefergegenstandVerwalten.NotARecordSelected executed"
    End If
    
End Function
