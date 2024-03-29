Attribute VB_Name = "basEinzelauftragUebersicht"
Option Compare Database
Option Explicit

Public Sub BuildEinzelauftragUebersicht()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.BuildEinzelauftragUebersicht"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragUebersicht"
    
    ' set subform name
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' clear form
    basEinzelauftragUebersicht.ClearForm strFormName
     
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
        intNumberOfRows = 2
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
    ' calculate domain sum 'angeboten' & 'abgerechnet'
    Dim strAngebotSumme As String
    strAngebotSumme = "= Nz(DSum('AngebotBrutto', 'qryEinzelauftragUebersicht'))"
    
    Dim strAbgerechnetSumme As String
    strAbgerechnetSumme = "= Nz(DSum('RechnungBrutto', 'qryEinzelauftragUebersicht'))"
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = strAngebotSumme
            .Left = basEinzelauftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "Angeboten (Summe, Brutto)"
            .Left = basEinzelauftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = strAbgerechnetSumme
            .Left = basEinzelauftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Abgerechnet (Summe, brutto)"
            .Left = basEinzelauftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
                
        ' create form title
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
            lblLabel.Name = "lblTitle"
            lblLabel.Visible = True
            lblLabel.Left = 566
            lblLabel.Top = 227
            lblLabel.Width = 9210
            lblLabel.Height = 507
            lblLabel.Caption = "Einzelauftrag �bersicht"
            
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
            btnButton.OnClick = "=SearchAndReloadEinzelauftragUebersicht()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schlie�en"
            btnButton.OnClick = "=CloseFormEinzelauftragUebersicht()"
            
        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = strSubformName
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmEinzelauftragUebersichtSub"
            .Locked = True
        End With
            
    ' close form and save changes
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.BuildEinzelauftragUebersicht executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.ClearForm"
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
        Debug.Print "basEinzelauftragUebersicht.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragUebersicht"
    
    ' delete form
    basEinzelauftragUebersicht.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basEinzelauftragUebersicht.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basEinzelauftragUebersicht.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basEinzelauftragUebersicht.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.CalculateGrid"
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
        Debug.Print "basEinzelauftragUebersicht.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestCalculateGrid"
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
    
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basEinzelauftragUebersicht.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basEinzelauftragUebersicht.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basEinzelauftragUebersicht.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basEinzelauftragUebersicht.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragUebersicht.GetLeft: column 0 is not available"
        MsgBox "basEinzelauftragUebersicht.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basEinzelauftragUebersicht.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basEinzelauftragUebersicht.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragUebersicht.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragUebersicht.GetTop: column 0 is not available"
        MsgBox "basEinzelauftragUebersicht.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basEinzelauftragUebersicht.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basEinzelauftragUebersicht.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragUebersicht.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragUebersicht.TestGetHeight: column 0 is not available"
        MsgBox "basEinzelauftragUebersicht.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basEinzelauftragUebersicht.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basEinzelauftragUebersicht.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragUebersicht.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragUebersicht.TestGetWidth: column 0 is not available"
        MsgBox "basEinzelauftragUebersicht.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basEinzelauftragUebersicht.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basEinzelauftragUebersicht.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragUebersicht.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadEinzelauftragUebersicht()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.SearchAndReloadEinzelauftragUebersicht"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragUebersicht"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search EinzelauftragUebersicht
    Dim strQueryName As String
    strQueryName = "qryEinzelauftragUebersicht"
    
    Dim strQuerySource As String
    ' -> strQuerySource = "tblEinzelauftragUebersicht"
    
    Dim strPrimaryKey As String
    ' -> strPrimaryKey = "EAkurzKey"
    
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    ' -> basEinzelauftragUebersichtSub.SearchEinzelauftragUebersicht strQueryName, strQuerySource, strPrimaryKey, varSearchTerm
    basEinzelauftragUebersichtSub.SearchEinzelauftragUebersicht strQueryName, varSearchTerm
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.SearchAndReloadEinzelauftragUebersicht executed"
    End If
    
End Function

Public Function CloseFormEinzelauftragUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragUebersicht.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragUebersicht"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragUebersicht.CloseForm executed"
    End If
    
End Function
