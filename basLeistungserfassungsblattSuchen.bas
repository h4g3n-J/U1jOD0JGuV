Attribute VB_Name = "basLeistungserfassungsblattSuchen"
Option Compare Database
Option Explicit

Public Sub BuildLeistungserfassungsblattSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.BuildLeistungserfassungsblattSuchen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattSuchen"
    
    ' clear form
     basLeistungserfassungsblattSuchen.ClearForm strFormName
     
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
        intNumberOfRows = 8
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basLeistungserfassungsblattSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "RechnungNr"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Bemerkung"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Rechnung (Link)"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Technisch Richtig Datum"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl04
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Ist Teilrechnung"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 2
    intRow = 6
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl05
    intColumn = 1
    intRow = 6
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "Ist Schlussrechnung"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 2
    intRow = 7
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl06
    intColumn = 1
    intRow = 7
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "Kalkulation LNW (Link)"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 2
    intRow = 8
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl07
    intColumn = 1
    intRow = 8
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "Rechnung Brutto"
            .Left = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    ' column added? -> update intNumberOfColumns
    
    ' create lifecycle grid
    Dim aintLifecycleGrid() As Integer
    
        ' grid settings
        intNumberOfColumns = 1
        intNumberOfRows = 1
        intLeft = 510
        intTop = 1700
        intWidth = 2730
        intHeight = 330
        
        ReDim aintLifecycleGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        aintLifecycleGrid = basLeistungserfassungsblattSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Rechnung erstellen" button
        intColumn = 1
        intRow = 1
        
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateOffer"
                .Left = basLeistungserfassungsblattSuchen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basLeistungserfassungsblattSuchen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basLeistungserfassungsblattSuchen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basLeistungserfassungsblattSuchen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Rechnung erstellen"
' insert editing here ----> .OnClick = "=OpenFormCreateOffer()"
                .Visible = False
            End With
            
        ' create form title
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
            lblLabel.Name = "lblTitle"
            lblLabel.Visible = True
            lblLabel.Left = 566
            lblLabel.Top = 227
            lblLabel.Width = 9210
            lblLabel.Height = 507
            lblLabel.Caption = "Leistungserfassungblatt Suchen"
            
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
            btnButton.OnClick = "=SearchAndReloadLeistungserfassungsblattSuchen()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schlie�en"
            btnButton.OnClick = "=CloseFormLeistungserfassungsblattSuchen()"

        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmLeistungserfassungsblattSuchenSub"
            .Locked = True
        End With
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.BuildLeistungserfassungsblattSuchen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.ClearForm"
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
        Debug.Print "basLeistungserfassungsblattSuchen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattSuchen"
    
    ' delete form
    basLeistungserfassungsblattSuchen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basLeistungserfassungsblattSuchen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basLeistungserfassungsblattSuchen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basLeistungserfassungsblattSuchen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.CalculateGrid"
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
        Debug.Print "basLeistungserfassungsblattSuchen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.TestCalculateGrid"
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
    
    aintInformationGrid = basLeistungserfassungsblattSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basLeistungserfassungsblattSuchen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basLeistungserfassungsblattSuchen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basLeistungserfassungsblattSuchen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "basLeistungserfassungsblattSuchen.TestCalculateGrid"
        Case 3
            MsgBox "basLeistungserfassungsblattSuchen.TestCalculateGrid: Test feiled: Error Code 3", vbCritical, "basLeistungserfassungsblattSuchen.TestCalculateGrid"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchen.GetLeft: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basLeistungserfassungsblattSuchen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basLeistungserfassungsblattSuchen.TestGetLeft: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchen.GetTop: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basLeistungserfassungsblattSuchen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basLeistungserfassungsblattSuchen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchen.TestGetTop executed"
    End If
    
End Sub
