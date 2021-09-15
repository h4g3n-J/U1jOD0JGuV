Attribute VB_Name = "basLeistungserfassungsblattErstellen"
Option Compare Database
Option Explicit

Public Sub buildLeistungserfassungsblattErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.buildLeistungserfassungsblattErstellen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    ' clear form
     basLeistungserfassungsblattErstellen.ClearForm strFormName
     
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
        intNumberOfRows = 13
        intLeft = 566
        intTop = 960
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "Leistungserfassungsblatt"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "RechnungNr"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Bemerkung"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "BelegID"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl04
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Brutto"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            lblLabel.Caption = "Leistungserfassungsblatt erfassen"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 7653
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schlieﬂen"
            btnButton.OnClick = "=CloseFormLeistungserfassungsblattErstellen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=LeistungserfassungsblattErstellenCreateRecordset()"
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.buildLeistungserfassungsblattErstellen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.ClearForm"
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
        Debug.Print "basLeistungserfassungsblattErstellen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    ' delete form
    basLeistungserfassungsblattErstellen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basLeistungserfassungsblattErstellen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basLeistungserfassungsblattErstellen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basLeistungserfassungsblattErstellen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.CalculateGrid"
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
        Debug.Print "basLeistungserfassungsblattErstellen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestCalculateGrid"
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
    
    aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basLeistungserfassungsblattErstellen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basLeistungserfassungsblattErstellen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basLeistungserfassungsblattErstellen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basLeistungserfassungsblattErstellen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattErstellen.GetLeft: column 0 is not available"
        MsgBox "basLeistungserfassungsblattErstellen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basLeistungserfassungsblattErstellen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattErstellen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattErstellen.GetTop: column 0 is not available"
        MsgBox "basLeistungserfassungsblattErstellen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basLeistungserfassungsblattErstellen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattErstellen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetHeight: column 0 is not available"
        MsgBox "basLeistungserfassungsblattErstellen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basLeistungserfassungsblattErstellen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattErstellen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetWidth: column 0 is not available"
        MsgBox "basLeistungserfassungsblattErstellen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basLeistungserfassungsblattErstellen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattErstellen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.TestGetWidth executed"
    End If
    
End Sub

Public Function CloseFormLeistungserfassungsblattErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.CloseForm executed"
    End If
    
End Function

Public Function LeistungserfassungsblattErstellenCreateRecordset()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.LeistungserfassungsblattErstellenCreateRecordset"
    End If
    
    Dim strTableName As String
    strTableName = "tblLeistungserfassungsblatt"
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    ' transfer values from form to clsLeistungserfassungsblatt
    With Forms.Item(strFormName)
        rstRecordset.Leistungserfassungsblatt = !txt00
        rstRecordset.RechnungNr = !txt01
        rstRecordset.Bemerkung = !txt02
        rstRecordset.BelegID = !txt03
        rstRecordset.Brutto = !txt04
    End With
    
    ' create Recordset
    rstRecordset.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.LeistungserfassungsblattErstellenCreateRecordset executed"
    End If
    
End Function


