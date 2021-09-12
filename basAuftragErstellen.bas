Attribute VB_Name = "basAuftragErstellen"
Option Compare Database
Option Explicit

Public Sub buildAuftragErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.buildAuftragErstellen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    ' clear form
     basAuftragErstellen.ClearForm strFormName
     
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
        intNumberOfRows = 6
        intLeft = 566
        intTop = 960
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basAuftragErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "AftrID"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "AftrTitel"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "BeginnSoll"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl04
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "EndeSoll"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 2
    intRow = 6
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl05
    intColumn = 1
    intRow = 6
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "Erstellt"
            .Left = basAuftragErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
        
        aintLifecycleGrid = basAuftragErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Einzelauftrag erstellen" button
        intColumn = 1
        intRow = 1
        
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateOffer"
                .Left = basAuftragErstellen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragErstellen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragErstellen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragErstellen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Einzelauftrag erstellen"
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
            lblLabel.Caption = "Auftrag Erstellen"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 7653
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schlieﬂen"
            btnButton.OnClick = "=CloseFormAuftragErstellen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AuftragErstellenCreateRecordset()"
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.buildAuftragErstellen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.ClearForm"
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
        Debug.Print "basAuftragErstellen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    ' delete form
    basAuftragErstellen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAuftragErstellen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAuftragErstellen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAuftragErstellen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.CalculateGrid"
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
        Debug.Print "basAuftragErstellen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestCalculateGrid"
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
    
    aintInformationGrid = basAuftragErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basAuftragErstellen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAuftragErstellen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAuftragErstellen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAuftragErstellen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErstellen.GetLeft: column 0 is not available"
        MsgBox "basAuftragErstellen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basAuftragErstellen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAuftragErstellen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErstellen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErstellen.GetTop: column 0 is not available"
        MsgBox "basAuftragErstellen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basAuftragErstellen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAuftragErstellen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErstellen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErstellen.TestGetHeight: column 0 is not available"
        MsgBox "basAuftragErstellen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAuftragErstellen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAuftragErstellen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErstellen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErstellen.TestGetWidth: column 0 is not available"
        MsgBox "basAuftragErstellen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAuftragErstellen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAuftragErstellen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErstellen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadAuftragErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.SearchAndReloadAuftragErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search Rechnung
    Dim strQueryName As String
    strQueryName = "qryAuftragErstellen"
    
    Dim strQuerySource As String
    strQuerySource = "tblEinzelauftrag"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "EAkurzKey"
    
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    basAuftragErstellenSub.SearchEinzelauftrag strQueryName, strQuerySource, strPrimaryKey, varSearchTerm
    
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.SearchAndReloadAuftragErstellen executed"
    End If
    
End Function

Public Function CloseFormAuftragErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.CloseForm executed"
    End If
    
End Function

Public Function AuftragErstellenCreateRecordset()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErstellen.AuftragErstellenCreateRecordset"
    End If
    
    Dim strTableName As String
    strTableName = "tblAuftrag"
    
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    Dim rstRecordset As clsAuftrag
    Set rstRecordset = New clsAuftrag
    
    ' transfer values from form to clsAuftrag
    With Forms.Item(strFormName)
        rstRecordset.AftrID = !txt00
        rstRecordset.AftrTitel = !txt01
        rstRecordset.Bemerkung = !txt02
        rstRecordset.BeginnSoll = !txt03
        rstRecordset.EndeSoll = !txt04
        rstRecordset.Erstellt = !txt05
    End With
    
    ' create Recordset
    rstRecordset.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErstellen.AuftragErstellenCreateRecordset executed"
    End If
    
End Function



