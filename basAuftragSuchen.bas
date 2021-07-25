Attribute VB_Name = "basAuftragSuchen"
Option Compare Database
Option Explicit

Public Sub BuildAuftragSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.BuildAuftragSuchen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
    ' clear form
     basAuftragSuchen.ClearForm strFormName
     
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
        intNumberOfRows = 11
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basAuftragSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
        'txt00
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
                
        'txt01
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt01"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl01
        intColumn = 1
        intRow = 2
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "AftrTitel"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt02
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl02
        intColumn = 1
        intRow = 3
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
            With lblLabel
                .Name = "lbl02"
                .Caption = "StatusKey"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt03
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl03
        intColumn = 1
        intRow = 4
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
            With lblLabel
                .Name = "lbl03"
                .Caption = "OwnerKey"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt04
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl04
        intColumn = 1
        intRow = 5
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
            With lblLabel
                .Name = "lbl04"
                .Caption = "PrioritaetKey"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt05
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt05"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl05
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "ParentKey"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt06
        intColumn = 2
        intRow = 7
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt06"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl06
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
            With lblLabel
                .Name = "lbl06"
                .Caption = "Bemerkung"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt07
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt07"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl07
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
            With lblLabel
                .Name = "lbl07"
                .Caption = "BeginnSoll"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt08
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt08"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl08
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
            With lblLabel
                .Name = "lbl08"
                .Caption = "EndeSoll"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt09
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl09
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Erstellt"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        'txt10
        intColumn = 2
        intRow = 11
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        'lbl10
        intColumn = 1
        intRow = 11
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Kunde"
                .Left = basAuftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
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
        
        aintLifecycleGrid = basAuftragSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Angebot erstellen" button
        intColumn = 1
        intRow = 1
        
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateOffer"
                .Left = basAuftragSuchen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragSuchen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragSuchen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragSuchen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Auftrag erstellen"
                .OnClick = "=OpenFormCreateOffer()"
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
            lblLabel.Caption = "Auftrag Suchen"
            
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
            btnButton.OnClick = "=SearchAngebot()"
            btnButton.Visible = False
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schlieﬂen"
            btnButton.OnClick = "=CloseFrmAuftragSuchen()"
            
        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmAuftragSuchenSub"
            .Locked = True
        End With
        
        ' close form
        DoCmd.Close acForm, strTempFormName, acSaveYes
    
        ' rename form
        DoCmd.Rename strFormName, acForm, strTempFormName
        
        ' event message
        If gconVerbatim Then
            Debug.Print "basAuftragSuchen.BuildAuftragSuchen: " & strFormName & " erstellt"
        End If
            
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.ClearForm"
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
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
    basAuftragSuchen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " created."
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " exists."
            Exit For
        End If
    Next
    
    basAuftragSuchen.ClearForm strFormName
    
    Debug.Print "basAuftragSuchen.TestClearForm: ClearForm executed."
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " exists."
            Exit For
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.TestClearForm"
    End If
    
End Sub

Public Function CloseFrmAuftragSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.CloseForm executed"
    End If
    
End Function

' get left from grid
Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchen.GetLeft: column 0 is not available"
        MsgBox "basAuftragSuchen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.GetLeft executed"
    End If
    
End Function

' get left from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchen.GetTop: column 0 is not available"
        MsgBox "basAuftragSuchen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.GetTop executed"
    End If
    
End Function

' get left from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchen.GetWidth: column 0 is not available"
        MsgBox "basAuftragSuchen.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.GetWidth executed"
    End If
    
End Function

' get left from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchen.GetHeight: column 0 is not available"
        MsgBox "basAuftragSuchen.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.GetHeight executed"
    End If
    
End Function

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.CalculateGrid"
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
        Debug.Print "basAuftragSuchen.CalculateGrid executed"
    End If
    
End Function
