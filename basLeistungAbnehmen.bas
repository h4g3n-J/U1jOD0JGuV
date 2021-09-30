Attribute VB_Name = "basLeistungAbnehmen"
Option Compare Database
Option Explicit

Public gvarLeistungAbnehmenClipboardAftrID As Variant
Public gvarLeistungAbnehmenClipboardBWIKey As Variant

Public Sub buildLeistungAbnehmen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.buildLeistungAbnehmen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungAbnehmen"
    
    ' clear form
     basLeistungAbnehmen.ClearForm strFormName
     
     ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' set form caption
    objForm.Caption = strFormName
    
    ' set On Open event
    objForm.OnOpen = "=OnOpenFrmLeistungAbnehmen()"
    
    ' set On Close event
    objForm.OnClose = "=OnCloseFrmLeistungAbnehmen()"
    
    ' declare command button
    Dim btnButton As CommandButton
    
    ' declare label
    Dim lblLabel As Label
    
    ' declare textbox
    Dim txtTextbox As TextBox
    
    ' declare combobox
    Dim cboCombobox As ComboBox
    
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
        intNumberOfRows = 10
        intLeft = 566
        intTop = 960
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basLeistungAbnehmen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
        ' txt00
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl00
        intColumn = 1
        intRow = 3
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
            With lblLabel
                .Name = "lbl00"
                .Caption = "Angebot ID"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
        
        ' txt01
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt01"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 1
            End With
            
        ' lbl01
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "Abgenommen am*"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt02
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
                .BorderStyle = 0
            End With
            
        ' lbl02
        intColumn = 1
        intRow = 4
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
            With lblLabel
                .Name = "lbl02"
                .Caption = "Mengengerüst"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt03
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
                .BorderStyle = 0
            End With
            
        ' lbl03
        intColumn = 1
        intRow = 5
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
            With lblLabel
                .Name = "lbl03"
                .Caption = "Leistungsbeschreibung"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt04
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl04
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
            With lblLabel
                .Name = "lbl04"
                .Caption = "Bemerkung"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt05
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt05"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl05
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "Beauftragt am"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt06
        ' intColumn = 2
        ' intRow = 7
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt06"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
            ' End With
            
        ' ' lbl06
        ' intColumn = 1
        ' intRow = 7
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
            ' With lblLabel
                ' .Name = "lbl06"
                ' .Caption = "Abgebrochen"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt07
        ' intColumn = 2
        ' intRow = 8
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt07"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
            ' End With
            
        ' ' lbl07
        ' intColumn = 1
        ' intRow = 8
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
            ' With lblLabel
                ' .Name = "lbl07"
                ' .Caption = "Angeboten"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt08
        ' intColumn = 2
        ' intRow = 9
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt08"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
            ' End With
            
        ' ' lbl08
        ' intColumn = 1
        ' intRow = 9
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
            ' With lblLabel
                ' .Name = "lbl08"
                ' .Caption = "Abgenommen"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt09
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl09
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Auftrag Beginn"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt10
        intColumn = 2
        intRow = 7
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl10
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Auftrag Ende"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt11
        ' intColumn = 2
        ' intRow = 12
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt11"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
            ' End With
            
        ' ' lbl11
        ' intColumn = 1
        ' intRow = 12
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
            ' With lblLabel
                ' .Name = "lbl11"
                ' .Caption = "Storniert"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt12
        ' intColumn = 2
        ' intRow = 7
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt12"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
                ' .BorderStyle = 0
            ' End With
            
        ' ' lbl12
        ' intColumn = 1
        ' intRow = 7
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
            ' With lblLabel
                ' .Name = "lbl12"
                ' .Caption = "Preis Brutto"
                ' .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt13
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt13"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl13
        intColumn = 1
        intRow = 1
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt13")
            With lblLabel
                .Name = "lbl13"
                .Caption = "Ticket ID"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt14
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt14"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl14
        intColumn = 1
        intRow = 2
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt14")
            With lblLabel
                .Name = "lbl14"
                .Caption = "Zusammenfassung"
                .Left = basLeistungAbnehmen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basLeistungAbnehmen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basLeistungAbnehmen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basLeistungAbnehmen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            lblLabel.Caption = "Auftrag erteilen"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 7653
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormLeistungAbnehmen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=LeistungAbnehmenSaveRecordset()"
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.buildLeistungAbnehmen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.ClearForm"
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
        Debug.Print "basLeistungAbnehmen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungAbnehmen"
    
    ' delete form
    basLeistungAbnehmen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basLeistungAbnehmen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basLeistungAbnehmen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basLeistungAbnehmen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.CalculateGrid"
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
        Debug.Print "basLeistungAbnehmen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestCalculateGrid"
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
    
    aintInformationGrid = basLeistungAbnehmen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basLeistungAbnehmen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basLeistungAbnehmen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basLeistungAbnehmen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basLeistungAbnehmen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungAbnehmen.GetLeft: column 0 is not available"
        MsgBox "basLeistungAbnehmen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungAbnehmen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basLeistungAbnehmen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basLeistungAbnehmen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungAbnehmen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungAbnehmen.GetTop: column 0 is not available"
        MsgBox "basLeistungAbnehmen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungAbnehmen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basLeistungAbnehmen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basLeistungAbnehmen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungAbnehmen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungAbnehmen.TestGetHeight: column 0 is not available"
        MsgBox "basLeistungAbnehmen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungAbnehmen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basLeistungAbnehmen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basLeistungAbnehmen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungAbnehmen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungAbnehmen.TestGetWidth: column 0 is not available"
        MsgBox "basLeistungAbnehmen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungAbnehmen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basLeistungAbnehmen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basLeistungAbnehmen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungAbnehmen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.TestGetWidth executed"
    End If
    
End Sub

Public Function LeistungAbnehmenSaveRecordset()
' Error Code 1: no value assgined to BWIKey
' Error Code 2: a recordset of that name already exists
' Error Code 3: input value is not on the value list

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.LeistungAbnehmenSaveRecordset"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungAbnehmen"
    
    Dim varBeauftragtDatum As Variant
    varBeauftragtDatum = Forms.Item(strFormName)!txt01
        
    ' check if varBeauftragDatum is empty
    If IsNull(varBeauftragtDatum) Then
        MsgBox "Sie haben im Pflichtfeld 'Abgenommen an' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basLeistungAbnehmen.LeistungAbnehmenSaveRecordset, Error Code 1"
        Exit Function
    End If
               
    ' transfer values to clsAngebot
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    
    ' select recordset
    rstAngebot.SelectRecordset gvarLeistungAbnehmenClipboardBWIKey
    
    With Forms.Item(strFormName)
        rstAngebot.Bemerkung = !txt04
        rstAngebot.AbgenommenDatum = !txt01
    End With
    
    ' save recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Angebot '" & gvarLeistungAbnehmenClipboardBWIKey & "' wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
    If varUserInput = 1 Then
        rstAngebot.SaveRecordset
        MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
    Else
        Debug.Print "Error: basLeistungAbnehmen.AuftragSuchenSaveRecordset aborted, Error Code 2"
        MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.LeistungAbnehmenSaveRecordset executed"
    End If
    
ExitProc:
    Set rstAngebot = Nothing
End Function

Public Function OnCloseFrmLeistungAbnehmen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.OnCloseFrmLeistungAbnehmen"
    End If
    
    gvarLeistungAbnehmenClipboardAftrID = Null
    gvarLeistungAbnehmenClipboardBWIKey = Null
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.OnCloseFrmLeistungAbnehmen executed"
    End If
    
End Function

Public Function CloseFormLeistungAbnehmen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.CloseFormLeistungAbnehmen"
    End If
    
    gvarLeistungAbnehmenClipboardAftrID = Null
    gvarLeistungAbnehmenClipboardBWIKey = Null
    
    Dim strFormName As String
    strFormName = "frmLeistungAbnehmen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.CloseFormLeistungAbnehmen executed"
    End If
    
End Function

Public Function OnOpenFrmLeistungAbnehmen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungAbnehmen.OnOpenfrmLeistungAbnehmen"
    End If
    
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
    rstTicket.SelectRecordset (gvarLeistungAbnehmenClipboardAftrID)
    
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    rstAngebot.SelectRecordset (gvarLeistungAbnehmenClipboardBWIKey)
    
    Forms!frmLeistungAbnehmen.Form!txt13 = rstTicket.AftrID
    Forms!frmLeistungAbnehmen.Form!txt14 = rstTicket.AftrTitel
    Forms!frmLeistungAbnehmen.Form!txt00 = rstAngebot.BWIKey
    Forms!frmLeistungAbnehmen.Form!txt02 = rstAngebot.MengengeruestLink
    Forms!frmLeistungAbnehmen.Form!txt03 = rstAngebot.LeistungsbeschreibungLink
    Forms!frmLeistungAbnehmen.Form!txt09 = rstAngebot.AftrBeginn
    Forms!frmLeistungAbnehmen.Form!txt10 = rstAngebot.AftrEnde
    Forms!frmLeistungAbnehmen.Form!txt05 = rstAngebot.BeauftragtDatum
    Forms!frmLeistungAbnehmen.Form!txt01 = rstAngebot.AbgenommenDatum
    Forms!frmLeistungAbnehmen.Form!txt04 = rstAngebot.Bemerkung
    Forms!frmLeistungAbnehmen.Form!txt01 = rstAngebot.AbgenommenDatum
    
    ' set focus
    Forms!frmLeistungAbnehmen!txt01.SetFocus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungAbnehmen.OnOpenfrmLeistungAbnehmen executed"
    End If
    
End Function



