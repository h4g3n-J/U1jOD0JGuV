Attribute VB_Name = "basAuftragErteilen"
Option Compare Database
Option Explicit

Public gvarAuftragErteilenClipboardAftrID As Variant
Public gvarAuftragErteilenClipboardBWIKey As Variant
Public gvarAuftragErteilenClipboardEinzelauftrag As Variant

Public Sub buildAuftragErteilen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.buildAuftragErteilen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragErteilen"
    
    ' clear form
     basAuftragErteilen.ClearForm strFormName
     
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
    objForm.OnOpen = "=OnOpenFrmAuftragErteilen()"
    
    ' set On Close event
    objForm.OnClose = "=OnCloseFrmAuftragErteilen()"
    
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
        intNumberOfRows = 15
        intLeft = 566
        intTop = 960
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basAuftragErteilen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
        ' txt00
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
        
        ' txt01
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt01"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl01
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "Einzelauftrag"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt02
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt03
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt04
        intColumn = 2
        intRow = 11
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl04
        intColumn = 1
        intRow = 11
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
            With lblLabel
                .Name = "lbl04"
                .Caption = "Bemerkung"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt05
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt05"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl05
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "Beauftragt am*"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt06
        ' intColumn = 2
        ' intRow = 7
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt06"
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt07
        ' intColumn = 2
        ' intRow = 8
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt07"
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt08
        ' intColumn = 2
        ' intRow = 9
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt08"
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt09
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 1
            End With
            
        ' lbl09
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Auftrag Beginn*"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt10
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 1
            End With
            
        ' lbl10
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Auftrag Ende*"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt11
        ' intColumn = 2
        ' intRow = 12
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt11"
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt12
        intColumn = 2
        intRow = 7
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt12"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl12
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
            With lblLabel
                .Name = "lbl12"
                .Caption = "Preis Brutto"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt13
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt13"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt14
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt14"
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAuftragErteilen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAuftragErteilen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAuftragErteilen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAuftragErteilen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            btnButton.OnClick = "=CloseFormAuftragErteilen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AuftragErteilenCreateRecordset()"
            
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.buildAuftragErteilen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.ClearForm"
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
        Debug.Print "basAuftragErteilen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErteilen"
    
    ' delete form
    basAuftragErteilen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAuftragErteilen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAuftragErteilen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAuftragErteilen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.CalculateGrid"
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
        Debug.Print "basAuftragErteilen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestCalculateGrid"
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
    
    aintInformationGrid = basAuftragErteilen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basAuftragErteilen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAuftragErteilen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAuftragErteilen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAuftragErteilen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErteilen.GetLeft: column 0 is not available"
        MsgBox "basAuftragErteilen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErteilen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basAuftragErteilen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAuftragErteilen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErteilen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErteilen.GetTop: column 0 is not available"
        MsgBox "basAuftragErteilen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErteilen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basAuftragErteilen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAuftragErteilen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErteilen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErteilen.TestGetHeight: column 0 is not available"
        MsgBox "basAuftragErteilen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErteilen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAuftragErteilen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAuftragErteilen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErteilen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragErteilen.TestGetWidth: column 0 is not available"
        MsgBox "basAuftragErteilen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragErteilen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAuftragErteilen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAuftragErteilen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragErteilen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.TestGetWidth executed"
    End If
    
End Sub

Public Function AuftragErteilenCreateRecordset()
' Error Code 1: no value assgined to BWIKey
' Error Code 2: a recordset of that name already exists
' Error Code 3: input value is not on the value list

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.AuftragErteilenCreateRecordset"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErteilen"
    
    Dim varAuftragBeginn As Variant
    varAuftragBeginn = Forms.Item(strFormName)!txt09
    
    Dim varAuftragEnde As Variant
    varAuftragEnde = Forms.Item(strFormName)!txt10
    
    Dim varBeauftragtDatum As Variant
    varBeauftragtDatum = Forms.Item(strFormName)!txt05
    
    ' check if varAuftragBeginn is empty
    If IsNull(varAuftragBeginn) Then
        MsgBox "Sie haben im Pflichtfeld 'Auftrag Beginn' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAuftragErteilen.AuftragErteilenCreateRecordset, Error Code 1"
        Exit Function
    End If
    
    ' check if varAuftragEnde is empty
    If IsNull(varAuftragEnde) Then
        MsgBox "Sie haben im Pflichtfeld 'Auftrag Ende' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAuftragErteilen.AuftragErteilenCreateRecordset, Error Code 1"
        Exit Function
    End If
    
    ' check if varBeauftragDatum is empty
    If IsNull(varBeauftragtDatum) Then
        MsgBox "Sie haben im Pflichtfeld 'Beauftragt' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAuftragErteilen.AuftragErteilenCreateRecordset, Error Code 1"
        Exit Function
    End If
               
    ' transfer values to clsAngebot
    Dim rstRecordset01 As clsAngebot
    Set rstRecordset01 = New clsAngebot
    
    ' select recordset
    rstRecordset01.SelectRecordset gvarAuftragErteilenClipboardBWIKey
    
    With Forms.Item(strFormName)
        rstRecordset01.AftrBeginn = !txt09
        rstRecordset01.AftrEnde = !txt10
        rstRecordset01.Bemerkung = !txt04
        rstRecordset01.BeauftragtDatum = !txt05
    End With
    
    ' save recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Angebot '" & gvarAuftragErteilenClipboardBWIKey & "' wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
    If varUserInput = 1 Then
        rstRecordset01.SaveRecordset
        MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
    Else
        Debug.Print "Error: basAuftragErteilen.AuftragSuchenSaveRecordset aborted, Error Code 2"
        MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.AuftragErteilenCreateRecordset executed"
    End If
    
ExitProc:
    Set rstRecordset01 = Nothing
End Function

Public Function OnCloseFrmAuftragErteilen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.OnCloseFrmAuftragErteilen"
    End If
    
    gvarAuftragErteilenClipboardAftrID = Null
    gvarAuftragErteilenClipboardBWIKey = Null
    gvarAuftragErteilenClipboardEinzelauftrag = Null
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.OnCloseFrmAuftragErteilen executed"
    End If
    
End Function

Public Function CloseFormAuftragErteilen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragErteilen.CloseFormAuftragErteilen"
    End If
    
    gvarAuftragErteilenClipboardAftrID = Null
    gvarAuftragErteilenClipboardBWIKey = Null
    gvarAuftragErteilenClipboardEinzelauftrag = Null
    
    Dim strFormName As String
    strFormName = "frmAuftragErteilen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragErteilen.CloseFormAuftragErteilen executed"
    End If
    
End Function

Public Function OnOpenFrmAuftragErteilen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.OnOpenfrmAuftragErteilen"
    End If
    
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
    rstTicket.SelectRecordset (gvarAuftragErteilenClipboardAftrID)
    
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    rstAngebot.SelectRecordset (gvarAuftragErteilenClipboardBWIKey)
    
    Forms!frmAuftragErteilen.Form!txt13 = rstTicket.AftrID
    Forms!frmAuftragErteilen.Form!txt14 = rstTicket.AftrTitel
    Forms!frmAuftragErteilen.Form!txt00 = rstAngebot.BWIKey
    Forms!frmAuftragErteilen.Form!txt02 = rstAngebot.MengengeruestLink
    Forms!frmAuftragErteilen.Form!txt03 = rstAngebot.LeistungsbeschreibungLink
    Forms!frmAuftragErteilen.Form!txt01 = gvarAuftragErteilenClipboardEinzelauftrag
    Forms!frmAuftragErteilen.Form!txt12 = rstAngebot.AngebotBrutto
    Forms!frmAuftragErteilen.Form!txt09 = rstAngebot.AftrBeginn
    Forms!frmAuftragErteilen.Form!txt10 = rstAngebot.AftrEnde
    Forms!frmAuftragErteilen.Form!txt05 = rstAngebot.BeauftragtDatum
    Forms!frmAuftragErteilen.Form!txt04 = rstAngebot.Bemerkung
    
    ' set focus
    Forms!frmAuftragErteilen!txt09.SetFocus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.OnOpenfrmAuftragErteilen executed"
    End If
    
End Function
