Attribute VB_Name = "basAngebotErstellen"
Option Compare Database
Option Explicit

Public gvarAngebotErstellenClipboardAftrID As Variant
Public gvarAngebotErstellenClipboardAftrTitel As Variant

Public Sub buildAngebotErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.buildAngebotErstellen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    ' clear form
     basAngebotErstellen.ClearForm strFormName
     
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
    objForm.OnOpen = "=OnOpenFrmAngebotErstellen()"
    
    ' set On Close event
    objForm.OnClose = "=OnCloseFrmAngebotErstellen()"
    
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
        aintInformationGrid = basAngebotErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
        ' txt00
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl00
        intColumn = 1
        intRow = 3
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
            With lblLabel
                .Name = "lbl00"
                .Caption = "Angebot ID*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
        
        ' cbo01
        intColumn = 2
        intRow = 7
        Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
            With cboCombobox
                .Name = "cbo01"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .RowSource = "tblEinzelauftrag"
                .AllowValueListEdits = False
                .ListItemsEditForm = "frmEinzelauftragErstellen"
            End With
            
        ' lbl01
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "Einzelauftrag*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt02
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
            End With
            
        ' lbl02
        intColumn = 1
        intRow = 5
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
            With lblLabel
                .Name = "lbl02"
                .Caption = "Mengengerüst*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt03
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
            End With
            
        ' lbl03
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
            With lblLabel
                .Name = "lbl03"
                .Caption = "Leistungsbeschreibung*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt04
        intColumn = 2
        intRow = 11
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt05
        ' intColumn = 2
        ' intRow = 6
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt05"
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
                ' .IsHyperlink = False
            ' End With
            
        ' ' lbl05
        ' intColumn = 1
        ' intRow = 6
        ' Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            ' With lblLabel
                ' .Name = "lbl05"
                ' .Caption = "Beauftragt"
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' ' txt06
        ' intColumn = 2
        ' intRow = 7
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt06"
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt07
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt07"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .Format = "short date"
            End With
            
        ' lbl07
        intColumn = 1
        intRow = 4
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
            With lblLabel
                .Name = "lbl07"
                .Caption = "Angeboten am*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt08
        ' intColumn = 2
        ' intRow = 9
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt08"
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt09
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .Format = "short date"
            End With
            
        ' lbl09
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Auftrag Beginn"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt10
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .Format = "short date"
            End With
            
        ' lbl10
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Auftrag Ende"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' ' txt11
        ' intColumn = 2
        ' intRow = 12
        ' Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            ' With txtTextbox
                ' .Name = "txt11"
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                ' .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                ' .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                ' .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                ' .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                ' .Visible = True
            ' End With
            
        ' txt12
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt12"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .Format = "Currency"
            End With
            
        ' lbl12
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
            With lblLabel
                .Name = "lbl12"
                .Caption = "Preis Brutto*"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt13
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt13"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt14
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt14"
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basAngebotErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            lblLabel.Caption = "Angebot erstellen"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 7653
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormAngebotErstellen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AngebotErstellenCreateRecordset()"
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.buildAngebotErstellen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.ClearForm"
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
        Debug.Print "basAngebotErstellen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    ' delete form
    basAngebotErstellen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAngebotErstellen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAngebotErstellen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAngebotErstellen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.CalculateGrid"
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
        Debug.Print "basAngebotErstellen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestCalculateGrid"
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
    
    aintInformationGrid = basAngebotErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basAngebotErstellen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAngebotErstellen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAngebotErstellen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAngebotErstellen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotErstellen.GetLeft: column 0 is not available"
        MsgBox "basAngebotErstellen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basAngebotErstellen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAngebotErstellen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotErstellen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotErstellen.GetTop: column 0 is not available"
        MsgBox "basAngebotErstellen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basAngebotErstellen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAngebotErstellen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotErstellen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotErstellen.TestGetHeight: column 0 is not available"
        MsgBox "basAngebotErstellen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAngebotErstellen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAngebotErstellen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotErstellen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAngebotErstellen.TestGetWidth: column 0 is not available"
        MsgBox "basAngebotErstellen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAngebotErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAngebotErstellen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAngebotErstellen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAngebotErstellen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.TestGetWidth executed"
    End If
    
End Sub

Public Function CloseFormAngebotErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.CloseFormAngebotErstellen"
    End If
    
    gvarAngebotErstellenClipboardAftrID = Null
    gvarAngebotErstellenClipboardAftrTitel = Null
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.CloseFormAngebotErstellen executed"
    End If
    
End Function

Public Function AngebotErstellenCreateRecordset()
' Error Code 1: no value assigned to BWIKey
' Error Code 2: recordset of that name already exists
' Error Code 3: input value is not on the value list

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.AngebotErstellenCreateRecordset"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    ' get mandatory values
    Dim varBWIKey As Variant
    varBWIKey = Forms.Item(strFormName)!txt00
    
    Dim varEAID As Variant
    varEAID = Forms.Item(strFormName)!cbo01
    
    Dim varMengengeruest As Variant
    varMengengeruest = Forms.Item(strFormName)!txt02
    
    Dim varLeistungsbeschreibug As Variant
    varLeistungsbeschreibug = Forms.Item(strFormName)!txt03
    
    Dim varAngebotDatum As Variant
    varAngebotDatum = Forms.Item(strFormName)!txt07
    
    Dim varAngebotBrutto As Variant
    varAngebotBrutto = Forms.Item(strFormName)!txt12
    
    ' check mandatory fields
    ' check if varBWIKey is empty
    If IsNull(varBWIKey) Then
        MsgBox "Sie haben im Pflichtfeld 'Angebot ID' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
        ' check if varBWIKey is taken
        ElseIf DCount("[BWIKey]", "tblAngebot", "[BWIkey] Like '" & varBWIKey & "'") > 0 Then
            MsgBox "Das Angebot '" & varBWIKey & "' wurde bereits erfasst.", vbCritical, "Speichern abgebrochen"
            Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset failed, Error Code 2"
            Exit Function
    End If
        
    ' check if varAngebotDatum is empty
    If IsNull(varAngebotDatum) Then
        MsgBox "Sie haben im Pflichtfeld 'Angeboten am' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
    End If
    
    ' check if varMengengeruest is empty
    If IsNull(varMengengeruest) Then
        MsgBox "Sie haben im Pflichtfeld 'Mengengerüst' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
    End If
    
    ' check if varLeistungsbeschreibug is empty
    If IsNull(varLeistungsbeschreibug) Then
        MsgBox "Sie haben im Pflichtfeld 'Leistungsbeschreibung' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
    End If
        
    ' check if varEAID is empty
    If IsNull(varEAID) Then
        MsgBox "Sie haben im Pflichtfeld 'Einzelauftrag' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
    ' check for forbidden values in varEAID
    ElseIf DCount("[EAkurzKey]", "tblEinzelauftrag", "[EAkurzKey] Like '" & varEAID & "'") = 0 Then
        MsgBox "Bitte wählen Sie im Feld 'Einzelauftrag ID' ausschließlich Werte aus der Drop-Down-Liste.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 3"
        Exit Function
    End If
    
    
    ' check if varAngebotBrutto is empty
    If IsNull(varAngebotBrutto) Then
        MsgBox "Sie haben im Pflichtfeld 'Preis Brutto' keinen Wert eingegeben.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: basAngebotErstellen.AngebotErstellenCreateRecordset, Error Code 1"
        Exit Function
    End If
    
    Dim strTestObject01 As String
    strTestObject01 = varEAID & varBWIKey
    
    If DCount("[checksum]", "qryChecksumEinzelauftragZuAngebot", "[checksum] Like '" & strTestObject01 & "'") > 0 Then
        MsgBox "Das Angebot '" & varBWIKey & "' wurde bereits mit dem Einzelauftrag '" & varEAID & "' verknüpft.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: clsEinzelauftragZuAngebot.CreateRecordset, Error Code 1"
        Exit Function
    End If
    
    Dim strTestObject02 As String
    ' strTestObject02 = varAuftragID & varBWIKey
    strTestObject02 = gvarAngebotErstellenClipboardAftrID & varBWIKey
    
    If DCount("[checksum]", "qryChecksumAuftragZuAngebot", "[checksum] Like '" & strTestObject02 & "'") > 0 Then
        MsgBox "Das Angebot '" & varBWIKey & "' wurde bereits mit dem Ticket '" & gvarAngebotErstellenClipboardAftrID & "' verknüpft.", vbCritical, "Speichern abgebrochen"
        Debug.Print "Error: clsEinzelauftragZuAngebot.CreateRecordset, Error Code 1"
        Exit Function
    End If
    
    ' transfer values to clsAngebot
    Dim rstRecordset01 As clsAngebot
    Set rstRecordset01 = New clsAngebot
    
    With Forms.Item(strFormName)
        rstRecordset01.BWIKey = varBWIKey
        rstRecordset01.MengengeruestLink = !txt02
        rstRecordset01.LeistungsbeschreibungLink = !txt03
        rstRecordset01.Bemerkung = !txt04
        ' rstRecordset01.BeauftragtDatum = !txt05
        ' rstRecordset01.AbgebrochenDatum = !txt06
        rstRecordset01.AngebotDatum = !txt07
        ' rstRecordset01.AbgenommenDatum = !txt08
        rstRecordset01.AftrBeginn = !txt09
        rstRecordset01.AftrEnde = !txt10
        ' rstRecordset01.StorniertDatum = !txt11
        rstRecordset01.AngebotBrutto = !txt12
    End With
    
    ' transfer values to clsEinzelauftratZuAngebot
    Dim rstRecordset02 As clsEinzelauftragZuAngebot
    Set rstRecordset02 = New clsEinzelauftragZuAngebot
    
    rstRecordset02.RefBWIkey = varBWIKey
    rstRecordset02.RefEAkurzKey = varEAID
    
    ' transfer values to clsAuftragZuAngebot
    Dim rstRecordset03 As clsAuftragZuAngebot
    Set rstRecordset03 = New clsAuftragZuAngebot
    
    rstRecordset03.RefAftrID = gvarAngebotErstellenClipboardAftrID
    rstRecordset03.RefBWIkey = varBWIKey
    
    ' create Recordset clsAngebot
    rstRecordset01.CreateRecordset
    
    ' create recordset clsEinzelauftragZuAngebot
    rstRecordset02.CreateRecordset
    
    ' create recordset clsAuftragZuAngebot
    rstRecordset03.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.AngebotErstellenCreateRecordset executed"
    End If
    
ExitProc:
    Set rstRecordset01 = Nothing
    Set rstRecordset02 = Nothing
    Set rstRecordset03 = Nothing
    
End Function

Public Function OnOpenFrmAngebotErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.OnOpenFrmAngebotErstellen"
    End If
    
    Forms!frmAngebotErstellen.Form!txt13 = gvarAngebotErstellenClipboardAftrID
    Forms!frmAngebotErstellen.Form!txt14 = gvarAngebotErstellenClipboardAftrTitel
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.OnOpenFrmAngebotErstellen executed"
    End If
    
End Function

Public Function OnCloseFrmAngebotErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotErstellen.OnCloseFrmAngebotErstellen"
    End If
    
    gvarAngebotErstellenClipboardAftrID = Null
    gvarAngebotErstellenClipboardAftrTitel = Null
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotErstellen.OnCloseFrmAngebotErstellen executed"
    End If
    
End Function
