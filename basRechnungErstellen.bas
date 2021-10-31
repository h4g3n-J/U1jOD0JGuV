Attribute VB_Name = "basRechnungErstellen"
Option Compare Database
Option Explicit

Public gvarRechnungErstellenClipboardAftrID As Variant
Public gvarRechnungErstellenClipboardBWIKey As Variant
Public gvarRechnungErstellenClipboardEAIDAngebot As Variant
Public gvarRechnungErstellenClipboardRechnungNr As Variant
Public gvarRechnungErstellenClipboardEAIDRechnung As Variant

Public Sub buildRechnungErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.buildRechnungErstellen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' clear form
     basRechnungErstellen.ClearForm strFormName
     
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
    objForm.OnOpen = "=OnOpenFrmRechnungErstellen()"
    
    ' set On Close event
    objForm.OnClose = "=OnCloseFrmRechnungErstellen()"
    
    ' declare command button
    Dim btnButton As CommandButton
    
    ' declare label
    Dim lblLabel As Label
    
    ' declare textbox
    Dim txtTextbox As TextBox
    
    ' declare combobox
    Dim cboCombobox As ComboBox
    
    ' declare checkbox
    Dim chkCheckbox As CheckBox
    
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
        intNumberOfRows = 18
        intLeft = 566
        intTop = 960
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basRechnungErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
        ' txt00
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
        
        ' txt01
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt01"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
                .Format = "Short Date"
            End With
            
        ' lbl01
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "Abgenommen am"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt02
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Caption = "Mengenger¸st"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt03
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .BorderStyle = 0
            End With
            
        ' txt04
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl04
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
            With lblLabel
                .Name = "lbl04"
                .Caption = "Bemerkung (Angebot)"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt05
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt05"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
                .Format = "Short Date"
            End With
            
        ' lbl05
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "Beauftragt am"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt06
        intColumn = 2
        intRow = 11
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt06"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl06
        intColumn = 1
        intRow = 11
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
            With lblLabel
                .Name = "lbl06"
                .Caption = "Rechnung Nummer*"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' chk07
        intColumn = 2
        intRow = 12
        Set chkCheckbox = CreateControl(strTempFormName, acCheckBox, acDetail)
            With chkCheckbox
                .Name = "chk07"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' lbl07
        intColumn = 1
        intRow = 12
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "chk07")
            With lblLabel
                .Name = "lbl07"
                .Caption = "Teilrechnung"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' chk08
        intColumn = 2
        intRow = 13
        Set chkCheckbox = CreateControl(strTempFormName, acCheckBox, acDetail)
            With chkCheckbox
                .Name = "chk08"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' lbl08
        intColumn = 1
        intRow = 13
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "chk08")
            With lblLabel
                .Name = "lbl08"
                .Caption = "Schlussrechnung"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt09
        intColumn = 2
        intRow = 14
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 1
                .Format = "Currency"
            End With
            
        ' lbl09
        intColumn = 1
        intRow = 14
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Rechnung Brutto*"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt10
        intColumn = 2
        intRow = 16
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
                .BorderStyle = 1
            End With
            
        ' lbl10
        intColumn = 1
        intRow = 16
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Rechnung Link*"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt11
        intColumn = 2
        intRow = 17
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt11"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
            End With
            
        ' lbl11
        intColumn = 1
        intRow = 17
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
            With lblLabel
                .Name = "lbl11"
                .Caption = "LNW-Kontrolle (Link)"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt12
        intColumn = 2
        intRow = 18
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt12"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 1
            End With
            
        ' lbl12
        intColumn = 1
        intRow = 18
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
            With lblLabel
                .Name = "lbl12"
                .Caption = "Bemerkung (Rechnung)"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt13
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt13"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt14
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt14"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt15
        intColumn = 2
        intRow = 7
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt15"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
            End With
            
        ' lbl15
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt15")
            With lblLabel
                .Name = "lbl15"
                .Caption = "Einzelauftrag (Angebot)"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' cbo16
        intColumn = 2
        intRow = 15
        Set cboCombobox = CreateControl(strTempFormName, acComboBox, acDetail)
            With cboCombobox
                .Name = "cbo16"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .RowSource = "tblEinzelauftrag"
                .AllowValueListEdits = False
                .ListItemsEditForm = "frmEinzelauftragErstellen"
            End With
            
        ' lbl16
        intColumn = 1
        intRow = 15
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "cbo16")
            With lblLabel
                .Name = "lbl16"
                .Caption = "Einzelauftrag (Rechnung)*"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt17
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt17"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
                .BorderStyle = 0
                .Format = "Currency"
            End With
            
        ' lbl17
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt17")
            With lblLabel
                .Name = "lbl17"
                .Caption = "Angebot Brutto"
                .Left = basRechnungErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basRechnungErstellen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basRechnungErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basRechnungErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
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
            btnButton.Caption = "Schlieﬂen"
            btnButton.OnClick = "=CloseFormRechnungErstellen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=RechnungSaveOrCreateRecordset()"
            
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.buildRechnungErstellen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.ClearForm"
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
        Debug.Print "basRechnungErstellen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' delete form
    basRechnungErstellen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basRechnungErstellen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basRechnungErstellen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basRechnungErstellen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.CalculateGrid"
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
        Debug.Print "basRechnungErstellen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestCalculateGrid"
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
    
    aintInformationGrid = basRechnungErstellen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basRechnungErstellen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basRechnungErstellen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basRechnungErstellen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basRechnungErstellen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungErstellen.GetLeft: column 0 is not available"
        MsgBox "basRechnungErstellen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basRechnungErstellen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basRechnungErstellen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungErstellen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungErstellen.GetTop: column 0 is not available"
        MsgBox "basRechnungErstellen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basRechnungErstellen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basRechnungErstellen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungErstellen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungErstellen.TestGetHeight: column 0 is not available"
        MsgBox "basRechnungErstellen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basRechnungErstellen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basRechnungErstellen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungErstellen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungErstellen.TestGetWidth: column 0 is not available"
        MsgBox "basRechnungErstellen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungErstellen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basRechnungErstellen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basRechnungErstellen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungErstellen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.TestGetWidth executed"
    End If
    
End Sub

Public Function OnCloseFrmRechnungErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.OnCloseFrmRechnungErstellen"
    End If
    
    gvarRechnungErstellenClipboardAftrID = Null
    gvarRechnungErstellenClipboardBWIKey = Null
    gvarRechnungErstellenClipboardEAIDAngebot = Null
    gvarRechnungErstellenClipboardEAIDRechnung = Null
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.OnCloseFrmRechnungErstellen executed"
    End If
    
End Function

Public Function CloseFormRechnungErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.CloseFormRechnungErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.CloseFormRechnungErstellen executed"
    End If
    
End Function

Public Function OnOpenFrmRechnungErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.OnOpenfrmRechnungErstellen"
    End If
    
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
    rstTicket.SelectRecordset (gvarRechnungErstellenClipboardAftrID)
    
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    rstAngebot.SelectRecordset (gvarRechnungErstellenClipboardBWIKey)
    
    Dim rstRechnung As clsRechnung
    Set rstRechnung = New clsRechnung
    rstRechnung.SelectRecordset (gvarRechnungErstellenClipboardRechnungNr)
    
    Forms!frmRechnungErstellen.Form!txt13 = rstTicket.AftrID
    Forms!frmRechnungErstellen.Form!txt14 = rstTicket.AftrTitel
    Forms!frmRechnungErstellen.Form!txt00 = rstAngebot.BWIKey
    Forms!frmRechnungErstellen.Form!txt02 = rstAngebot.MengengeruestLink
    Forms!frmRechnungErstellen.Form!txt03 = rstAngebot.LeistungsbeschreibungLink
    Forms!frmRechnungErstellen.Form!txt17 = rstAngebot.AngebotBrutto
    Forms!frmRechnungErstellen.Form!txt15 = gvarRechnungErstellenClipboardEAIDAngebot
    Forms!frmRechnungErstellen.Form!txt05 = rstAngebot.BeauftragtDatum
    Forms!frmRechnungErstellen.Form!txt01 = rstAngebot.AbgenommenDatum
    Forms!frmRechnungErstellen.Form!txt04 = rstAngebot.Bemerkung
    Forms!frmRechnungErstellen.Form!txt06 = rstRechnung.RechnungNr
    Forms!frmRechnungErstellen.Form!chk07 = Nz(rstRechnung.IstTeilrechnung, False)
    Forms!frmRechnungErstellen.Form!chk08 = Nz(rstRechnung.IstSchlussrechnung, False)
    Forms!frmRechnungErstellen.Form!txt09 = rstRechnung.RechnungBrutto
    Forms!frmRechnungErstellen.Form!cbo16 = gvarRechnungErstellenClipboardEAIDRechnung
    Forms!frmRechnungErstellen.Form!txt10 = rstRechnung.RechnungLink
    Forms!frmRechnungErstellen.Form!txt11 = rstRechnung.KalkulationLNWLink
    Forms!frmRechnungErstellen.Form!txt12 = rstRechnung.Bemerkung
    
    ' set focus
    Forms!frmRechnungErstellen!txt06.SetFocus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.OnOpenfrmRechnungErstellen executed"
    End If
    
End Function

Public Function RechnungSaveOrCreateRecordset()
    ' Error Code 7: AngebotZuRechung is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.RechnungSaveOrCreateRecordset"
    End If
    
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    Dim varBWIKey As Variant
    varBWIKey = Forms.Item(strFormName).Form!txt13
    
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName).Form!txt06
    
    Dim varRechnungBrutto As Variant
    varRechnungBrutto = Forms.Item(strFormName).Form!txt09
    
    Dim varRechnungLink As Variant
    varRechnungLink = Forms.Item(strFormName).Form!txt10
    
    Dim varEAIDRechnung As Variant
    varEAIDRechnung = Forms.Item(strFormName).Form!cbo16
    
    Dim strTestRelationsship As String
    
    ' # check mandatory values #
    If IsNull(varBWIKey) Then
        Debug.Print "Error: basRechnungErstellen.RechnugSaveOrCreateRecordset, Error Code 1"
        MsgBox "Es wurde keine Angebot Nummer ¸bergeben.", vbCritical, "Speichern"
    ElseIf IsNull(varRechnungNr) Then
        Debug.Print "Error: basRechnungErstellen.RechnugSaveOrCreateRecordset, Error Code 2"
        MsgBox "Sie haben in dem Pflichtfeld 'Rechnung Nummer' keinen Wert eingegeben.", vbCritical, "Speichern"
    ElseIf IsNull(varRechnungBrutto) Then
        Debug.Print "Error: basRechnungErstellen.RechnugSaveOrCreateRecordset, Error Code 3"
        MsgBox "Sie haben in dem Pflichtfeld 'Rechnung Brutto' keinen Wert eingegeben.", vbCritical, "Speichern"
    ElseIf IsNull(varRechnungLink) Then
        Debug.Print "Error: basRechnungErstellen.RechnugSaveOrCreateRecordset, Error Code 4"
        MsgBox "Sie haben in dem Pflichtfeld 'Rechnung Link' keinen Wert eingegeben.", vbCritical, "Speichern"
    ElseIf IsNull(varEAIDRechnung) Then
        Debug.Print "Error: basRechnungErstellen.RechnugSaveOrCreateRecordset, Error Code 5"
        MsgBox "Sie haben in dem Pflichtfeld 'Einzelauftrag (Rechnung)' keinen Wert eingegeben.", vbCritical, "Speichern"
    End If
    
    Dim intUserSelection As Integer
    ' check if RechnungNr is taken
    If DCount("RechnungNr", "tblRechnung", "RechnugNr like'" & varRechnungNr & "'") = 0 Then
        
        ' create Rechnung
        basRechnungErstellen.RechnungCreateRecordset
        ' create AngebotZuRechnung
        basRechnungErstellen.AngebotZuRechungCreateRecordset
        ' create EinzelauftragZuRechnung
        basRechnungErstellen.EinzelauftragZuRechnung
        
    Else
        ' get user consent to save changes on Rechnung
        intUserSelection = MsgBox("Die Rechnung " & varRechnungNr & " wurde bereits erfasst. Mˆchten Sie Ihre ƒnderungen speichern?", vbYesNo, "Speichern")
        ' evaluate input
        Select Case intUserSelection
            ' Yes
            Case 6
                ' save changes
                basRechnungErstellen.RechnungErstellenSaveRecordset
                
                ' check if AngebotZuRechnung is taken
                strTestRelationsship = varBWIKey & varRechnungNr
                
                If DCount("checksum", "qryChecksumAngebotZuRechnung", "checksum like'" & strTestRelationsship & "'") = 0 Then
                    basRechnung.AngebotZuRechnungCreateRecordset
                End If
                
                ' check if EinzelauftragZuRechnung is taken
                strTestRelationsship = varEAIDRechnung & varRechnungNr
                
                If DCount("checksum", "qryEinzelauftragZuRechnung", "checksum like '" & strTestRelationsship & "'") = 0 Then
                    basRechnung.EinzelauftragZuRechnungCreateRecordset
                End If
                
            ' No
            Case 7
                Debug.Print "Error: basRechnungErstellen.RechnungSaveOrCreateRecordset, Error Code 6"
                ExitSub
        End Select
    End If
    
    Dim strTestRelationsship As String
    
    ' check if AngebotZuRechnung is taken
    strTestRelationsship = varBWIKey & varRechnungNr
    If DCount("checksum", "qryChecksumAngebotZuRechnung", "checksum like '" & strTestRelationsship & "'") > 0 Then
        Debug.Print "Error: basRechnungErstellen.RechnungSaveOrCreateRecordset, Error Code 7"
    Else
        basRechnungErstellen.AngebotZuRechnungCreateRecordset
    End If
        
    ' check if EinzelauftragZuRechnung is taken
    strTestRelationsship = varEAIDRechnung & varRechnungNr
    If DCount("checksum", "qryChecksumEinzelauftragZuRechnung", "checksum like '" & strTestRelationsship & "'") > 0 Then
        Debug.Print "Error: basRechnungErstellen.RechnungSaveOrCreateRecordset, Error Code 8"
    Else
        basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.RechnungSaveOrCreateRecordset executed"
    End If

End Function

Private Sub RechnungCreateRecordset()
    ' Error Code 1: RechnungNr was not supplied
    ' Error Code 2: RechnungNr is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.RechnungCreateRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' declare Rechnung
    Dim rstRechnung As clsRechnung
    Set rstRechnung = New clsRechnung
    
    ' get RechnungNr
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName)!txt06
    
    ' check if RechnungNr not IsNull
    If IsNull(varRechnungNr) Then
        Debug.Print "Error: basRechnungErstellen.RechnungCreateRecordset, Error Code 1"
        Exit Sub
    End If
    
    ' check if RechnungNr is taken
    If DCount("RechnungNr", "tblRechnung", "RechnungNr like '" & varRechnungNr & "'") > 0 Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset, Error Code 2"
        Exit Sub
    End If
    
    ' move values from form to clsRechnung
    With Forms.Item(strFormName)
        rstRechnung.RechnungNr = varRechnungNr
        rstRechnung.RechnungBrutto = Forms.Item(strFormName)!txt09
        rstRechnung.RechnungLink = Forms.Item(strFormName)!txt10
    End With
    
    ' create recordset
    rstRechnung.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.RechnungCreateRecordset executed"
    End If

End Sub

Private Sub RechnungSaveRecordset()
    ' Error Code 1: RechnungNr was not supplied
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.RechnungSaveRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' get RechnungNr from form
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName)!txt06
    
    ' check if RechnungNr was supplied
    If IsNull(varRechnungNr) Then
        Debug.Print "Error: basRechnungErstellen.RechnungSaveRecordset, Error Code 1"
        Exit Sub
    End If
    
    ' select recordset
    Dim rstRechnung As clsRechnung
    Set rstRechnung = New clsRechnung
    
    rstRechnung.SelectRecordset (varRechnungNr)
    
    ' transfer values from form to Rechnung
    rstRechnung.RechnungBrutto = Forms.Item(strFormName)!txt09
    rstRechnung.RechnungLink = Forms.Item(strFormName)!txt10
    
    ' save changes to recordset
    rstRechnung.SaveRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.RechnungSaveRecordset executed"
    End If

End Sub

Private Sub AngebotZuRechnungCreateRecordset()
    ' Error Code 1: BWIKey was not supplied
    ' Error Code 2: RechnungNr was not supplied
    ' Error Code 3: RechnungZuLeistungserfassungsblattID is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.AngebotZuRechnungCreateRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' get RechnungNr from form
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName)!txt06
    
    ' get BWIKey from form
    Dim varBWIKey As Variant
    varBWIKey = Forms.Item(strFormName)!txt00
    
    ' check if RechnungNr and BWIKey were supplied
    If IsNull(varBWIKey) Then
        Debug.Print "Error: basRechnungErstellen.AngebotZuRechnungCreateRecordset, Error Code 1"
        Exit Sub
    ElseIf IsNull(varRechnungNr) Then
        Debug.Print "Error: basRechnungErstellen.AngebotZuRechnungCreateRecordset, Error Code 2"
        Exit Sub
    End If
    
    ' check if AngebotZuRechnung is taken
    Dim strTestRelationship As String
    strTestRelationship = varBWIKey & varRechnungNr
    
    If DCount("checksum", "qryChecksumAngebotZuRechnung", "checksum like '" & strTestRelationship & "'") > 0 Then
        Debug.Print "Error: basRechnungErstellen.AngebotZuRechnungCreateRecordset, Error Code 3"
        Exit Sub
    End If
    
    ' move values to AngebotZuRechnung
    Dim rstAngebotZuRechnung As clsAngebotZuRechnung
    Set rstAngebotZuRechnung = New clsAngebotZuRechnung
    
    rstAngebotZuRechnung.RefBWIkey = varBWIKey
    rstAngebotZuRechnung.RefRechnungNr = varRechnungNr
    
    ' create AngebotZuRechnung recordset
    rstAngebotZuRechnung.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.AngebotZuRechnungCreateRecordset executed"
    End If

End Sub

Private Sub EinzelauftragZuRechnungCreateRecordset()
    ' Error Code 1: EAIDRechnung was not supplied
    ' Error Code 2: RechnungNr was not supplied
    ' Error Code 3: EinzelauftragZuRechnung is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' get RechnungNr from form
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName)!txt06
    
    ' get EAIDRechnung from form
    Dim varEAIDRechnung As Variant
    varEAIDRechnung = Forms.Item(strFormName)!cbo16
    
    ' check if RechnungNr and EAIDRechnung were supplied
    If IsNull(varRechnungNr) Then
        Debug.Print "Error: basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset, Error Code 2"
        Exit Sub
    ElseIf IsNull(varEAIDRechnung) Then
        Debug.Print "Error: basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset, Error Code 1"
        Exit Sub
    End If
    
    ' check if EinzelauftragZuRechnung is taken
    Dim strTestRelationship As String
    strTestRelationship = varEAIDRechnung & varRechnungNr
    
    If DCount("checksum", "qryChecksumEinzelauftragZuRechnung", "checksum like'" & strTestRelationship & "'") > 0 Then
        Debug.Print "Error: basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset, Error Code 3"
        Exit Sub
    End If
    
    ' move values to EinzelauftragZuRechnung
    Dim rstEinzelauftragZuRechnung As clsEinzelauftragZuRechnung
    Set rstEinzelauftragZuRechnung = New clsEinzelauftragZuRechnung
    
    rstEinzelauftragZuRechnung.RefEAkurzKey = varEAIDRechnung
    rstEinzelauftragZuRechnung.RefRechnungNr = varRechnungNr
    
    ' create EinzelauftragZuRechnung
    rstEinzelauftragZuRechnung.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungErstellen.EinzelauftragZuRechnungCreateRecordset executed"
    End If

End Sub

