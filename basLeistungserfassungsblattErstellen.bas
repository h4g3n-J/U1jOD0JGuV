Attribute VB_Name = "basLeistungserfassungsblattErstellen"
Option Compare Database
Option Explicit

Public gvarLeistungserfassungsblattErstellenClipboardAftrID As Variant
Public gvarLeistungserfassungsblattErstellenClipboardBWIKey As Variant
Public gvarLeistungserfassungsblattErstellenClipboardEAIDAngebot As Variant
Public gvarLeistungserfassungsblattErstellenClipboardRechnungNr As Variant
Public gvarLeistungserfassungsblattErstellenClipboardEAIDRechnung As Variant
Public gvarLeistungserfassungsblattErstellenClipboardLeistungserfassungsblattID As Variant
Public gvarLeistungserfassungsblattErstellenClipboardRechnungZuLeistungserfassungsblattID As Variant

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
    
    ' set On Open event
    objForm.OnOpen = "=OnOpenFrmLeistungserfassungsblattErstellen()"
    
    ' set On Close event
    objForm.OnClose = "=OnCloseFrmLeistungserfassungsblattErstellen()"
    
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
        intNumberOfRows = 19
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
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
    
    'lbl00
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "Angebot ID"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 6
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .Format = "short date"
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl01
    intColumn = 1
    intRow = 6
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Abgenommen Datum"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl02
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Mengengerüst"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl03
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Leistungsbeschreibung"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 7
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl04
    intColumn = 1
    intRow = 7
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Bemerkung (Angebot)"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 2
    intRow = 15
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl05
    intColumn = 1
    intRow = 15
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "LEB Nummer*"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 2
    intRow = 8
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl06
    intColumn = 1
    intRow = 8
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "Rechnung Nummer"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'chk07
    intColumn = 2
    intRow = 9
    Set chkCheckbox = CreateControl(strTempFormName, acCheckBox, acDetail)
        With chkCheckbox
            .Name = "chk07"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .Locked = True
        End With
        
    'lbl07
    intColumn = 1
    intRow = 9
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "chk07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "Teilrechnung"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'chk08
    intColumn = 2
    intRow = 10
    Set chkCheckbox = CreateControl(strTempFormName, acCheckBox, acDetail)
        With chkCheckbox
            .Name = "chk08"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .Locked = True
        End With
        
    'lbl08
    intColumn = 1
    intRow = 10
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "chk08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "Schlussrechnung"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt09
    intColumn = 2
    intRow = 11
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .Format = "Currency"
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl09
    intColumn = 1
    intRow = 11
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "Rechnung Brutto"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt10
    intColumn = 2
    intRow = 13
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl10
    intColumn = 1
    intRow = 13
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "Rechnung Link"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt11
    intColumn = 2
    intRow = 16
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl11
    intColumn = 1
    intRow = 16
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "Bemerkung (LEB)"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt12
    intColumn = 2
    intRow = 14
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt12"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl12
    intColumn = 1
    intRow = 14
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
        With lblLabel
            .Name = "lbl12"
            .Caption = "Bemerkung (Rechnung)"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt13
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt13"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl13
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt13")
        With lblLabel
            .Name = "lbl13"
            .Caption = "Ticket ID"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt14
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt14"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl14
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt14")
        With lblLabel
            .Name = "lbl14"
            .Caption = "Zusammenfassung"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt15
    intColumn = 2
    intRow = 17
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt15"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl15
    intColumn = 1
    intRow = 17
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt15")
        With lblLabel
            .Name = "lbl15"
            .Caption = "Beleg ID"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt16
    intColumn = 2
    intRow = 12
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt16"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .BorderStyle = 0
            .Locked = True
        End With
        
    'lbl16
    intColumn = 1
    intRow = 12
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt16")
        With lblLabel
            .Name = "lbl16"
            .Caption = "Einzelauftrag (Rechnung)"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt17
    intColumn = 2
    intRow = 18
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt17"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .Format = "Currency"
        End With
        
    'lbl17
    intColumn = 1
    intRow = 18
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt17")
        With lblLabel
            .Name = "lbl17"
            .Caption = "LEB Brutto*"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt18
    intColumn = 2
    intRow = 19
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt18"
            .Left = basLeistungserfassungsblattErstellen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattErstellen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattErstellen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattErstellen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl18
    intColumn = 1
    intRow = 19
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt18")
        With lblLabel
            .Name = "lbl18"
            .Caption = "Haushaltsjahr"
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
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormLeistungserfassungsblattErstellen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 7653
            btnButton.Top = 1350
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=LeistungserfassungsblattSaveOrCreateRecordset()"
                
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

Private Sub LeistungserfassungsblattCreateRecordset()
    ' Error Code 1: LeistungserfassungsblattID was not supplied
    ' Error Code 2: LeistungserfassungsblattID is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
        
    ' get LeistungserfassungsblattID
    Dim strLeistungserfassungsblattID As String
    strLeistungserfassungsblattID = Forms.Item(strFormName)!txt05

    ' check if LeistungserfassungsblattID is supplied
    If IsNull(strLeistungserfassungsblattID) Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset, Error Code 1"
        Exit Sub
    End If

    ' check if LeistungserfassungsblattID is taken
    If DCount("LeistungserfassungsblattID", "tblLeistungserfassungsblatt", "LeistungserfassungsblattID like '" & strLeistungserfassungsblattID & "'") > 0 Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset, Error Code 2"
        Exit Sub
    End If

    ' declare Leistungserfassungsblatt
    Dim rstLeistungserfassungsblatt As clsLeistungserfassungsblatt
    Set rstLeistungserfassungsblatt = New clsLeistungserfassungsblatt

    ' transfer values from form to clsLeistungserfassungsblatt
    With Forms.Item(strFormName)
        rstLeistungserfassungsblatt.LeistungserfassungsblattID = strLeistungserfassungsblattID
        rstLeistungserfassungsblatt.Brutto = !txt17
        rstLeistungserfassungsblatt.Bemerkung = !txt11
        rstLeistungserfassungsblatt.BelegID = !txt15
        rstLeistungserfassungsblatt.Haushaltsjahr = !txt18
    End With
    
    ' create Recordset
    rstLeistungserfassungsblatt.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset executed"
    End If
    
End Sub

Private Sub LeistungserfassungsblattSaveRecordset()
    ' Error Code 1: recordsetID was not supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveRecordset"
    End If
        
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
        
    ' get LeistungserfassungsblattID from form
    Dim strLeistungserfassungsblattID As String
    strLeistungserfassungsblattID = Forms.Item(strFormName)!txt05

    ' check if LeistungserfassungsblattID is supplied
    If IsNull(strLeistungserfassungsblattID) Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveRecordset, Error Code 1"
        Exit Function
    End If

    ' declare Leistungserfassungsblatt
    Dim rstLeistungserfassungsblatt As clsLeistungserfassungsblatt
    Set rstLeistungserfassungsblatt = New clsLeistungserfassungsblatt

    ' select recordset
    rstLeistungserfassungsblatt.SelectRecordset (strLeistungserfassungsblattID)
    
    ' transfer values from form to clsLeistungserfassungsblatt
    With Forms.Item(strFormName)
        rstLeistungserfassungsblatt.RechnungNr = !txt06
        rstLeistungserfassungsblatt.Bemerkung = !txt11
        rstLeistungserfassungsblatt.BelegID = !txt03
        rstLeistungserfassungsblatt.Brutto = !txt17
        rstLeistungserfassungsblatt.Haushaltsjahr = !txt18
    End With
    
    ' save changes to recordset
    rstLeistungserfassungsblatt.SaveRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveRecordset executed"
    End If
    
End Sub

Private Sub RechnungZuLeistungserfassungsblattCreateRecordset()
    ' Error Code 1: LeistungserfassungsblattID was not supplied
    ' Error Code 2: RechnungNr was not supplied
    ' Error Code 3: RechnungZuLeistungserfassungsblattID is taken

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    ' get RechnungNr from form
    Dim strRechnungNr As String
    strRechnungNr = Forms.Item(strFormName)!txt06
    
    ' get LeistungserfassungsblattID from form
    Dim strLeistungserfassungsblattID As String
    strLeistungserfassungsblattID = Forms.Item(strFormName)!txt05
    
    ' declare RechnungZuLeistungserfassungsblatt
    Dim rstRechnungZuLeistungserfassungsblatt As clsRechnungZuLeistungserfassungsblatt
    Set rstRechnungZuLeistungserfassungsblatt = New clsRechnungZuLeistungserfassungsblatt
    
    ' check if RechnungNr and LeistungserfassungsblattID are supplied
    If IsNull(strLeistungserfassungsblattID) Then
        Debug.Print "Error: LeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset Error Code 1"
        Exit Sub
    ElseIf IsNull(strRechnungNr) Then
        Debug.Print "Error: LeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset Error Code 2"
        Exit Sub
    End If

    Dim strTestRelationship As String
    strTestRelationship = strRechnungNr & strLeistungserfassungsblattID

    ' check if RechnungZuLeistungserfassungsblatt is taken
    If DCount("checksum", "qryChecksumRechnungZuLeistungserfassungsblatt", "checksum like '" & strTestRelationship & "'") > 0 Then
        Debug.Print "Error: RechnungZuLeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset, Error Code 3"
        Exit Sub
    End If

    ' transfer values to RechnungZuLeistungserfassungsblatt
    rstRechnungZuLeistungserfassungsblatt.RefLeistungserfassungsblattID = strLeistungserfassungsblattID
    rstRechnungZuLeistungserfassungsblatt.RefRechnungNr = strRechnungNr
    
    ' create RechnungZuLeistungserfassungsblatt recordset
    rstRechnungZuLeistungserfassungsblatt.CreateRecordset
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset executed"
    End If

End Sub

Public Function LeistungserfassungsblattSaveOrCreateRecordset()
    ' Error Code 1: LeistungserfassungsblattID was not supplied
    ' Error Code 2: RechnungNr was not supplied
    ' Error Code 3: Brutto was not supplied
    ' Error Code 4: user canceled function
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset"
    End If

    Dim intUserSelection As Integer
    intUserSelection = 0
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    ' get RechnungNr from form
    Dim varRechnungNr As Variant
    varRechnungNr = Forms.Item(strFormName)!txt06
    
    ' get LeistungserfassungsblattID from form
    Dim varLeistungserfassungsblattID As Variant
    varLeistungserfassungsblattID = Forms.Item(strFormName)!txt05

    ' get Brutto from form
    Dim varBrutto As Variant
    varBrutto = Forms.Item(strFormName)!txt17
    
    ' # check mandatory values #
    ' check isNull(LeistungserfassungsblattID)
    If IsNull(varLeistungserfassungsblattID) Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset, Error Code 1"
        MsgBox "Sie haben im Pflichtfeld 'LEB Nummer' keinen Wert eingegeben.", vbCritical, "Speichern"
        GoTo ExitProc
    ' check isNull(RechnungNr)
    ElseIf IsNull(varRechnungNr) Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset, Error Code 2"
        MsgBox "Es wurde keine Rechnung Nummer übergeben. Speichern abgebrochen", vbCritical, "Speichern"
        GoTo ExitProc
    ' check isNull(Brutto)
    ElseIf IsNull(varBrutto) Then
        Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset, Error Code 3"
        MsgBox "Sie haben im Pflichtfeld 'LEB Brutto' keinen Wert eingegeben.", vbCritical, "Speichern"
        GoTo ExitProc
    End If
    
    ' check if LeistungserfassungsblattID is taken
    If DCount("LeistungserfassungsblattID", "tblLeistungserfassungsblatt", "LeistungserfassungsblattID like '" & varLeistungserfassungsblattID & "'") = 0 Then
        
        ' create Leistungserfassungsblatt
        basLeistungserfassungsblattErstellen.LeistungserfassungsblattCreateRecordset
        
        ' create RechnungZuLeistungserfassungsblatt
        basLeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset
        
        GoTo ExitProc
        
    Else
        ' get user consent to save changes to Leistungserfassungsblatt
        intUserSelection = MsgBox("Das Leistungserfassungsblatt '" & varLeistungserfassungsblattID & "' wurde bereits erfasst. Möchten Sie Ihre Änderungen speichern?", vbYesNo, "Speichern")
            ' evaluate input
            Select Case intUserSelection
                ' Yes
                Case 6
                    ' save changes
                    basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveRecordset
                ' No
                Case 7
                    Debug.Print "Error: basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset, Error Code 4"
                    GoTo ExitProc
            End Select
    End If
    
    ' # check if RechnungZuLeistungserfassungsblatt is taken #
    Dim strTestRelationship As String
    strTestRelationship = varRechnungNr & varLeistungserfassungsblattID

    If DCount("checksum", "qryChecksumRechnungZuLeistungserfassungsblatt", "checksum like '" & strTestRelationship & "'") = 0 Then
        basLeistungserfassungsblattErstellen.RechnungZuLeistungserfassungsblattCreateRecordset
    End If

ExitProc:
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.LeistungserfassungsblattSaveOrCreateRecordset executed"
    End If

End Function

Public Function OnOpenFrmLeistungserfassungsblattErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.OnOpenfrmLeistungserfassungsblattErstellen"
    End If
    
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
    rstTicket.SelectRecordset (gvarLeistungserfassungsblattErstellenClipboardAftrID)
    
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    rstAngebot.SelectRecordset (gvarLeistungserfassungsblattErstellenClipboardBWIKey)
    
    Dim rstRechnung As clsRechnung
    Set rstRechnung = New clsRechnung
    rstRechnung.SelectRecordset (gvarLeistungserfassungsblattErstellenClipboardRechnungNr)
    
    Dim rstLeistungserfassungsblatt As clsLeistungserfassungsblatt
    Set rstLeistungserfassungsblatt = New clsLeistungserfassungsblatt
    rstLeistungserfassungsblatt.SelectRecordset (gvarLeistungserfassungsblattErstellenClipboardLeistungserfassungsblattID)
    
    Forms!frmLeistungserfassungsblattErstellen.Form!txt13 = rstTicket.AftrID
    Forms!frmLeistungserfassungsblattErstellen.Form!txt14 = rstTicket.AftrTitel
    Forms!frmLeistungserfassungsblattErstellen.Form!txt00 = rstAngebot.BWIKey
    Forms!frmLeistungserfassungsblattErstellen.Form!txt02 = rstAngebot.MengengeruestLink
    Forms!frmLeistungserfassungsblattErstellen.Form!txt03 = rstAngebot.LeistungsbeschreibungLink
    Forms!frmLeistungserfassungsblattErstellen.Form!txt01 = rstAngebot.AbgenommenDatum
    Forms!frmLeistungserfassungsblattErstellen.Form!txt04 = rstAngebot.Bemerkung
    Forms!frmLeistungserfassungsblattErstellen.Form!txt06 = rstRechnung.RechnungNr
    Forms!frmLeistungserfassungsblattErstellen.Form!chk07 = Nz(rstRechnung.IstTeilrechnung, False)
    Forms!frmLeistungserfassungsblattErstellen.Form!chk08 = Nz(rstRechnung.IstSchlussrechnung, False)
    Forms!frmLeistungserfassungsblattErstellen.Form!txt09 = rstRechnung.RechnungBrutto
    Forms!frmLeistungserfassungsblattErstellen.Form!txt16 = gvarLeistungserfassungsblattErstellenClipboardEAIDRechnung
    Forms!frmLeistungserfassungsblattErstellen.Form!txt10 = rstRechnung.RechnungLink
    Forms!frmLeistungserfassungsblattErstellen.Form!txt12 = rstRechnung.Bemerkung
    Forms!frmLeistungserfassungsblattErstellen.Form!txt05 = rstLeistungserfassungsblatt.LeistungserfassungsblattID
    Forms!frmLeistungserfassungsblattErstellen.Form!txt11 = rstLeistungserfassungsblatt.Bemerkung
    Forms!frmLeistungserfassungsblattErstellen.Form!txt15 = rstLeistungserfassungsblatt.BelegID
    Forms!frmLeistungserfassungsblattErstellen.Form!txt17 = rstLeistungserfassungsblatt.Brutto
    Forms!frmLeistungserfassungsblattErstellen.Form!txt18 = rstLeistungserfassungsblatt.Haushaltsjahr

    ' set focus
    Forms!frmLeistungserfassungsblattErstellen!txt05.SetFocus
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.OnOpenfrmLeistungserfassungsblattErstellen executed"
    End If
    
End Function

Public Function CloseFormLeistungserfassungsblattErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.CloseFormLeistungserfassungsblattErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.CloseFormLeistungserfassungsblattErstellen executed"
    End If
    
End Function

Public Function OnCloseFrmLeistungserfassungsblattErstellen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattErstellen.OnCloseFrmLeistungserfassungsblattErstellen"
    End If
    
    gvarLeistungserfassungsblattErstellenClipboardAftrID = Null
    gvarLeistungserfassungsblattErstellenClipboardBWIKey = Null
    gvarLeistungserfassungsblattErstellenClipboardEAIDAngebot = Null
    gvarLeistungserfassungsblattErstellenClipboardRechnungNr = Null
    gvarLeistungserfassungsblattErstellenClipboardEAIDRechnung = Null
    gvarLeistungserfassungsblattErstellenClipboardLeistungserfassungsblattID = Null

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattErstellen.OnCloseFrmLeistungserfassungsblattErstellen executed"
    End If
    
End Function
