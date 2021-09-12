Attribute VB_Name = "basEinzelauftragSuchen"
Option Compare Database
Option Explicit

Public Sub BuildEinzelauftragSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.BuildEinzelauftragSuchen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    ' clear form
     basEinzelauftragSuchen.ClearForm strFormName
     
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
        intNumberOfRows = 15
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "EAkurzKey"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "MengengeruestLink"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "LeistungsbeschreibungLink"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Bemerkung"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl04
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "BeauftragtDatum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 2
    intRow = 6
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl05
    intColumn = 1
    intRow = 6
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "AbgebrochenDatum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 2
    intRow = 7
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl06
    intColumn = 1
    intRow = 7
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "MitzeichnungI21Datum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 2
    intRow = 8
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl07
    intColumn = 1
    intRow = 8
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "MitzeichnungI25Datum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt08
    intColumn = 2
    intRow = 9
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl08
    intColumn = 1
    intRow = 9
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "AngebotDatum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt09
    intColumn = 2
    intRow = 10
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl09
    intColumn = 1
    intRow = 10
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "AbgenommenDatum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt10
    intColumn = 2
    intRow = 11
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl10
    intColumn = 1
    intRow = 11
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "StorniertDatum"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt11
    intColumn = 2
    intRow = 12
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl11
    intColumn = 1
    intRow = 12
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "AngebotBrutto"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt12
    intColumn = 2
    intRow = 13
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt12"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl12
    intColumn = 1
    intRow = 13
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
        With lblLabel
            .Name = "lbl12"
            .Caption = "BWIKey"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt13
    intColumn = 2
    intRow = 14
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt13"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl13
    intColumn = 1
    intRow = 14
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt13")
        With lblLabel
            .Name = "lbl13"
            .Caption = "AftrBeginn"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt14
    intColumn = 2
    intRow = 15
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt14"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl14
    intColumn = 1
    intRow = 15
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt14")
        With lblLabel
            .Name = "lbl14"
            .Caption = "AftrEnde"
            .Left = basEinzelauftragSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchen.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
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
        
        aintLifecycleGrid = basEinzelauftragSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Einzelauftrag erstellen" button
        intColumn = 1
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateEinzelauftrag"
                .Left = basEinzelauftragSuchen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basEinzelauftragSuchen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basEinzelauftragSuchen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basEinzelauftragSuchen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Einzelauftrag erstellen"
                .OnClick = "=OpenFormEinzelauftragErstellen()"
                .Visible = True
            End With
            
        ' create form title
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
            lblLabel.Name = "lblTitle"
            lblLabel.Visible = True
            lblLabel.Left = 566
            lblLabel.Top = 227
            lblLabel.Width = 9210
            lblLabel.Height = 507
            lblLabel.Caption = "Einzelauftrag Suchen"
            
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
            btnButton.OnClick = "=SearchAndReloadEinzelauftragSuchen()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormEinzelauftragSuchen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 13180
            btnButton.Top = 1425
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=EinzelauftragSuchenSaveRecordset()"
            
        ' create deleteRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdDeleteRecordset"
            btnButton.Left = 13180
            btnButton.Top = 1875
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Datensatz löschen"
            btnButton.OnClick = "=EinzelauftragSuchenDeleteRecordset()"

        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmEinzelauftragSuchenSub"
            .Locked = True
        End With
                
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.BuildEinzelauftragSuchen executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.ClearForm"
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
        Debug.Print "basEinzelauftragSuchen.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    ' delete form
    basEinzelauftragSuchen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basEinzelauftragSuchen.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basEinzelauftragSuchen.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basEinzelauftragSuchen.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.CalculateGrid"
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
        Debug.Print "basEinzelauftragSuchen.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestCalculateGrid"
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
    
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basEinzelauftragSuchen.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basEinzelauftragSuchen.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basEinzelauftragSuchen.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basEinzelauftragSuchen.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchen.GetLeft: column 0 is not available"
        MsgBox "basEinzelauftragSuchen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basEinzelauftragSuchen.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basEinzelauftragSuchen.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchen.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchen.GetTop: column 0 is not available"
        MsgBox "basEinzelauftragSuchen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basEinzelauftragSuchen.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basEinzelauftragSuchen.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchen.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchen.TestGetHeight: column 0 is not available"
        MsgBox "basEinzelauftragSuchen.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basEinzelauftragSuchen.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basEinzelauftragSuchen.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchen.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchen.TestGetWidth: column 0 is not available"
        MsgBox "basEinzelauftragSuchen.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchen.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basEinzelauftragSuchen.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basEinzelauftragSuchen.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchen.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadEinzelauftragSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.SearchAndReloadEinzelauftragSuchen"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search Rechnung
    Dim strQueryName As String
    strQueryName = "qryEinzelauftragSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblEinzelauftrag"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "EAkurzKey"
    
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    basEinzelauftragSuchenSub.SearchEinzelauftrag strQueryName, strQuerySource, strPrimaryKey, varSearchTerm
    
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.SearchAndReloadEinzelauftragSuchen executed"
    End If
    
End Function

Public Function CloseFormEinzelauftragSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.CloseForm executed"
    End If
    
End Function

Public Function OpenFormEinzelauftragErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.OpenFormEinzelauftragErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.OpenFormEinzelauftragErstellen executed"
    End If
    
End Function

Public Function EinzelauftragSuchenSaveRecordset()
    ' Error Code 1: no recordset was supplied
    ' Error Code 2: user aborted function

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    ' declare subform name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "EAkurzKey"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Einzelauftrag
    Dim Einzelauftrag As clsEinzelauftrag
    Set Einzelauftrag = New clsEinzelauftrag
    
    ' check primary key value
    If IsNull(varRecordsetName) Then
        Debug.Print "Error: basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt. Speichern abgebrochen.", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' select recordset
    Einzelauftrag.SelectRecordset varRecordsetName
    
    ' allocate values to recordset properties
    With Einzelauftrag
        .MengengeruestLink = Forms.Item(strFormName).Controls("txt01")
        .LeistungsbeschreibungLink = Forms.Item(strFormName).Controls("txt02")
        .Bemerkung = Forms.Item(strFormName).Controls("txt03")
        .BeauftragtDatum = Forms.Item(strFormName).Controls("txt04")
        .AbgebrochenDatum = Forms.Item(strFormName).Controls("txt05")
        .MitzeichnungI21Datum = Forms.Item(strFormName).Controls("txt06")
        .MitzeichnungI25Datum = Forms.Item(strFormName).Controls("txt07")
        .AngebotDatum = Forms.Item(strFormName).Controls("txt08")
        .AbgenommenDatum = Forms.Item(strFormName).Controls("txt09")
        .StorniertDatum = Forms.Item(strFormName).Controls("txt10")
        .AngebotBrutto = Forms.Item(strFormName).Controls("txt11")
        .BWIKey = Forms.Item(strFormName).Controls("txt12")
        .AftrBeginn = Forms.Item(strFormName).Controls("txt13")
        .AftrEnde = Forms.Item(strFormName).Controls("txt14")
    End With
    
    ' delete recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Datensatz " & varRecordsetName & " wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
    If varUserInput = 1 Then
        Einzelauftrag.SaveRecordset
        MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
    Else
        Debug.Print "Error: basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset aborted, Error Code 2"
        MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset execute"
    End If
    
End Function

Public Function EinzelauftragSuchenDeleteRecordset()
    ' Error Code 1: no recordset was supplied
    ' Error Code 2: user aborted function
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchen.EinzelauftragSuchenDeleteRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "EAkurzKey"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Einzelauftrag
    Dim Einzelauftrag As clsEinzelauftrag
    Set Einzelauftrag = New clsEinzelauftrag
    
    ' check primary key value
    If IsNull(varRecordsetName) Then
        Debug.Print "Error: basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt. Speichern abgebrochen.", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' select recordset
    Einzelauftrag.SelectRecordset varRecordsetName
    
    ' delete recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Soll der Datensatz " & varRecordsetName & " wirklich gelöscht werden?", vbOKCancel, "Datensatz löschen")
    
    If varUserInput = 1 Then
        Einzelauftrag.DeleteRecordset
        MsgBox "Datensatz gelöscht", vbInformation, "Datensatz löschen"
    Else
        Debug.Print "Error: basEinzelauftragSuchen.AuftragSuchenDeleteRecordset aborted, Error Code 2"
        MsgBox "löschen abgebrochen", vbInformation, "Datensatz löschen"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchen.EinzelauftragSuchenSaveRecordset execute"
    End If
    
End Function

