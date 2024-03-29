Attribute VB_Name = "basAuftragUebersicht"
Option Compare Database
Option Explicit

Public Sub BuildAuftragUebersicht()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.BuildAuftragUebersicht"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragUebersicht"
    
    ' set subform name
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' clear form
    basAuftragUebersicht.ClearForm strFormName
     
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
        intNumberOfRows = 12
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
    ' calculate information grid
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 2
    intRow = 1
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![AftrID]"
            .BorderStyle = 0
    End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "Ticket ID"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![AftrTitel]"
            .BorderStyle = 0
        End With
        
    'lbl01
    intColumn = 1
    intRow = 2
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Zusammenfassung"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 2
    intRow = 3
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![BWIKey]"
            .BorderStyle = 0
        End With
        
    'lbl02
    intColumn = 1
    intRow = 3
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Angebot"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 2
    intRow = 4
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
            .ControlSource = "=[frbSubForm].[Form]![LeistungsbeschreibungLink]"
            .BorderStyle = 0
        End With
        
    'lbl03
    intColumn = 1
    intRow = 4
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Leistungsbeschreibung (Link)"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 2
    intRow = 5
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
            .ControlSource = "=[frbSubForm].[Form]![MengengeruestLink]"
            .BorderStyle = 0
        End With
        
    'lbl04
    intColumn = 1
    intRow = 5
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Mengengerüst (Link)"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 2
    intRow = 6
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![EaAngebot]"
            .BorderStyle = 0
        End With
        
    'lbl05
    intColumn = 1
    intRow = 6
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "Angeboten zum EA"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 2
    intRow = 7
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![BeauftragtDatum]"
            .Format = "Short Date"
            .BorderStyle = 0
        End With
        
    'lbl06
    intColumn = 1
    intRow = 7
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "Beauftragt am"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 2
    intRow = 8
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![AbgenommenDatum]"
            .Format = "Short Date"
            .BorderStyle = 0
        End With
        
    'lbl07
    intColumn = 1
    intRow = 8
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "Abgenommen am"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt08
    intColumn = 2
    intRow = 9
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![RechnungNr]"
            .BorderStyle = 0
        End With
        
    'lbl08
    intColumn = 1
    intRow = 9
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "RechnungsNr"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt09
    intColumn = 2
    intRow = 10
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![RechnungBrutto]"
            .Format = "Currency"
            .BorderStyle = 0
        End With
        
    'lbl09
    intColumn = 1
    intRow = 10
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "Rechnung (Brutto)"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt10
    intColumn = 2
    intRow = 11
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![EaRechnung]"
            .BorderStyle = 0
        End With
        
    'lbl10
    intColumn = 1
    intRow = 11
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "Abgerechnet im EA"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
   
    'txt11
    intColumn = 2
    intRow = 12
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
            .ControlSource = "=[frbSubForm].[Form]![LeistungserfassungsblattID]"
            .BorderStyle = 0
        End With
        
    'lbl11
    intColumn = 1
    intRow = 12
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "Leistungserfassungsblatt"
            .Left = basAuftragUebersicht.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragUebersicht.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragUebersicht.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragUebersicht.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' create lifecycle grid
    Dim aintLifecycleGrid() As Integer
    
        ' grid settings
        intNumberOfColumns = 6
        intNumberOfRows = 1
        intLeft = 510
        intTop = 1700
        intWidth = 2450
        intHeight = 330
        
        ReDim aintLifecycleGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        aintLifecycleGrid = basAuftragUebersicht.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Ticket erstellen" button
        intColumn = 1
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateTicket"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Ticket erstellen"
                .OnClick = "=OpenAuftragUebersichtTicketErstellen()"
                .Visible = True
            End With
            
        ' create "Angebot erstellen" button
        intColumn = 2
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateAngebot"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Angebot erstellen"
                .OnClick = "=OpenAuftragUebersichtAngebotErstellen()"
                .Visible = True
            End With
            
        ' create "Auftrag erteilen" button
        intColumn = 3
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdExecuteOrder"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Auftrag erteilen"
                .OnClick = "=OpenAuftragUebersichtAuftragErteilen()"
                .Visible = True
            End With
            
        ' create "Leistung abnehmen" button
        intColumn = 4
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdConfirmService"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Leistung abnehmen"
                .OnClick = "=OpenAuftragUebersichtLeistungAbnehmen()"
                .Visible = True
            End With
            
        ' create "Rechnung erstellen" button
        intColumn = 5
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateInvoice"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Rechnung erstellen"
                .OnClick = "=OpenAuftragUebersichtRechnungErstellen()"
                .Visible = True
            End With
            
        ' OpenAuftragUebersichtLeistungserfassungsblattErstellen
        ' create "Rechnung erstellen" button
        intColumn = 6
        intRow = 1
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateLeistungserfassungsblatt"
                .Left = basAuftragUebersicht.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAuftragUebersicht.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAuftragUebersicht.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAuftragUebersicht.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "LEB erstellen"
                .OnClick = "=OpenAuftragUebersichtLeistungserfassungsblattErstellen()"
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
            lblLabel.Caption = "Auftragsübersicht"
            
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
            btnButton.OnClick = "=SearchAndReloadAuftragUebersicht()"
            btnButton.Visible = True
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFormAuftragUebersicht()"
            
        ' create subform
        Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = strSubformName
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmAuftragUebersichtSub"
            .Locked = True
        End With
            
    ' close form and save changes
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.BuildAuftragUebersicht executed"
    End If

End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.ClearForm"
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
        Debug.Print "basAuftragUebersicht.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragUebersicht"
    
    ' delete form
    basAuftragUebersicht.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create empty form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' delete form
    basAuftragUebersicht.ClearForm strFormName
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            MsgBox "basAuftragUebersicht.TestClearForm: Test failed", vbCritical, "Test Result"
            Exit For
        End If
    Next
    
    MsgBox "basAuftragUebersicht.TestClearForm: Test succesfull", vbOKOnly, "Test Result"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestClearForm executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.CalculateGrid"
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
        Debug.Print "basAuftragUebersicht.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestCalculateGrid"
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
    
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basAuftragUebersicht.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basAuftragUebersicht.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basAuftragUebersicht.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basAuftragUebersicht.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragUebersicht.GetLeft: column 0 is not available"
        MsgBox "basAuftragUebersicht.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basAuftragUebersicht.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basAuftragUebersicht.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragUebersicht.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragUebersicht.GetTop: column 0 is not available"
        MsgBox "basAuftragUebersicht.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basAuftragUebersicht.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basAuftragUebersicht.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragUebersicht.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragUebersicht.TestGetHeight: column 0 is not available"
        MsgBox "basAuftragUebersicht.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basAuftragUebersicht.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basAuftragUebersicht.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragUebersicht.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basAuftragUebersicht.TestGetWidth: column 0 is not available"
        MsgBox "basAuftragUebersicht.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basAuftragUebersicht.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basAuftragUebersicht.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basAuftragUebersicht.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basAuftragUebersicht.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.TestGetWidth executed"
    End If
    
End Sub

Public Function SearchAndReloadAuftragUebersicht()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.SearchAndReloadAuftragUebersicht"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragUebersicht"
    
    Dim strSearchTextboxName As String
    strSearchTextboxName = "txtSearchBox"
    
    ' search AuftragUebersicht
    Dim strQueryName As String
    strQueryName = "qryAuftragUebersicht"
    
    Dim strQuerySource As String
    ' -> strQuerySource = "tblAuftragUebersicht"
    
    Dim strPrimaryKey As String
    ' -> strPrimaryKey = "EAkurzKey"
    
    Dim varSearchTerm As Variant
    varSearchTerm = Application.Forms.Item(strFormName).Controls(strSearchTextboxName)
    
    ' -> basAuftragUebersichtSub.SearchAuftragUebersicht strQueryName, strQuerySource, strPrimaryKey, varSearchTerm
    basAuftragUebersichtSub.SearchAuftragUebersicht strQueryName, varSearchTerm
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.SearchAndReloadAuftragUebersicht executed"
    End If
    
End Function

Public Function CloseFormAuftragUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.CloseForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragUebersicht"
    
    DoCmd.Close acForm, strFormName, acSaveYes
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.CloseForm executed"
    End If
    
End Function

Public Function OpenFormAuftragUebersichtErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenFormAuftragUebersichtErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragUebersichtErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenFormAuftragUebersichtErstellen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtTicketErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.OpenAuftragUebersichtTicketErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.OpenAuftragUebersichtTicketErstellen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtAngebotErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenAuftragUebersichtAngebotErstellen"
    End If
    
    ' name the opening form
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    ' name the current form
    Dim strFormNameClipboardSource As String
    strFormNameClipboardSource = "frmAuftragUebersicht"
    
    ' name the subform of the current form
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' name the attributes that will be use in the opening form
    Dim strExportField01 As String
    strExportField01 = "txt00"
    
    Dim strExportField02 As String
    strExportField02 = "txt01"
    
    ' reset frmAngebotErstellenClipboard
    gvarAngebotErstellenClipboardAftrID = Null
    gvarAngebotErstellenClipboardAftrTitel = Null
    
    ' send AftrID and AftrTitel to frmAngebotErstellen's Clipboard
    gvarAngebotErstellenClipboardAftrID = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strExportField01)
    gvarAngebotErstellenClipboardAftrTitel = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strExportField02)
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenAuftragUebersichtAngebotErstellen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtAuftragErteilen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenAuftragUebersichtAuftragErteilen"
    End If
    
    ' name the opening form
    Dim strFormName As String
    strFormName = "frmAuftragErteilen"
    
    ' name the current form
    Dim strFormNameClipboardSource As String
    strFormNameClipboardSource = "frmAuftragUebersicht"
    
    ' name the subform of the current form
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' name the attributes that will be used in the opening form
    
    ' select Angebot
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    
    Dim strFieldBWIKey As String
    strFieldBWIKey = "txt02"
    
    rstAngebot.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldBWIKey))
    
    ' select Ticket
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
        
    Dim strFieldAftrID As String
    strFieldAftrID = "txt00"
    
    rstTicket.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldAftrID))
    
    ' get Einzelauftrag
    Dim strFieldEAID As String
    strFieldEAID = "txt05"
    
    ' reset frmAuftragErteilenClipboard
    gvarAuftragErteilenClipboardAftrID = Null
    gvarAuftragErteilenClipboardBWIKey = Null
    gvarAuftragErteilenClipboardEinzelauftrag = Null
    
    ' send AftrID and AftrTitel to frmAuftragErteilen's Clipboard
    gvarAuftragErteilenClipboardAftrID = rstTicket.AftrID
    gvarAuftragErteilenClipboardBWIKey = rstAngebot.BWIKey
    gvarAuftragErteilenClipboardEinzelauftrag = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldEAID)
    
    If IsNull(gvarAuftragErteilenClipboardBWIKey) Then
        MsgBox "Zu dem ausgewählten Datensatz wurde bisher kein Angebot erfasst.", vbCritical, "Angebot erfassen"
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtAuftragErteilen, no offer to recordset"
        Exit Function
    End If
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenAuftragUebersichtAuftragErteilen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtLeistungAbnehmen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen"
    End If
    
    ' name the opening form
    Dim strFormName As String
    strFormName = "frmLeistungAbnehmen"
    
    ' name the current form
    Dim strFormNameClipboardSource As String
    strFormNameClipboardSource = "frmAuftragUebersicht"
    
    ' name the subform of the current form
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' select Ticket
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
        
    Dim strFieldAftrID As String
    strFieldAftrID = "txt00"
    
    rstTicket.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldAftrID))
    
    ' select Angebot
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    
    Dim strFieldBWIKey As String
    strFieldBWIKey = "txt02"
    
    rstAngebot.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldBWIKey))
    
    ' reset frmLeistungAbnehmenClipboard
    gvarLeistungAbnehmenClipboardAftrID = Null
    gvarLeistungAbnehmenClipboardBWIKey = Null
    
    ' send AftrID and AftrTitel to frmLeistungAbnehmen's Clipboard
    gvarLeistungAbnehmenClipboardAftrID = rstTicket.AftrID
    gvarLeistungAbnehmenClipboardBWIKey = rstAngebot.BWIKey
    
    ' check BWIKey
    If IsNull(gvarLeistungAbnehmenClipboardBWIKey) Then
        MsgBox "Zu dem ausgewählten Datensatz wurde bisher kein Angebot erfasst.", vbCritical, "Angebot erfassen"
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen, no offer to recordset"
        Exit Function
    End If
    
    ' check BeauftragtDatum
    If IsNull(rstAngebot.BeauftragtDatum) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Das ausgewählte Angebot wurde bisher nicht beauftragt.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtRechnungErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenAuftragUebersichtRechnungErstellen"
    End If
    
    ' name the opening form
    Dim strFormName As String
    strFormName = "frmRechnungErstellen"
    
    ' name the current form
    Dim strFormNameClipboardSource As String
    strFormNameClipboardSource = "frmAuftragUebersicht"
    
    ' name the subform of the current form
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' name the attributes that will be used in the opening form
    ' get TicketID
    Dim rstTicket As clsAuftrag
    Set rstTicket = New clsAuftrag
    
    Dim strFieldAftrID As String
    strFieldAftrID = "txt00"
    
    rstTicket.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldAftrID))
    
    ' get AngebotID
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    
    Dim strFieldBWIKey As String
    strFieldBWIKey = "txt02"
    rstAngebot.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldBWIKey))
    
    ' get Einzelauftrag
    Dim strFieldEAIDAngebot As String
    strFieldEAIDAngebot = "txt05"
    
    Dim strFieldEAIDRechnung As String
    strFieldEAIDRechnung = "txt10"
    
    ' get Rechnung
    Dim strFieldRechnungNr As String
    strFieldRechnungNr = "txt08"
    
    ' reset frmRechnungErstellenClipboard
    gvarRechnungErstellenClipboardAftrID = Null
    gvarRechnungErstellenClipboardBWIKey = Null
    gvarRechnungErstellenClipboardEAIDAngebot = Null
    gvarRechnungErstellenClipboardRechnungNr = Null
    gvarRechnungErstellenClipboardEAIDRechnung = Null
    
    ' send AftrID and AftrTitel to frmRechnungErstellen's Clipboard
    gvarRechnungErstellenClipboardAftrID = rstTicket.AftrID
    gvarRechnungErstellenClipboardBWIKey = rstAngebot.BWIKey
    gvarRechnungErstellenClipboardEAIDAngebot = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldEAIDAngebot)
    gvarRechnungErstellenClipboardRechnungNr = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldRechnungNr)
    gvarRechnungErstellenClipboardEAIDRechnung = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldEAIDRechnung)
    
    ' check BWIKey
    If IsNull(gvarRechnungErstellenClipboardBWIKey) Then
        MsgBox "Zu dem ausgewählten Datensatz wurde bisher kein Angebot erfasst.", vbCritical, "Angebot erfassen"
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen, no offer to recordset"
        Exit Function
    End If
    
    ' check BeauftragtDatum
    If IsNull(rstAngebot.BeauftragtDatum) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Das ausgewählte Angebot wurde bisher nicht beauftragt.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' check AbgenommenDatum
    If IsNull(rstAngebot.AbgenommenDatum) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Die Leistung zu dem ausgewählten Angebot wurde bisher nicht abgenommen.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenAuftragUebersichtRechnungErstellen executed"
    End If
    
End Function

Public Function OpenAuftragUebersichtLeistungserfassungsblattErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragUebersicht.OpenAuftragUebersichtLeistungserfassungsblattErstellen"
    End If
    
    ' name the opening form
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattErstellen"
    
    ' name the current form
    Dim strFormNameClipboardSource As String
    strFormNameClipboardSource = "frmAuftragUebersicht"
    
    ' name the subform of the current form
    Dim strSubformName As String
    strSubformName = "frbSubForm"
    
    ' name the attributes that will be used in the to be opened form
    ' get TicketID
    ' Dim rstTicket As clsAuftrag
    ' Set rstTicket = New clsAuftrag
    
    Dim strFieldAftrID As String
    strFieldAftrID = "txt00"
    
    ' rstTicket.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldAftrID))
    
    ' get AngebotID
    Dim rstAngebot As clsAngebot
    Set rstAngebot = New clsAngebot
    
    Dim strFieldBWIKey As String
    strFieldBWIKey = "txt02"
    rstAngebot.SelectRecordset (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldBWIKey))
    
    ' get Einzelauftrag
    Dim strFieldEAIDAngebot As String
    strFieldEAIDAngebot = "txt05"
    
    Dim strFieldEAIDRechnung As String
    strFieldEAIDRechnung = "txt10"
    
    ' get Rechnung
    Dim strFieldRechnungNr As String
    strFieldRechnungNr = "txt08"
    
    ' get Leistungserfassungsblatt
    Dim strFieldLeistungserfassungsblattID As String
    strFieldLeistungserfassungsblattID = "txt11"
    
    ' get RechnungZuLeistungserfassungsblatt
    Dim strFieldRechnungZuLeistungserfassungsblattID As String
    strFieldRechnungZuLeistungserfassungsblattID = "txt16"
    
    ' reset frmLeistungserfassungsblattErstellenClipboard
    gvarLeistungserfassungsblattErstellenClipboardAftrID = Null
    gvarLeistungserfassungsblattErstellenClipboardBWIKey = Null
    gvarLeistungserfassungsblattErstellenClipboardEAIDAngebot = Null
    gvarLeistungserfassungsblattErstellenClipboardRechnungNr = Null
    gvarLeistungserfassungsblattErstellenClipboardEAIDRechnung = Null
    gvarLeistungserfassungsblattErstellenClipboardLeistungserfassungsblattID = Null
    
    ' send AftrID and AftrTitel to frmLeistungserfassungsblattErstellen's Clipboard
    gvarLeistungserfassungsblattErstellenClipboardAftrID = (Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldAftrID))
    gvarLeistungserfassungsblattErstellenClipboardBWIKey = rstAngebot.BWIKey
    gvarLeistungserfassungsblattErstellenClipboardEAIDAngebot = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldEAIDAngebot)
    gvarLeistungserfassungsblattErstellenClipboardRechnungNr = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldRechnungNr)
    gvarLeistungserfassungsblattErstellenClipboardEAIDRechnung = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldEAIDRechnung)
    gvarLeistungserfassungsblattErstellenClipboardLeistungserfassungsblattID = Forms.Item(strFormNameClipboardSource).Controls(strSubformName).Form(strFieldLeistungserfassungsblattID)
    
    ' check BWIKey
    If IsNull(gvarLeistungserfassungsblattErstellenClipboardBWIKey) Then
        MsgBox "Zu dem ausgewählten Datensatz wurde bisher kein Angebot erfasst.", vbCritical, "Angebot erfassen"
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen, no offer to recordset"
        Exit Function
    End If
    
    ' check BeauftragtDatum
    If IsNull(rstAngebot.BeauftragtDatum) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Das ausgewählte Angebot wurde bisher nicht beauftragt.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' check AbgenommenDatum
    If IsNull(rstAngebot.AbgenommenDatum) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Die Leistung zu dem ausgewählten Angebot wurde bisher nicht abgenommen.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' check RechnungNr
    If IsNull(gvarLeistungserfassungsblattErstellenClipboardRechnungNr) Then
        Debug.Print "Error: basAuftragUebersicht.OpenAuftragUebersichtLeistungAbnehmen canceled, Error Code 1"
        MsgBox "Zu dem ausgewählten Auftrag wurde keine Rechnung erfasst.", vbCritical, "Speichern"
        Exit Function
    End If
    
    ' open form
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAuftragUebersicht.OpenAuftragUebersichtLeistungserfassungsblattErstellen executed"
    End If
    
End Function

