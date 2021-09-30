Attribute VB_Name = "basHauptmenue"
' basHauptmenue

Option Compare Database
Option Explicit

Public Const gconVerbatim As Boolean = True

Public Sub BuildHauptmenue()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.BuildFormHauptmenue"
    End If
    
    ' define form name
    Dim strFormName As String
    strFormName = "frmHauptmenue"
    
    ' clear form
    basHauptmenue.ClearForm strFormName
    
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
    
    ' declare labels
    Dim lblLabel As Label
    
    ' declare grid variables
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intWidth As Integer
        Dim intHeight As Integer
        
        Dim intColumn As Integer
        Dim intRow As Integer
    
    ' create control grid
    Dim aintControlGrid() As Integer
    
        ' grid settings
        intNumberOfColumns = 4
        intNumberOfRows = 10
        intLeft = 100
        intTop = 100
        intWidth = 3800
        intHeight = 660
    
    ReDim aintControlGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    ' calculate control grid
    aintControlGrid = basHauptmenue.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
    ' define layout
    Dim intNumberOfElements As Integer
    intNumberOfElements = 22

    Dim avarLayout As Variant
    ReDim avarLayout(intNumberOfElements, 3)
    
    avarLayout(0, 0) = ""  ' caption
        avarLayout(0, 1) = 0  ' column
        avarLayout(0, 2) = 0  ' row
        avarLayout(0, 3) = ""  ' function
    avarLayout(19, 0) = "Übersicht"
        avarLayout(19, 1) = 1
        avarLayout(19, 2) = 1
        avarLayout(19, 3) = ""
    ' ---
    avarLayout(1, 0) = "Ticketübersicht"
        avarLayout(1, 1) = 1
        avarLayout(1, 2) = 2
        avarLayout(1, 3) = "=" & "OpenFormAuftragUebersicht" & "()"
    avarLayout(2, 0) = "Liefergegenstand" & " " & "Übersicht"
        avarLayout(2, 1) = 2
        avarLayout(2, 2) = 2
        avarLayout(2, 3) = "=" & "OpenFormLiefergegenstandUebersicht" & "()"
    avarLayout(3, 0) = "Einzelauftrag" & " " & "Übersicht"
        avarLayout(3, 1) = 3
        avarLayout(3, 2) = 2
        avarLayout(3, 3) = "=" & "OpenFormEinzelauftragUebersicht" & "()"
    ' ---
    avarLayout(20, 0) = "Bearbeiten"
        avarLayout(20, 1) = 1
        avarLayout(20, 2) = 3
        avarLayout(20, 3) = ""
    ' ---
    avarLayout(4, 0) = "Ticket" & " " & "bearbeiten"
        avarLayout(4, 1) = 1
        avarLayout(4, 2) = 4
        avarLayout(4, 3) = "=" & "OpenFormAuftragSuchen" & "()"
    avarLayout(5, 0) = "Angebot" & " " & "bearbeiten"
        avarLayout(5, 1) = 2
        avarLayout(5, 2) = 4
        avarLayout(5, 3) = "=" & "OpenFormAngebotSuchen" & "()"
    avarLayout(6, 0) = "Rechnung" & " " & "bearbeiten"
        avarLayout(6, 1) = 3
        avarLayout(6, 2) = 4
        avarLayout(6, 3) = "=" & "OpenFormRechnungSuchen" & "()"
    avarLayout(7, 0) = "Leistungserfassungsblatt" & " " & "bearbeiten"
        avarLayout(7, 1) = 4
        avarLayout(7, 2) = 4
        avarLayout(7, 3) = "=" & "OpenFormLeistungserfassungsblattSuchen" & "()"
    ' ---
    avarLayout(8, 0) = "Liefergegenstand" & " " & "bearbeiten"
        avarLayout(8, 1) = 1
        avarLayout(8, 2) = 5
        avarLayout(8, 3) = "=" & "OpenFormLiefergegenstandSuchen" & "()"
    avarLayout(9, 0) = "Einzelauftrag" & " " & "bearbeiten"
        avarLayout(9, 1) = 2
        avarLayout(9, 2) = 5
        avarLayout(9, 3) = "=" & "OpenFormEinzelauftragSuchen" & "()"
    avarLayout(10, 0) = "Kontinuierliche" & " " & "Leistungen" & " " & "berbeiten"
        avarLayout(10, 1) = 3
        avarLayout(10, 2) = 5
        avarLayout(10, 3) = "=" & "OpenFormKontinuierlicheLeistungenSuchen" & "()"
    ' ---
    avarLayout(21, 0) = "Beziehungen verwalten"
        avarLayout(21, 1) = 1
        avarLayout(21, 2) = 6
        avarLayout(21, 3) = ""
    ' ---
    avarLayout(11, 0) = "Ticket - Angebot verwalten"
        avarLayout(11, 1) = 1
        avarLayout(11, 2) = 7
        avarLayout(11, 3) = "=" & "OpenFormAuftragZuAngebotVerwalten" & "()"
    avarLayout(12, 0) = "Angebot - Rechnung verwalten"
        avarLayout(12, 1) = 2
        avarLayout(12, 2) = 7
        avarLayout(12, 3) = "=" & "OpenFormAngebotZuRechnungVerwalten" & "()"
    avarLayout(13, 0) = "Rechnung - Leistungserfassungsblatt verwalten"
        avarLayout(13, 1) = 3
        avarLayout(13, 2) = 7
        avarLayout(13, 3) = "=" & "OpenFormRechnungZuLeistungserfassungsblattVerwalten" & "()"
    ' ---
    avarLayout(14, 0) = "Angebot - Liefergegenstand verwalten"
        avarLayout(14, 1) = 1
        avarLayout(14, 2) = 8
        avarLayout(14, 3) = "=" & "OpenFormAngebotZuLiefergegenstandVerwalten" & "()"
    avarLayout(15, 0) = "Einzelauftrag - Angebot verwalten"
        avarLayout(15, 1) = 2
        avarLayout(15, 2) = 8
        avarLayout(15, 3) = "=" & "OpenFormEinzelauftragZuAngebotVerwalten" & "()"
    avarLayout(16, 0) = "Einzelauftrag - Rechnung verwalten"
        avarLayout(16, 1) = 3
        avarLayout(16, 2) = 8
        avarLayout(16, 3) = "=" & "OpenFormEinzelauftragZuRechnungVerwalten" & "()"
    avarLayout(17, 0) = "Kontinuierliche Leistungen - Rechnung verwalten"
        avarLayout(17, 1) = 4
        avarLayout(17, 2) = 8
        avarLayout(17, 3) = "=" & "OpenFormKontinuierlicheLeistungenZuRechnungVerwalten" & "()"
    ' ---
    avarLayout(22, 0) = "Anwendung laden"
        avarLayout(22, 1) = 1
        avarLayout(22, 2) = 9
        avarLayout(22, 3) = ""
    ' ---
    avarLayout(18, 0) = "Build Application"
        avarLayout(18, 1) = 1
        avarLayout(18, 2) = 10
        avarLayout(18, 3) = "=" & "BuildApplication" & "()"
           
    Dim strValueWanted As String
    
    ' create control elements
    strValueWanted = "Übersicht"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
        With lblLabel
            .Name = "lbl01"
            .Visible = True
            .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
            .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
            .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
            .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
            .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
        End With
        
    strValueWanted = "Bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
        With lblLabel
            .Name = "lbl02"
            .Visible = True
            .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
            .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
            .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
            .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
            .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
        End With
        
    strValueWanted = "Beziehungen verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
        With lblLabel
            .Name = "lbl03"
            .Visible = True
            .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
            .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
            .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
            .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
            .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
        End With
        
    strValueWanted = "Anwendung laden"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
        With lblLabel
            .Name = "lbl04"
            .Visible = True
            .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
            .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
            .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
            .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
            .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
        End With
    
    strValueWanted = "Ticketübersicht"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd00"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Liefergegenstand" & " " & "Übersicht"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd01"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Einzelauftrag" & " " & "Übersicht"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd02"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Ticket" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd03"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Angebot" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd04"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Rechnung" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd05"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Leistungserfassungsblatt" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd06"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Liefergegenstand" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd07"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Einzelauftrag" & " " & "bearbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd08"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Kontinuierliche" & " " & "Leistungen" & " " & "berbeiten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd09"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Ticket - Angebot verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd10"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Angebot - Rechnung verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd11"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Rechnung - Leistungserfassungsblatt verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd12"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Angebot - Liefergegenstand verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd13"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Einzelauftrag - Angebot verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd14"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Einzelauftrag - Rechnung verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd15"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Kontinuierliche Leistungen - Rechnung verwalten"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd16"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    
    strValueWanted = "Build Application"
    intColumn = ReturnValueByName(avarLayout, strValueWanted, "column")
    intRow = ReturnValueByName(avarLayout, strValueWanted, "row")
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd17"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = ReturnValueByName(avarLayout, strValueWanted, "name")
                .OnClick = ReturnValueByName(avarLayout, strValueWanted, "function")
                .Visible = True
            End With
    ' column added? -> update intNumberOfColumns
            
        ' close form
        DoCmd.Close acForm, strTempFormName, acSaveYes
    
        ' rename form
        DoCmd.Rename strFormName, acForm, strTempFormName
        
        ' open form
        DoCmd.OpenForm strFormName, acNormal
        
        ' event message
        If gconVerbatim Then
            Debug.Print "basHauptmenue.BuildHauptmenue executed"
        End If
    
End Sub

Public Function OpenFormAuftragSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAuftragSuchen"
    End If
    
    DoCmd.OpenForm "frmAuftragSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAuftragSuchen executed"
    End If
    
End Function

Public Function OpenFormAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAngebotSuchen"
    End If

    DoCmd.OpenForm "frmAngebotSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAngebotSuchen executed"
    End If
End Function

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.CalculateGrid"
    End If
    
    Const cintHorizontalSpacing As Integer = 60
    Const cintVerticalSpacing As Integer = 80
    
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
        Debug.Print "basHauptmenue.CalculateGrid executed"
    End If
    
End Function

' get left from grid
Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basHauptmenue.GetLeft: column 0 is not available"
        MsgBox "basHauptmenue.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.GetLeft executed"
    End If
    
End Function

' get left from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basHauptmenue.GetTop: column 0 is not available"
        MsgBox "basHauptmenue.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.GetTop executed"
    End If
    
End Function

' get left from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basHauptmenue.GetWidth: column 0 is not available"
        MsgBox "basHauptmenue.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.GetWidth executed"
    End If
    
End Function

' get left from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basHauptmenue.GetHeight: column 0 is not available"
        MsgBox "basHauptmenue.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.GetHeight executed"
    End If
    
End Function

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.ClearForm"
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
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenu.ClearForm executed"
    End If
    
End Sub

Public Function OpenFormRechnungSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormRechnungSuchen"
    End If

    DoCmd.OpenForm "frmRechnungSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormRechnungSuchen executed"
    End If
End Function

Public Function OpenFormLeistungserfassungsblattSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormLeistungserfassungsblattSuchen"
    End If

    DoCmd.OpenForm "frmLeistungserfassungsblattSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormLeistungserfassungsblattSuchen executed"
    End If
End Function

Public Function OpenFormLiefergegenstandSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormLiefergegenstandSuchen"
    End If

    DoCmd.OpenForm "frmLiefergegenstandSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormLiefergegenstandSuchen executed"
    End If
End Function

Public Function OpenFormEinzelauftragSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormEinzelauftragSuchen"
    End If

    DoCmd.OpenForm "frmEinzelauftragSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormEinzelauftragSuchen executed"
    End If
End Function

Public Function OpenFormKontinuierlicheLeistungenSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormKontinuierlicheLeistungenSuchen"
    End If

    DoCmd.OpenForm "frmKontinuierlicheLeistungenSuchen", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormKontinuierlicheLeistungenSuchen executed"
    End If
End Function


Public Function OpenFormAuftragUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAuftragUebersicht"
    End If

    DoCmd.OpenForm "frmAuftragUebersicht", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAuftragUebersicht executed"
    End If
End Function

Public Function OpenFormAuftragZuAngebotVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAuftragZuAngebotVerwalten"
    End If

    DoCmd.OpenForm "frmAuftragZuAngebotVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAuftragZuAngebotVerwalten executed"
    End If
End Function

Public Function OpenFormEinzelauftragZuAngebotVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormEinzelauftragZuAngebotVerwalten"
    End If

    DoCmd.OpenForm "frmEinzelauftragZuAngebotVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormEinzelauftragZuAngebotVerwalten executed"
    End If
End Function


Public Function OpenFormAngebotZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAngebotZuRechnungVerwalten"
    End If

    DoCmd.OpenForm "frmAngebotZuRechnungVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAngebotZuRechnungVerwalten executed"
    End If
End Function

Public Function OpenFormEinzelauftragZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormEinzelauftragZuRechnungVerwalten"
    End If

    DoCmd.OpenForm "frmEinzelauftragZuRechnungVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormEinzelauftragZuRechnungVerwalten executed"
    End If
End Function


Public Function OpenFormAngebotZuLiefergegenstandVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormAngebotZuLiefergegenstandVerwalten"
    End If

    DoCmd.OpenForm "frmAngebotZuLiefergegenstandVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormAngebotZuLiefergegenstandVerwalten executed"
    End If
End Function

Public Function OpenFormRechnungZuLeistungserfassungsblattVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormRechnungZuLeistungserfassungsblattVerwalten"
    End If

    DoCmd.OpenForm "frmRechnungZuLeistungserfassungsblattVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormRechnungZuLeistungserfassungsblattVerwalten executed"
    End If
End Function

Public Function OpenFormKontinuierlicheLeistungenZuRechnungVerwalten()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormKontinuierlicheLeistungenZuRechnungVerwalten"
    End If

    DoCmd.OpenForm "frmKontinuierlicheLeistungenZuRechnungVerwalten", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormKontinuierlicheLeistungenZuRechnungVerwalten executed"
    End If
End Function

Public Function OpenFormLiefergegenstandUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormLiefergegenstandUebersicht"
    End If

    DoCmd.OpenForm "frmLiefergegenstandUebersicht", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormLiefergegenstandUebersicht executed"
    End If
End Function

Public Function OpenFormEinzelauftragUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.OpenFormEinzelauftragUebersicht"
    End If

    DoCmd.OpenForm "frmEinzelauftragUebersicht", acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.OpenFormEinzelauftragUebersicht executed"
    End If
End Function


' builds the application form scratch
' work in progress
Public Function BuildApplication()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basHauptmenue.BuildApplication"
    End If
        
    ' build forms
    basAngebotSuchenSub.BuildAngebotSuchenSub
    basAngebotSuchen.BuildAngebotSuchen
    basAngebotErstellen.buildAngebotErstellen
    
    basAuftragSuchenSub.BuildAuftragSuchenSub
    basAuftragSuchen.BuildAuftragSuchen
    basAngebotErstellen.buildAngebotErstellen
    
    basRechnungSuchenSub.BuildRechnungSuchenSub
    basRechnungSuchen.BuildRechnungSuchen
    basRechnungErstellen.buildRechnungErstellen
    
    basLeistungserfassungsblattSuchenSub.BuildLeistungserfassungsblattSuchenSub
    basLeistungserfassungsblattSuchen.BuildLeistungserfassungsblattSuchen
    basLeistungserfassungsblattErstellen.buildLeistungserfassungsblattErstellen
    
    basLiefergegenstandSuchenSub.BuildLiefergegenstandSuchenSub
    basLiefergegenstandSuchen.BuildLiefergegenstandSuchen
    basLiefergegenstandErstellen.buildLiefergegenstandErstellen
    
    basEinzelauftragSuchenSub.BuildEinzelauftragSuchenSub
    basEinzelauftragSuchen.BuildEinzelauftragSuchen
    basEinzelauftragErstellen.buildEinzelauftragErstellen
    
    basKontinuierlicheLeistungenSuchenSub.BuildKontinuierlicheLeistungenSuchenSub
    basKontinuierlicheLeistungenSuchen.BuildKontinuierlicheLeistungenSuchen
    basKontinuierlicheLeistungenErstellen.buildKontinuierlicheLeistungenErstellen
    
    basAuftragUebersichtSub.BuildAuftragUebersichtSub
    basAuftragUebersicht.BuildAuftragUebersicht
    
    basAuftragZuAngebotVerwaltenSub.BuildAuftragZuAngebotVerwaltenSub
    basAuftragZuAngebotVerwalten.BuildAuftragZuAngebotVerwalten
    
    basEinzelauftragZuAngebotVerwaltenSub.BuildEinzelauftragZuAngebotVerwaltenSub
    basEinzelauftragZuAngebotVerwalten.BuildEinzelauftragZuAngebotVerwalten
    
    basAngebotZuRechnungVerwaltenSub.BuildAngebotZuRechnungVerwaltenSub
    basAngebotZuRechnungVerwalten.BuildAngebotZuRechnungVerwalten
    
    basEinzelauftragZuRechnungVerwaltenSub.BuildEinzelauftragZuRechnungVerwaltenSub
    basEinzelauftragZuRechnungVerwalten.BuildEinzelauftragZuRechnungVerwalten
    
    basAngebotZuLiefergegenstandVerwaltenSub.buildAngebotZuLiefergegenstandVerwaltenSub
    basAngebotZuLiefergegenstandVerwalten.BuildEinzelauftragZuRechnungVerwalten
    
    basRechnungZuLeistungserfassungsblattVerwaltenSub.buildRechnungZuLeistungserfassungsblattVerwaltenSub
    basRechnungZuLeistungserfassungsblattVerwalten.BuildRechnungZuLeistungserfassungsblattVerwalten
    
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.BuildKontinuierlicheLeistungenZurRechnungVerwaltenSub
    basKontinuierlicheLeistungenZuRechnungVerwalten.BuildKontinuierlicheLeistungenZuRechnungVerwalten
    
    basLiefergegenstandUebersichtSub.BuildLiefergegenstandUebersichtSub
    basLiefergegenstandUebersicht.BuildLiefergegenstandUebersicht
    
    basEinzelauftragUebersichtSub.BuildEinzelauftragUebersichtSub
    basEinzelauftragUebersicht.BuildEinzelauftragUebersicht
    
    basAuftragErstellen.buildAuftragErstellen
    basAngebotErstellen.buildAngebotErstellen
    basAuftragErteilen.buildAuftragErteilen
    
    ' open frmHauptmenue
    DoCmd.OpenForm "frmHauptmenue", acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.BuildApplication executed"
    End If
    
End Function

' search varWanted in two dimensional array
' avarArray style: name, column, row, alias
' array style A: (intIndex, strField)
' array style B: (strField, intIndex)
' strField feasible values: caption, column, row, function
Private Function ReturnValueByName(ByVal avarArray, varWanted As Variant, ByVal strField As String) As Variant
    
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute basHauptmenue.ReturnValueByName"
    End If
   
    Dim intIndex As Integer
    intIndex = 0
    
    Dim intValue As Integer
    intValue = 0
    
    Select Case strField
        Case "caption"
            intValue = 0
        Case "column"
            intValue = 1
        Case "row"
            intValue = 2
        Case "function"
            intValue = 3
    End Select
    
    ' scan array until match
    Do While avarArray(intIndex, 0) <> varWanted
        If intIndex = UBound(avarArray, 1) Then
            Debug.Print "basHauptmenue.ReturnValueByName: '" & varWanted & "' im übergebenen Array nicht gefunden"
            ReturnValueByName = Null
            Exit Function
        Else
            intIndex = intIndex + 1
        End If
    Loop
    
    ' return value
    ReturnValueByName = avarArray(intIndex, intValue)
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "basHauptmenue.ReturnValueByName executed"
    End If
    
End Function

Private Sub TestReturnValueByName()

    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 3

    Dim varTestArray As Variant
    ReDim varTestArray(intNumberOfAttributes, 3)

    varTestArray(0, 0) = "caption"
        varTestArray(0, 1) = "column"
        varTestArray(0, 2) = "row"
        varTestArray(0, 3) = "function"
    varTestArray(1, 0) = "Auftrag Suchen"
        varTestArray(1, 1) = 1
        varTestArray(1, 2) = 1
        varTestArray(1, 3) = "=OpenFormAuftragSuchen()"
    varTestArray(2, 0) = "Angebot Suchen"
        varTestArray(2, 1) = 1
        varTestArray(2, 2) = 2
        varTestArray(2, 3) = "=OpenFormAngebotSuchen()"
    varTestArray(3, 0) = "Rechnung Suchen"
        varTestArray(3, 1) = 2
        varTestArray(3, 2) = 3
        varTestArray(3, 3) = "=OpenFormRechnungSuchen()"
        
    Dim varReturnedValue01 As Variant
    varReturnedValue01 = ReturnValueByName(varTestArray, "Rechnung Suchen", "caption")
    
    Dim varReturnedValue02 As Variant
    varReturnedValue02 = ReturnValueByName(varTestArray, "Rechnung Suchen", "column")
    
    Dim varReturnedValue03 As Variant
    varReturnedValue03 = ReturnValueByName(varTestArray, "Rechnung Suchen", "row")
    
    Dim varReturnedValue04 As Variant
    varReturnedValue04 = ReturnValueByName(varTestArray, "Rechnung Suchen", "function")
    
    Dim varExpectedValue01 As Variant
    varExpectedValue01 = varTestArray(3, 0)
    
    Dim varExpectedValue02 As Variant
    varExpectedValue02 = varTestArray(3, 1)
    
    Dim varExpectedValue03 As Variant
    varExpectedValue03 = varTestArray(3, 2)
    
    Dim varExpectedValue04 As Variant
    varExpectedValue04 = varTestArray(3, 3)
    
    If varReturnedValue01 = varExpectedValue01 And varReturnedValue02 = varExpectedValue02 And varReturnedValue03 = varExpectedValue03 And varReturnedValue04 = varReturnedValue04 Then
        MsgBox "basHauptmenue.TestReturnValueByName passed", vbOKOnly, "Test Result"
    Else
        MsgBox "basHauptmenue.TestReturnValueByName failed", vbCritical, "Test Result"
    End If
    
End Sub

