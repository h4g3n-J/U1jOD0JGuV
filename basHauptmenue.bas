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
        intNumberOfColumns = 1
        intNumberOfRows = 17
        intLeft = 100
        intTop = 100
        intWidth = 3800
        intHeight = 660
    
    ReDim aintControlGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    ' calculate control grid
    aintControlGrid = basHauptmenue.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
    intColumn = 1
    intRow = 1
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd00"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Auftrag Suchen"
                .OnClick = "=OpenFormAuftragSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 2
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd01"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Angebot Suchen"
                .OnClick = "=OpenFormAngebotSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 3
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd02"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Rechnung Suchen"
                .OnClick = "=OpenFormRechnungSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 4
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd03"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Leistungserfassungsblatt Suchen"
                .OnClick = "=OpenFormLeistungserfassungsblattSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 5
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd04"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Liefergegenstand suchen"
                .OnClick = "=OpenFormLiefergegenstandSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 6
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd05"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Einzelauftrag suchen"
                .OnClick = "=OpenFormEinzelauftragSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 7
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd06"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Auftragsübersicht"
                .OnClick = "=OpenFormAuftragUebersicht()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 8
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd07"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Auftrag - Angebot" & vbCrLf & " Beziehungen verwalten"
                .OnClick = "=OpenFormAuftragZuAngebotVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 9
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd08"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Einzelauftrag - Angebot" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormEinzelauftragZuAngebotVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 10
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd09"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Angebot - Rechnung" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormAngebotZuRechnungVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 11
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd10"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Einzelauftrag - Rechnung" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormEinzelauftragZuRechnungVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 12
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd11"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Angebot - Liefergegenstand" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormAngebotZuLiefergegenstandVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 13
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd12"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Rechnung - Leistungserfassungsblatt" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormRechnungZuLeistungserfassungsblattVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 14
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd13"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Kontinuierliche Leistungen suchen"
                .OnClick = "=OpenFormKontinuierlicheLeistungenSuchen()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 15
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd14"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Kontinuierliche Leistungen - Rechnung" & vbCrLf & " Beziehung verwalten"
                .OnClick = "=OpenFormKontinuierlicheLeistungenZuRechnungVerwalten()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 16
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd15"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Liefergegenstand Übersicht"
                .OnClick = "=OpenFormLiefergegenstandUebersicht()"
                .Visible = True
            End With
            
    intColumn = 1
    intRow = 17
    Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmd16"
                .Left = basHauptmenue.GetLeft(aintControlGrid, intColumn, intRow)
                .Top = basHauptmenue.GetTop(aintControlGrid, intColumn, intRow)
                .Width = basHauptmenue.GetWidth(aintControlGrid, intColumn, intRow)
                .Height = basHauptmenue.GetHeight(aintControlGrid, intColumn, intRow)
                .Caption = "Build Application"
                .OnClick = "=BuildApplication()"
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
    
    ' open frmHauptmenue
    DoCmd.OpenForm "frmHauptmenue", acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basHauptmenue.BuildApplication executed"
    End If
    
End Function


