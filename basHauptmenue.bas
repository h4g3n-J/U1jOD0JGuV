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
        intNumberOfRows = 6
        intLeft = 100
        intTop = 100
        intWidth = 2600
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

