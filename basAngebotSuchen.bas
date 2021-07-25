Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

' build form AngebotSuchen
Public Sub BuildAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen ausführen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' declare temporary form name
    Dim strTempFormName As String

    ' clear form
    basAngebotSuchen.ClearForm strFormName

    ' declare form
    Dim frm As Form
    Set frm = CreateForm
    
    ' write temporary form name to strFormName
    strTempFormName = frm.Name
    
    ' set form caption
    frm.Caption = strFormName
        
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
        intNumberOfRows = 6
        intLeft = 10000
        intTop = 2430
        'intColumnWidth(0) = 2540
        'intColumnWidth(1) = 3120
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basAngebotSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
            
        ' txt00
        intColumn = 2
        intRow = 1
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt00"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl00
        intColumn = 1
        intRow = 1
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
            With lblLabel
                .Name = "lbl00"
                .Caption = "Angebot"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
        
        ' lbl01
        intColumn = 1
        intRow = 2
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
            With lblLabel
                .Name = "lbl01"
                .Caption = "Einzelauftrag"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt01
        intColumn = 2
        intRow = 2
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt01"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl02
        intColumn = 1
        intRow = 3
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
            With lblLabel
                .Name = "lbl02"
                .Caption = "Mengengerüst"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt02
        intColumn = 2
        intRow = 3
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt02"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
            End With
            
        ' lbl03
        intColumn = 1
        intRow = 4
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
            With lblLabel
                .Name = "lbl03"
                .Caption = "Leistungsbeschreibung"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt03
        intColumn = 2
        intRow = 4
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt03"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = True
            End With
            
        ' lbl04
        intColumn = 1
        intRow = 5
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
            With lblLabel
                .Name = "lbl04"
                .Caption = "Bemerkung"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt04
        intColumn = 2
        intRow = 5
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt04"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl05
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "Wildcard"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt05
        intColumn = 2
        intRow = 6
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt05"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
        
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
        
        aintLifecycleGrid = basAngebotSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
    
        ' create "Angebot erstellen" button
        intColumn = 1
        intRow = 1
        
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
            With btnButton
                .Name = "cmdCreateOffer"
                .Left = basAngebotSuchen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Angebot erstellen"
                .OnClick = "=OpenFormCreateOffer()"
            End With
            
        ' create form title
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail)
            lblLabel.Name = "lblTitle"
            lblLabel.Visible = True
            lblLabel.Left = 566
            lblLabel.Top = 227
            lblLabel.Width = 9210
            lblLabel.Height = 507
            lblLabel.Caption = "Angebot Suchen"
    
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
            btnButton.OnClick = "=SearchAngebot()"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFrmAngebotSuchen()"
    
    ' create subform
    Set frmSubForm = CreateControl(strTempFormName, acSubform, acDetail)
        With frmSubForm
            .Name = "frbSubForm"
            .Left = 510
            .Top = 2453
            .Width = 9218
            .Height = 5055
            .SourceObject = "frmAngebotSuchenSub"
            .Locked = True
        End With
        
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen: " & strFormName & " erstellt"
    End If

End Sub

Public Function CloseFrmAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function

Public Function SearchAngebot()

    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SearchAngebot ausfuehren"
    End If
    
    ' search term
    basBuild.BuildQryAngebotAuswahl Application.Forms.Item("frmAngebotSuchen").Controls("txtSearchBox")
    
    ' close form
    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
    
    ' open form
    DoCmd.OpenForm "frmAngebotSuchen", acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.SearchAngebot executed"
    End If
    
End Function

' set captions and values
Public Function CaptionAndValueSettings(ByVal intNumberOfRows As Integer) As String()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CaptionAndValue ausfuehren"
    End If
    
    Dim astrSettings() As String
    ReDim astrSettings(intNumberOfRows, 3)

    astrSettings(0, 0) = "label.Name"
        astrSettings(0, 1) = "label.Caption"
        astrSettings(0, 2) = "txtbox.Name"
        astrSettings(0, 3) = "txtbox"
    astrSettings(1, 0) = "lbl00"
        astrSettings(1, 1) = "Angebot"
        astrSettings(1, 2) = "txt00"
        astrSettings(1, 3) = "BWIKey"
    astrSettings(2, 0) = "lbl01"
        astrSettings(2, 1) = "Einzelauftrag"
        astrSettings(2, 2) = "txt01"
        astrSettings(2, 3) = "EAkurzKey"
    astrSettings(3, 0) = "lbl02"
        astrSettings(3, 1) = "Mengengerüst"
        astrSettings(3, 2) = "txt02"
        astrSettings(3, 3) = "MengengeruestLink"
    astrSettings(4, 0) = "lbl03"
        astrSettings(4, 1) = "Leistungsbeschreibung"
        astrSettings(4, 2) = "txt03"
        astrSettings(4, 3) = "LeistungsbeschreibungLink"
    astrSettings(5, 0) = "lbl04"
        astrSettings(5, 1) = "Bemerkung"
        astrSettings(5, 2) = "txt04"
        astrSettings(5, 3) = "Bemerkung"
    astrSettings(6, 0) = "lbl05"
        astrSettings(6, 1) = "wildcard"
        astrSettings(6, 2) = "txt05"
        astrSettings(6, 3) = "Bemerkung"
    
    CaptionAndValueSettings = astrSettings
End Function

' delete form
' 1. check if form exists
' 2. close form is loaded
' 3. delete form
Public Sub ClearForm(ByVal strFormName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.ClearForm"
    End If
    
    Dim objDummy As Object
    For Each objDummy In Application.CurrentProject.AllForms
        If objDummy.Name = strFormName Then
            
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
                
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basAngebotSuchen.ClearForm: " & strFormName & " ist geoeffnet, Formular schließen"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "basAngebotSuchen.ClearForm: " & strFormName & " existiert bereits, Formular loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.CalculateGrid"
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
        Debug.Print "basAngebotSuchen.CalculateGrid executed"
    End If
    
End Function

' get left from grid
Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchen.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuchen.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetLeft executed"
    End If
    
End Function

' get left from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchen.GetTop: column 0 is not available"
        MsgBox "basAngebotSuchen.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetTop executed"
    End If
    
End Function

' get left from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchen.GetWidth: column 0 is not available"
        MsgBox "basAngebotSuchen.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetWidth executed"
    End If
    
End Function

' get left from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchen.GetHeight: column 0 is not available"
        MsgBox "basAngebotSuchen.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.GetHeight executed"
    End If
    
End Function

Private Function OpenFormCreateOffer()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.OpenFormCreateOffer"
    End If

    Dim strFormName As String
    strFormName "frmCreateOffer"
    
    DoCmd.OpenForm strFormName, acNormal
    
    'event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.OpenFormCreateOffer executed"
    End If

End Function
