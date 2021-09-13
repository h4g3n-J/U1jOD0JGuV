Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

' build form AngebotSuchen
Public Sub BuildAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.BuildAngebotSuchen"
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
        intNumberOfRows = 15
        intLeft = 10000
        intTop = 2430
        intWidth = 3120
        intHeight = 330
        
        ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
        
        ' calculate information grid
        aintInformationGrid = basAngebotSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intWidth, intHeight)
        
        ' create textbox before label, so label can be associated
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
                .Caption = "Angebot Nr."
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
            
        ' lbl05
        intColumn = 1
        intRow = 6
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
            With lblLabel
                .Name = "lbl05"
                .Caption = "Beauftragt"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt06
        intColumn = 2
        intRow = 7
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt06"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl06
        intColumn = 1
        intRow = 7
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
            With lblLabel
                .Name = "lbl06"
                .Caption = "Abgebrochen"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt07
        intColumn = 2
        intRow = 8
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt07"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl07
        intColumn = 1
        intRow = 8
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
            With lblLabel
                .Name = "lbl07"
                .Caption = "Angeboten"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt08
        intColumn = 2
        intRow = 9
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt08"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl08
        intColumn = 1
        intRow = 9
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
            With lblLabel
                .Name = "lbl08"
                .Caption = "Abgenommen"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt09
        intColumn = 2
        intRow = 10
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt09"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl09
        intColumn = 1
        intRow = 10
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
            With lblLabel
                .Name = "lbl09"
                .Caption = "Auftrag Beginn"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt10
        intColumn = 2
        intRow = 11
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt10"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl10
        intColumn = 1
        intRow = 11
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
            With lblLabel
                .Name = "lbl10"
                .Caption = "Auftrag Ende"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt11
        intColumn = 2
        intRow = 12
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt11"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl11
        intColumn = 1
        intRow = 12
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
            With lblLabel
                .Name = "lbl11"
                .Caption = "Storniert"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
            End With
            
        ' txt12
        intColumn = 2
        intRow = 13
        Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
            With txtTextbox
                .Name = "txt12"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
                .IsHyperlink = False
            End With
            
        ' lbl12
        intColumn = 1
        intRow = 13
        Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
            With lblLabel
                .Name = "lbl12"
                .Caption = "Preis Brutto"
                .Left = basAngebotSuchen.GetLeft(aintInformationGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintInformationGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintInformationGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintInformationGrid, intColumn, intRow)
                .Visible = True
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
                .Name = "cmdCreateAngebot"
                .Left = basAngebotSuchen.GetLeft(aintLifecycleGrid, intColumn, intRow)
                .Top = basAngebotSuchen.GetTop(aintLifecycleGrid, intColumn, intRow)
                .Width = basAngebotSuchen.GetWidth(aintLifecycleGrid, intColumn, intRow)
                .Height = basAngebotSuchen.GetHeight(aintLifecycleGrid, intColumn, intRow)
                .Caption = "Angebot erstellen"
                .OnClick = "=OpenFormAngebotErstellen()"
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
            btnButton.OnClick = "=OpenSearchAngebot()"
            
        ' create exit button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdExit"
            btnButton.Left = 13180
            btnButton.Top = 960
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Schließen"
            btnButton.OnClick = "=CloseFrmAngebotSuchen()"
            
        ' create save button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdSave"
            btnButton.Left = 13180
            btnButton.Top = 1425
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Speichern"
            btnButton.OnClick = "=AngebotSuchenSaveRecordset()"
            
        ' create deleteRecordset button
        Set btnButton = CreateControl(strTempFormName, acCommandButton, acDetail)
        btnButton.Name = "cmdDeleteRecordset"
            btnButton.Left = 13180
            btnButton.Top = 1875
            btnButton.Width = 3120
            btnButton.Height = 330
            btnButton.Caption = "Datensatz löschen"
            btnButton.OnClick = "=AngebotSuchenDeleteRecordset()"
    
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
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen executed"
    End If

End Sub

Public Function CloseFrmAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.ClosefrmAngebotSuchen ausfuehren"
    End If

    DoCmd.Close acForm, "frmAngebotSuchen", acSaveYes
End Function

Public Function OpenSearchAngebot()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.SearchAngebot"
    End If
    
    ' search Angebot
    basAngebotSuchenSub.SearchAngebot Application.Forms.Item("frmAngebotSuchen").Controls("txtSearchBox")
    
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

Public Function OpenFormAngebotErstellen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.OpenFormAngebotErstellen"
    End If
    
    Dim strFormName As String
    strFormName = "frmAngebotErstellen"
    
    DoCmd.OpenForm strFormName, acNormal
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.OpenFormAngebotErstellen executed"
    End If
    
End Function

Public Function AngebotSuchenSaveRecordset()
    ' Error Code 1: no recordset was supplied
    ' Error Code 2: user aborted function

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.AngebotSuchenSaveRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' declare subform name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "BWIKey"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Angebot
    Dim Angebot As clsAngebot
    Set Angebot = New clsAngebot
    
    ' check primary key value
    If IsNull(varRecordsetName) Then
        Debug.Print "Error: basAngebotSuchen.AngebotSuchenSaveRecordset aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt. Speichern abgebrochen.", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' select recordset
    Angebot.SelectRecordset varRecordsetName
    
    ' allocate values to recordset properties
    With Angebot
        .EAkurzKey = Forms.Item(strFormName).Controls("txt01")
        .MengengeruestLink = Forms.Item(strFormName).Controls("txt02")
        .LeistungsbeschreibungLink = Forms.Item(strFormName).Controls("txt03")
        .Bemerkung = Forms.Item(strFormName).Controls("txt04")
        .BeauftragtDatum = Forms.Item(strFormName).Controls("txt05")
        .AbgebrochenDatum = Forms.Item(strFormName).Controls("txt06")
        .AngebotDatum = Forms.Item(strFormName).Controls("txt07")
        .AbgenommenDatum = Forms.Item(strFormName).Controls("txt08")
        .AftrBeginn = Forms.Item(strFormName).Controls("txt09")
        .AftrEnde = Forms.Item(strFormName).Controls("txt10")
        .StorniertDatum = Forms.Item(strFormName).Controls("txt11")
        .AngebotBrutto = Forms.Item(strFormName).Controls("txt12")
    End With
    
    ' delete recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Sollen die Änderungen am Datensatz " & varRecordsetName & " wirklich gespeichert werden?", vbOKCancel, "Speichern")
    
    If varUserInput = 1 Then
        Angebot.SaveRecordset
        MsgBox "Änderungen gespeichert", vbInformation, "Änderungen Speichern"
    Else
        Debug.Print "Error: basAngebotSuchen.AngebotSuchenSaveRecordset aborted, Error Code 2"
        MsgBox "Speichern abgebrochen", vbInformation, "Änderungen Speichern"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.AngebotSuchenSaveRecordset execute"
    End If
    
End Function

Public Function AngebotSuchenDeleteRecordset()
    ' Error Code 1: no recordset was supplied
    ' Error Code 2: user aborted function
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchen.AngebotSuchenDeleteRecordset"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "BWIKey"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Angebot
    Dim Angebot As clsAngebot
    Set Angebot = New clsAngebot
    
    ' check primary key value
    If IsNull(varRecordsetName) Then
        Debug.Print "Error: basAngebotSuchen.AngebotSuchenSaveRecordset aborted, Error Code 1"
        MsgBox "Es wurde kein Datensatz ausgewählt. Speichern abgebrochen.", vbCritical, "Fehler"
        Exit Function
    End If
    
    ' select recordset
    Angebot.SelectRecordset varRecordsetName
    
    ' delete recordset
    Dim varUserInput As Variant
    varUserInput = MsgBox("Soll der Datensatz " & varRecordsetName & " wirklich gelöscht werden?", vbOKCancel, "Datensatz löschen")
    
    If varUserInput = 1 Then
        Angebot.DeleteRecordset
        MsgBox "Datensatz gelöscht", vbInformation, "Datensatz löschen"
    Else
        Debug.Print "Error: basAngebotSuchen.AuftragSuchenDeleteRecordset aborted, Error Code 2"
        MsgBox "löschen abgebrochen", vbInformation, "Datensatz löschen"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.AngebotSuchenSaveRecordset execute"
    End If
    
End Function
