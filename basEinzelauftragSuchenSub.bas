Attribute VB_Name = "basEinzelauftragSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildEinzelauftragSuchenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.BuildEinzelauftragSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchenSub"
    
    ' clear form
    basEinzelauftragSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query
    Dim strQueryName As String
    strQueryName = "qryEinzelauftragSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblEinzelauftrag"
    
    basEinzelauftragSuchenSub.SearchEinzelauftrag strQueryName, strQuerySource
    
    ' set recordset source
    objForm.RecordSource = strQueryName
    
    ' build information grid
    Dim aintInformationGrid() As Integer
        
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intColumnWidth As Integer
        Dim intRowHeight As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intColumn As Integer
        Dim intRow As Integer
        
            intNumberOfColumns = 16
            intNumberOfRows = 2
            intColumnWidth = 1500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "EAkurzKey"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "MengengeruestLink"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "MengengeruestLink"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "LeistungsbeschreibungLink"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "LeistungsbeschreibungLink"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "Bemerkung"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Bemerkung"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "BeauftragtDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "BeauftragtDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "AbgebrochenDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl05
    intColumn = 6
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "AbgebrochenDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 7
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .ControlSource = "MitzeichnungI21Datum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl06
    intColumn = 7
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "MitzeichnungI21Datum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 8
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .ControlSource = "MitzeichnungI25Datum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl07
    intColumn = 8
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "MitzeichnungI25Datum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt08
    intColumn = 9
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .ControlSource = "AngebotDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl08
    intColumn = 9
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "AngebotDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt09
    intColumn = 10
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .ControlSource = "AbgenommenDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl09
    intColumn = 10
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "AbgenommenDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt10
    intColumn = 11
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .ControlSource = "StorniertDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl10
    intColumn = 11
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "StorniertDatum"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt11
    intColumn = 12
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .ControlSource = "AngebotBrutto"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl11
    intColumn = 12
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "AngebotBrutto"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt12
    intColumn = 13
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt12"
            .ControlSource = "BWIKey"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl12
    intColumn = 13
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
        With lblLabel
            .Name = "lbl12"
            .Caption = "BWIKey"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt13
    intColumn = 14
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt13"
            .ControlSource = "AftrBeginn"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl13
    intColumn = 14
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt13")
        With lblLabel
            .Name = "lbl13"
            .Caption = "AftrBeginn"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt14
    intColumn = 15
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt14"
            .ControlSource = "AftrEnde"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl14
    intColumn = 15
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt14")
        With lblLabel
            .Name = "lbl14"
            .Caption = "AftrEnde"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt15
    intColumn = 16
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt15"
            .ControlSource = "EATitel"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl15
    intColumn = 16
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt15")
        With lblLabel
            .Name = "lbl15"
            .Caption = "EATitel"
            .Left = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    objForm.OnCurrent = "=SelectEinzelauftrag()"
        
    ' set form properties
    objForm.AllowDatasheetView = True
    objForm.AllowFormView = False
    objForm.DefaultView = 2 '2 is for datasheet
    
    ' restore form size
    DoCmd.Restore
    
    ' close and rename form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.BuildEinzelauftragSuchenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.ClearForm"
    End If
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        
        If objForm.Name = strFormName Then
        
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
            End If
            
            ' delete Form
            DoCmd.DeleteObject acForm, strFormName
            Exit For
            
        End If
        
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchenSub"
    
    basEinzelauftragSuchenSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basEinzelauftragSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basEinzelauftragSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestClearForm executed"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.DeleteQuery"
    End If
    
    ' set dummy object
    Dim objDummy As Object
    ' search object list >>AllQueries<< for strQueryName
    For Each objDummy In Application.CurrentData.AllQueries
        If objDummy.Name = strQueryName Then
            
            ' check if query isloaded
            If objDummy.IsLoaded Then
                ' close query
                DoCmd.Close acQuery, strQueryName, acSaveYes
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basEinzelauftragSuchenSub.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
                End If
            End If
    
            ' delete query
            DoCmd.DeleteObject acQuery, strQueryName
            
            ' exit loop
            Exit For
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.DeleteQuery executed"
    End If
    
End Sub

Private Sub TestDeleteQuery()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestDeleteQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryEinzelauftragSuchen"
    
    ' delete query
    basEinzelauftragSuchenSub.DeleteQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not deleted.", vbCritical, "basEinzelauftragSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " was not detected", vbOKOnly, "basEinzelauftragSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestDeleteQuery executed"
    End If
    
End Sub

Public Sub SearchEinzelauftrag(ByVal strQueryName As String, ByVal strQuerySource As String, Optional varSearchTerm As Variant = Null)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.SearchEinzelauftrag"
    End If
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
    
    Dim strSearchField01 As String
    strSearchField01 = "EAkurzKey"
    
    Dim strSearchField02 As String
    strSearchField02 = "EATitel"
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basEinzelauftragSuchenSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT " & strQuerySource & ".*" & _
                    " FROM " & strQuerySource & _
                    " WHERE (" & strQuerySource & "." & strSearchField01 & " Like '*" & strSearchTerm & "*') OR (" & strQuerySource & "." & strSearchField02 & " Like '*" & strSearchTerm & "*')" & _
                    " ;"
    End With
    
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With

ExitProc:
    qdfQuery.Close
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.SearchEinzelauftrag executed"
    End If

End Sub

Private Sub TestSearchEinzelauftrag()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestSearchEinzelauftrag"
    End If
        
    ' build query qryRechnungSuchen
    Dim strQueryName As String
    strQueryName = "qryEinzelauftragSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblEinzelauftrag"
    
    basEinzelauftragSuchenSub.SearchEinzelauftrag strQueryName, strQuerySource
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basEinzelauftragSuchenSub.TestSearchEinzelauftrag"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basEinzelauftragSuchenSub.TestSearchEinzelauftrag"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestSearchEinzelauftrag executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.CalculateGrid"
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
        Debug.Print "basEinzelauftragSuchenSub.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestCalculateGrid"
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
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basEinzelauftragSuchenSub.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basEinzelauftragSuchenSub.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basEinzelauftragSuchenSub.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basEinzelauftragSuchenSub.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basEinzelauftragSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basEinzelauftragSuchenSub.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basEinzelauftragSuchenSub.TestGetLeft: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchenSub.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchenSub.GetTop: column 0 is not available"
        MsgBox "basEinzelauftragSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basEinzelauftragSuchenSub.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basEinzelauftragSuchenSub.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchenSub.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetHeight: column 0 is not available"
        MsgBox "basEinzelauftragSuchenSub.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basEinzelauftragSuchenSub.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basEinzelauftragSuchenSub.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchenSub.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetWidth: column 0 is not available"
        MsgBox "basEinzelauftragSuchenSub.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basEinzelauftragSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basEinzelauftragSuchenSub.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basEinzelauftragSuchenSub.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basEinzelauftragSuchenSub.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.TestGetWidth executed"
    End If
    
End Sub

Public Function SelectEinzelauftrag()
    ' Error Code 1: Form does not exist
    ' Error Code 2: Parent Form is not loaded

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basEinzelauftragSuchenSub.SelectEinzelauftrag"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmEinzelauftragSuchen"
    
    ' check if frmEinzelauftragSuchen exists (Error Code: 1)
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If Not bolFormExists Then
        Debug.Print "basEinzelauftragSuchenSub.SelectEinzelauftrag aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmEinzelauftragSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basEinzelauftragSuchenSub.SelectEinzelauftrag aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare primary key
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "EAkurzKey"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Auftrag
    Dim Einzelauftrag As clsEinzelauftrag
    Set Einzelauftrag = New clsEinzelauftrag
    
    ' select recordset
    Einzelauftrag.SelectRecordset varRecordsetName
    
    ' show recordset
    ' referes to the textboxes in basEinzelauftragSuchen
    ' Forms.Item(strFormName).Controls.Item("insert_textboxName_here") = CallByName(insert_Object_Name, "insert_Attribute_Name_here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(Einzelauftrag, "EAkurzKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt01") = CallByName(Einzelauftrag, "MengengeruestLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt02") = CallByName(Einzelauftrag, "LeistungsbeschreibungLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(Einzelauftrag, "Bemerkung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt04") = CallByName(Einzelauftrag, "BeauftragtDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt05") = CallByName(Einzelauftrag, "AbgebrochenDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt06") = CallByName(Einzelauftrag, "MitzeichnungI21Datum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt07") = CallByName(Einzelauftrag, "MitzeichnungI25Datum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt08") = CallByName(Einzelauftrag, "AngebotDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt09") = CallByName(Einzelauftrag, "AbgenommenDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt10") = CallByName(Einzelauftrag, "StorniertDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt11") = CallByName(Einzelauftrag, "AngebotBrutto", VbGet)
    Forms.Item(strFormName).Controls.Item("txt12") = CallByName(Einzelauftrag, "BWIKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt13") = CallByName(Einzelauftrag, "AftrBeginn", VbGet)
    Forms.Item(strFormName).Controls.Item("txt14") = CallByName(Einzelauftrag, "AftrEnde", VbGet)
    Forms.Item(strFormName).Controls.Item("txt15") = CallByName(Einzelauftrag, "EATitel", VbGet)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basEinzelauftragSuchenSub.SelectEinzelauftrag executed"
    End If
    
End Function

