Attribute VB_Name = "basAngebotSuchenSub"
Option Compare Database
Option Explicit

' build form
Public Sub BuildAngebotSuchenSub()
    
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.BuildAngebotSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchenSub"
    
    ' clear form
    basAngebotSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create query qryAngebotAuswahl
    Dim strQueryName As String
    strQueryName = "qryAngebotAuswahl"
    basAngebotSuchenSub.BuildQryAngebotAuswahl strQueryName
    
    ' set recordsetSource
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
            intColumnWidth = 1600
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basAngebotSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "BWIKey"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "BWIKey"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "EAkurzKey"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "EAkurzKey"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "MengengeruestLink"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "MengengeruestLink"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "LeistungsbeschreibungLink"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "LeistungsbeschreibungLink"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "Bemerkung"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Bemerkung"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "BeauftragtDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl05
    intColumn = 6
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "BeauftragtDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 7
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .ControlSource = "AbgebrochenDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl06
    intColumn = 7
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "AbgebrochenDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 8
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .ControlSource = "AngebotDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl07
    intColumn = 8
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "AngebotDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt08
    intColumn = 9
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .ControlSource = "AbgenommenDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl08
    intColumn = 9
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "AbgenommenDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt09
    intColumn = 10
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .ControlSource = "AftrBeginn"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl09
    intColumn = 10
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "AftrBeginn"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt10
    intColumn = 11
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .ControlSource = "AftrEnde"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl10
    intColumn = 11
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "AftrEnde"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt11
    intColumn = 12
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .ControlSource = "StorniertDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl11
    intColumn = 12
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "StorniertDatum"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt12
    intColumn = 13
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt12"
            .ControlSource = "AngebotBrutto"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl12
    intColumn = 13
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt12")
        With lblLabel
            .Name = "lbl12"
            .Caption = "AngebotBrutto"
            .Left = basAngebotSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAngebotSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAngebotSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAngebotSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    ' column added? -> update intNumberOfColumns
        
    ' set OnCurrent methode
    objForm.OnCurrent = "=SelectAngebot()"
    
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
        Debug.Print "basAngebotSuchenSub.BuildAuftragSuchenSub executed"
    End If
        
End Sub

' load recordset to destination form
Public Function SelectAngebot()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.SelectAngebot"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    ' check if frmAngebotSuchen exists (Error Code: 1)
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If Not bolFormExists Then
        Debug.Print "basAngebotSuchenSub.selectAngebot aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmAngebotSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basAngebotSuchenSub.selectAngebot aborted, Error Code: 2"
        Exit Function
    End If
    
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
    
    ' select recordset
    Angebot.SelectRecordset varRecordsetName
    
    ' show recordset
    ' Forms.Item(strFormName).Controls.Item("insert textboxName here") = CallByName(Auftrag, "insert Attribute Name here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(Angebot, "BWIKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt02") = CallByName(Angebot, "MengengeruestLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(Angebot, "LeistungsbeschreibungLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt04") = CallByName(Angebot, "Bemerkung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt05") = CallByName(Angebot, "BeauftragtDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt06") = CallByName(Angebot, "AbgebrochenDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt07") = CallByName(Angebot, "AngebotDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt08") = CallByName(Angebot, "AbgenommenDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt09") = CallByName(Angebot, "AftrBeginn", VbGet)
    Forms.Item(strFormName).Controls.Item("txt10") = CallByName(Angebot, "AftrEnde", VbGet)
    Forms.Item(strFormName).Controls.Item("txt11") = CallByName(Angebot, "StorniertDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt12") = CallByName(Angebot, "AngebotBrutto", VbGet)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.SelectAngebot executed"
    End If
    
ExitProc:
    Set Angebot = Nothing
End Function

' returns array
' (column, row, property)
' properties: 0 - Left, 1 - Top, 2 - Width, 3 - Height
' calculates left, top, width and height parameters
Private Function CalculateInformationGrid(ByVal intNumberOfColumns As Integer, ByRef aintColumnWidth() As Integer, ByVal intNumberOfRows As Integer, Optional ByVal intLeft As Integer = 10000, Optional ByVal intTop As Integer = 2430)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.CalculateTableSetting ausfuehren"
    End If
    
    intNumberOfColumns = intNumberOfColumns - 1
    intNumberOfRows = intNumberOfRows - 1
    
    ' column dimension
    Const cintHorizontalSpacing As Integer = 60
            
    ' row dimension
    Dim intRowHeight As Integer
    intRowHeight = 330
    
    Const cintVerticalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    Dim aintGridSettings() As Integer
    ReDim aintGridSettings(intNumberOfColumns, intNumberOfRows, cintNumberOfProperties)
    
    ' compute cell position properties
    Dim inti As Integer
    Dim intj As Integer
    For inti = 0 To intNumberOfColumns
        ' For intr = 0 To cintNumberOfRows
        For intj = 0 To intNumberOfRows
            ' set column left
            aintGridSettings(inti, intj, 0) = intLeft + inti * (aintColumnWidth(inti) + cintHorizontalSpacing)
            ' set row top
            aintGridSettings(inti, intj, 1) = intTop + intj * (intRowHeight + cintVerticalSpacing)
            ' set column width
            aintGridSettings(inti, intj, 2) = aintColumnWidth(inti)
            ' set row height
            aintGridSettings(inti, intj, 3) = intRowHeight
        Next
    Next

    CalculateInformationGrid = aintGridSettings
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.CalculateInformationGrid ausgefuehrt"
    End If

End Function

Private Sub ClearForm(ByVal strFormName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.ClearForm"
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
        Debug.Print "basAngebotSuchenSub executed"
    End If
    
End Sub

Public Sub SearchAngebot(Optional varSearchTerm As Variant)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.SearchAngebot"
    End If
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' define query name
    Dim strQueryName As String
    strQueryName = "qryAngebotAuswahl"
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basAngebotSuchenSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT qryAngebot.*" & _
                    " FROM qryAngebot" & _
                    " WHERE qryAngebot.BWIKey LIKE '*" & strSearchTerm & "*'" & _
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
        Debug.Print "basAngebotSuchenSub.SearchAngebot executed"
    End If

End Sub

' build qryAngebotAuswahl
Private Sub BuildQryAngebotAuswahl(ByVal strQueryName As String)
        
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.BuildQryAngebotAuswahl"
    End If
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    ' basAngebotSuchenSub.DeleteQueryName (strQueryName)
    basAngebotSuchenSub.DeleteQuery (strQueryName)
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT qryAngebot.*" & _
            " FROM qryAngebot" & _
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
        Debug.Print "basAngebotSuchenSub.BuildQryAngebotAuswahl executed"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.DeleteQuery"
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
                    Debug.Print "basAngebotSuchenSub.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
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
        Debug.Print "basAngebotSuchenSub.DeleteQuery executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Long, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAngebotSuchenSub.CalculateGrid"
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
        Debug.Print "basAngebotSuchenSub.CalculateGrid executed"
    End If
    
End Function

Private Function TestCalculateGrid()

    Dim aintInformationGrid() As Integer
        
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intColumnWidth As Integer
        Dim intRowHeight As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intColumn As Integer
        Dim intRow As Integer
        
            intNumberOfColumns = 15
            intNumberOfRows = 2
            intColumnWidth = 1600
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basAngebotSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = False
    
    If Not bolOutput Then
        TestCalculateGrid = aintInformationGrid
        Exit Function
    End If
    
    For intColumn = 0 To UBound(aintInformationGrid, 1)
        For intRow = 0 To UBound(aintInformationGrid, 2)
            Debug.Print "column " & intColumn & ", row " & intRow & ", left: " & aintInformationGrid(intColumn, intRow, 0)
            Debug.Print "column " & intColumn & ", row " & intRow & ", top: " & aintInformationGrid(intColumn, intRow, 1)
            Debug.Print "column " & intColumn & ", row " & intRow & ", width: " & aintInformationGrid(intColumn, intRow, 2)
            Debug.Print "column " & intColumn & ", row " & intRow & ", height: " & aintInformationGrid(intColumn, intRow, 3)
        Next
    Next
    
    TestCalculateGrid = aintInformationGrid
    
End Function

' get left from grid
Private Function GetLeft(alngGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Long
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = alngGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()

    Dim aintGrid() As Integer
    aintGrid = basAngebotSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If Not bolOutput Then
        Exit Sub
    End If
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Left: " & basAngebotSuchenSub.GetLeft(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
End Sub

' get top from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchenSub.GetTop: column 0 is not available"
        MsgBox "basAngebotSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()

    Dim aintGrid() As Integer
    aintGrid = basAngebotSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Top: " & basAngebotSuchenSub.GetTop(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub

' get width from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchenSub.GetWidth: column 0 is not available"
        MsgBox "basAngebotSuchenSub.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.GetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()

    Dim aintGrid() As Integer
    aintGrid = basAngebotSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Width: " & basAngebotSuchenSub.GetWidth(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub

' get height from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchenSub.GetHeight: column 0 is not available"
        MsgBox "basAngebotSuchenSub.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.GetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()

    Dim aintGrid() As Integer
    aintGrid = basAngebotSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Height: " & basAngebotSuchenSub.GetHeight(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub

