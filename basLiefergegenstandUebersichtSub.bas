Attribute VB_Name = "basLiefergegenstandUebersichtSub"
Option Compare Database
Option Explicit

Public Sub BuildLiefergegenstandUebersichtSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.BuildLiefergegenstandUebersichtSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLiefergegenstandUebersichtSub"
    
    ' clear form
    basLiefergegenstandUebersichtSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query
    Dim strQueryName As String
    strQueryName = "qryLiefergegenstandUebersicht"
    
    ' build query LiefergegenstandUebersichtSub
    basLiefergegenstandUebersichtSub.SearchLiefergegenstandUebersicht (strQueryName)
    
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
        
            intNumberOfColumns = 12
            intNumberOfRows = 2
            intColumnWidth = 1500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "AftrID"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "AftrID"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "AftrTitel"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "AftrTitel"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "BWIKey"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "BWIKey"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "LeistungsbeschreibungLink"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = True
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "LeistungsbeschreibungLink"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "MengengeruestLink"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "MengengeruestLink"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "EaAngebot"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl05
    intColumn = 6
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "EaAngebot"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 7
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .ControlSource = "BeauftragtDatum"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl06
    intColumn = 7
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "BeauftragtDatum"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 8
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .ControlSource = "AbgenommenDatum"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl07
    intColumn = 8
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "AbgenommenDatum"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt08
    intColumn = 9
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .ControlSource = "RechnungNr"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl08
    intColumn = 9
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "RechnungNr"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt09
    intColumn = 10
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .ControlSource = "RechnungBrutto"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl09
    intColumn = 10
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "RechnungBrutto"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt10
    intColumn = 11
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .ControlSource = "EaRechnung"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl10
    intColumn = 11
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "EaRechnung"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt11
    intColumn = 12
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt11"
            .ControlSource = "LeistungserfassungsblattID"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl11
    intColumn = 12
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt11")
        With lblLabel
            .Name = "lbl11"
            .Caption = "LeistungserfassungsblattID"
            .Left = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    ' ->
    objForm.OnCurrent = "=SelectLiefergegenstandUebersicht()"
        
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
        Debug.Print "basLiefergegenstandUebersichtSub.BuildLiefergegenstandUebersichtSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.ClearForm"
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
        Debug.Print "basLiefergegenstandUebersichtSub.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLiefergegenstandUebersichtSub"
    
    basLiefergegenstandUebersichtSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basLiefergegenstandUebersichtSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basLiefergegenstandUebersichtSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestClearForm executed"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.DeleteQuery"
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
                    Debug.Print "basLiefergegenstandUebersichtSub.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
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
        Debug.Print "basLiefergegenstandUebersichtSub.DeleteQuery executed"
    End If
    
End Sub

Private Sub TestDeleteQuery()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestDeleteQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryLiefergegenstandUebersicht"
    
    ' delete query
    basLiefergegenstandUebersichtSub.DeleteQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not deleted.", vbCritical, "basLiefergegenstandUebersichtSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " was not detected", vbOKOnly, "basLiefergegenstandUebersichtSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestDeleteQuery executed"
    End If
    
End Sub

Public Sub SearchLiefergegenstandUebersicht(ByVal strQueryName As String, Optional varSearchTerm As Variant = Null)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.SearchLiefergegenstandUebersicht"
    End If
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basLiefergegenstandUebersichtSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT tblAuftrag.AftrID, tblAuftrag.AftrTitel, tblAngebot.BWIKey, tblAngebot.MengengeruestLink, tblAngebot.LeistungsbeschreibungLink, tblLiefergegenstand.bezeichnung, tblLiefergegenstand.linkLieferschein, tblLiefergegenstand.seriennummer, tblLiefergegenstand.anzahl, tblLiefergegenstand.herstellerkennzeichen, tblLiefergegenstand.lieferdatum, tblLiefergegenstand.zielAftrID, tblLiefergegenstand.zielLinkLieferschein, tblLiefergegenstand.zielLieferdatum, tblLiefergegenstand.LiefergegenstandLink, tblLiefergegenstand.seriennummer2, tblLiefergegenstand.bemerkung" & _
                    " FROM (((tblAuftrag LEFT JOIN tblAuftragZuAngebot ON tblAuftrag.AftrID = tblAuftragZuAngebot.RefAftrID) LEFT JOIN tblAngebot ON tblAuftragZuAngebot.RefBWIkey = tblAngebot.BWIKey) LEFT JOIN tblAngebotZuLiefergegenstand ON tblAngebot.BWIKey = tblAngebotZuLiefergegenstand.RefBWIkey) LEFT JOIN tblLiefergegenstand ON tblAngebotZuLiefergegenstand.RefLiefergegenstandID = tblLiefergegenstand.LiefergegenstandID" & _
                    " WHERE (((tblAuftrag.AftrID) Like '" & varSearchTerm & "')) OR (((tblAuftrag.AftrTitel) Like '" & varSearchTerm & "')) OR (((tblAngebot.BWIKey) Like '" & varSearchTerm & "')) OR (((tblLiefergegenstand.bezeichnung) Like '" & varSearchTerm & "')) OR (((tblLiefergegenstand.seriennummer) Like '" & varSearchTerm & "')) OR (((tblLiefergegenstand.herstellerkennzeichen) Like '" & varSearchTerm & "')) OR (((tblLiefergegenstand.zielAftrID) Like '" & varSearchTerm & "')) OR (((tblLiefergegenstand.seriennummer2) Like '" & varSearchTerm & "'))" & _
                ";"
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
        Debug.Print "basLiefergegenstandUebersichtSub.SearchLiefergegenstandUebersicht executed"
    End If

End Sub

Private Sub TestSearchLiefergegenstandUebersicht()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestSearchLiefergegenstandUebersicht"
    End If
        
    ' build query qryRechnungSuchen
    Dim strQueryName As String
    strQueryName = "qryLiefergegenstandUebersicht"
    
    ' basBuildQueryLiefergegenstandUebersicht
    basLiefergegenstandUebersichtSub.SearchLiefergegenstandUebersicht strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basLiefergegenstandUebersichtSub.TestSearchLiefergegenstandUebersicht"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basLiefergegenstandUebersichtSub.TestSearchLiefergegenstandUebersicht"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestSearchLiefergegenstandUebersicht executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.CalculateGrid"
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
        Debug.Print "basLiefergegenstandUebersichtSub.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestCalculateGrid"
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
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basLiefergegenstandUebersichtSub.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basLiefergegenstandUebersichtSub.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basLiefergegenstandUebersichtSub.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basLiefergegenstandUebersichtSub.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basLiefergegenstandUebersichtSub.GetLeft: column 0 is not available"
        MsgBox "basLiefergegenstandUebersichtSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basLiefergegenstandUebersichtSub.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basLiefergegenstandUebersichtSub.TestGetLeft: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLiefergegenstandUebersichtSub.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basLiefergegenstandUebersichtSub.GetTop: column 0 is not available"
        MsgBox "basLiefergegenstandUebersichtSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basLiefergegenstandUebersichtSub.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basLiefergegenstandUebersichtSub.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLiefergegenstandUebersichtSub.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetHeight: column 0 is not available"
        MsgBox "basLiefergegenstandUebersichtSub.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basLiefergegenstandUebersichtSub.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basLiefergegenstandUebersichtSub.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLiefergegenstandUebersichtSub.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetWidth: column 0 is not available"
        MsgBox "basLiefergegenstandUebersichtSub.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLiefergegenstandUebersichtSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basLiefergegenstandUebersichtSub.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basLiefergegenstandUebersichtSub.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLiefergegenstandUebersichtSub.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.TestGetWidth executed"
    End If
    
End Sub

Public Function SelectLiefergegenstandUebersicht()
    ' Error Code 1: Form does not exist
    ' Error Code 2: Parent Form is not loaded

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandUebersichtSub.SelectLiefergegenstandUebersicht"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmLiefergegenstandUebersicht"
    
    ' check if frmLiefergegenstandUebersicht exists (Error Code: 1)
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If Not bolFormExists Then
        Debug.Print "basLiefergegenstandUebersichtSub.SelectLiefergegenstandUebersicht aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmLiefergegenstandUebersicht not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basLiefergegenstandUebersichtSub.SelectLiefergegenstandUebersicht aborted, Error Code: 2"
        Exit Function
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandUebersichtSub.SelectLiefergegenstandUebersicht executed"
    End If
    
End Function

