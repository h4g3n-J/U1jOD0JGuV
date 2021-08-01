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
        
            intNumberOfColumns = 11
            intNumberOfRows = 2
            intColumnWidth = 2500
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
            .ControlSource = "AftrID"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    
    ' set OnCurrent methode
    objForm.OnCurrent = "=SelectAngebot()"
    
    Dim aintColumnWidth(1) As Integer
    aintColumnWidth(0) = 2300
    aintColumnWidth(1) = 2300
    
    ' set number of rows
    Dim intNumberOfRows As Integer
    intNumberOfRows = 3
    
    ' set top left
    Dim intLeft As Integer
    intLeft = 100
    
    ' set top
    Dim intTop As Integer
    intTop = 100
    
    ' create table
        
        ' declare table settings
        Dim aintGridSettings() As Integer
        
        ' get grid settings
        aintGridSettings = basAngebotSuchenSub.CalculateInformationGrid(2, aintColumnWidth, intNumberOfRows, intLeft, intTop)
        
        ' create textboxes
        basAngebotSuchenSub.CreateTextbox objForm.Name, aintGridSettings, intNumberOfRows
    
        ' create labels
        basAngebotSuchenSub.CreateLabel objForm.Name, aintGridSettings, intNumberOfRows
    
    ' set Caption and ControlSource
    basAngebotSuchenSub.CaptionAndSource objForm.Name, intNumberOfRows
    
    Dim inti As Integer
    
    ' set form properties
        objForm.AllowDatasheetView = True
        objForm.AllowFormView = False
        objForm.DefaultView = 2 ' 2 is for datasheet
    
    ' restore form size
    DoCmd.Restore
    
    ' save temporary form name in strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = objForm.Name
    
    ' close and save form
    DoCmd.Close acForm, strFormNameTemp, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
        
End Function

Private Sub CreateTextbox(ByVal strFormName As String, aintTableSettings() As Integer, ByVal intNumberOfRows As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.CreateTextbox ausfuehren"
    End If
    
    ' declare textbox
    Dim txtTextbox As TextBox
    
    ' set default values
    Dim avarSettingsTable(15, 3) As Variant
    avarSettingsTable(0, 0) = "txt00" ' name
        avarSettingsTable(0, 1) = 1 ' column
        avarSettingsTable(0, 2) = 0 ' row
        avarSettingsTable(0, 3) = True 'visibility
    avarSettingsTable(1, 0) = "txt01"
        avarSettingsTable(1, 1) = 1
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
    avarSettingsTable(2, 0) = "txt02"
        avarSettingsTable(2, 1) = 1
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
    avarSettingsTable(3, 0) = "txt03"
        avarSettingsTable(3, 1) = 1
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = False
    avarSettingsTable(4, 0) = "txt04"
        avarSettingsTable(4, 1) = 1
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = False
    avarSettingsTable(5, 0) = "txt05"
        avarSettingsTable(5, 1) = 1
        avarSettingsTable(5, 2) = 5
        avarSettingsTable(5, 3) = False
    avarSettingsTable(6, 0) = "txt06"
        avarSettingsTable(6, 1) = 1
        avarSettingsTable(6, 2) = 6
        avarSettingsTable(6, 3) = False
    avarSettingsTable(7, 0) = "txt07"
        avarSettingsTable(7, 1) = 1
        avarSettingsTable(7, 2) = 7
        avarSettingsTable(7, 3) = False
    avarSettingsTable(8, 0) = "txt08"
        avarSettingsTable(8, 1) = 1
        avarSettingsTable(8, 2) = 8
        avarSettingsTable(8, 3) = False
    avarSettingsTable(9, 0) = "txt09"
        avarSettingsTable(9, 1) = 1
        avarSettingsTable(9, 2) = 9
        avarSettingsTable(9, 3) = False
    avarSettingsTable(10, 0) = "txt10"
        avarSettingsTable(10, 1) = 1
        avarSettingsTable(10, 2) = 10
        avarSettingsTable(10, 3) = False
    avarSettingsTable(11, 0) = "txt11"
        avarSettingsTable(11, 1) = 1
        avarSettingsTable(11, 2) = 11
        avarSettingsTable(11, 3) = False
    avarSettingsTable(12, 0) = "txt12"
        avarSettingsTable(12, 1) = 1
        avarSettingsTable(12, 2) = 12
        avarSettingsTable(12, 3) = False
    avarSettingsTable(13, 0) = "txt13"
        avarSettingsTable(13, 1) = 1
        avarSettingsTable(13, 2) = 12
        avarSettingsTable(13, 3) = False
    avarSettingsTable(14, 0) = "txt14"
        avarSettingsTable(14, 1) = 1
        avarSettingsTable(14, 2) = 13
        avarSettingsTable(14, 3) = False
    avarSettingsTable(15, 0) = "txt15"
        avarSettingsTable(15, 1) = 1
        avarSettingsTable(15, 2) = 14
        avarSettingsTable(15, 3) = False
        
    intNumberOfRows = intNumberOfRows - 1
    
    Dim intColumn As Integer
    Dim intRow As Integer
    
    Dim inti As Integer
    For inti = LBound(avarSettingsTable, 1) To intNumberOfRows
        Set txtTextbox = CreateControl(strFormName, acTextBox, acDetail)
        txtTextbox.Name = avarSettingsTable(inti, 0) ' set name
        txtTextbox.Visible = avarSettingsTable(inti, 3) ' set visibility
        
        intColumn = avarSettingsTable(inti, 1)
        intRow = avarSettingsTable(inti, 2)
        Set txtTextbox = basSupport.PositionObjectInTable(txtTextbox, aintTableSettings, intColumn, intRow) ' set position
    Next
    
End Sub

Private Sub CreateLabel(ByVal strFormName As String, ByRef aintTableSettings() As Integer, ByVal intNumberOfRows As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.CreateLabel ausfuehren"
    End If
    
    ' declare label
    Dim lblLabel As Label

    Dim avarSettingsTable(15, 4) As Variant
    avarSettingsTable(0, 0) = "lbl00"
        avarSettingsTable(0, 1) = 0
        avarSettingsTable(0, 2) = 0
        avarSettingsTable(0, 3) = True
        avarSettingsTable(0, 4) = "txt00"
    avarSettingsTable(1, 0) = "lbl01"
        avarSettingsTable(1, 1) = 0
        avarSettingsTable(1, 2) = 1
        avarSettingsTable(1, 3) = True
        avarSettingsTable(1, 4) = "txt01"
    avarSettingsTable(2, 0) = "lbl02"
        avarSettingsTable(2, 1) = 0
        avarSettingsTable(2, 2) = 2
        avarSettingsTable(2, 3) = True
        avarSettingsTable(2, 4) = "txt02"
    avarSettingsTable(3, 0) = "lbl03"
        avarSettingsTable(3, 1) = 0
        avarSettingsTable(3, 2) = 3
        avarSettingsTable(3, 3) = False
        avarSettingsTable(3, 4) = "txt03"
    avarSettingsTable(4, 0) = "lbl04"
        avarSettingsTable(4, 1) = 0
        avarSettingsTable(4, 2) = 4
        avarSettingsTable(4, 3) = False
        avarSettingsTable(4, 4) = "txt04"
    avarSettingsTable(5, 0) = "lbl05"
        avarSettingsTable(5, 1) = 0
        avarSettingsTable(5, 2) = 5
        avarSettingsTable(5, 3) = False
        avarSettingsTable(5, 4) = "txt05"
    avarSettingsTable(6, 0) = "lbl06"
        avarSettingsTable(6, 1) = 0
        avarSettingsTable(6, 2) = 6
        avarSettingsTable(6, 3) = False
        avarSettingsTable(6, 4) = "txt06"
    avarSettingsTable(7, 0) = "lbl07"
        avarSettingsTable(7, 1) = 0
        avarSettingsTable(7, 2) = 7
        avarSettingsTable(7, 3) = False
        avarSettingsTable(7, 4) = "txt07"
    avarSettingsTable(8, 0) = "lbl08"
        avarSettingsTable(8, 1) = 0
        avarSettingsTable(8, 2) = 8
        avarSettingsTable(8, 3) = False
        avarSettingsTable(8, 4) = "txt08"
    avarSettingsTable(9, 0) = "lbl09"
        avarSettingsTable(9, 1) = 0
        avarSettingsTable(9, 2) = 9
        avarSettingsTable(9, 3) = False
        avarSettingsTable(9, 4) = "txt09"
    avarSettingsTable(10, 0) = "lbl10"
        avarSettingsTable(10, 1) = 0
        avarSettingsTable(10, 2) = 10
        avarSettingsTable(10, 3) = False
        avarSettingsTable(10, 4) = "txt10"
    avarSettingsTable(11, 0) = "lbl11"
        avarSettingsTable(11, 1) = 0
        avarSettingsTable(11, 2) = 11
        avarSettingsTable(11, 3) = False
        avarSettingsTable(11, 4) = "txt11"
    avarSettingsTable(12, 0) = "lbl12"
        avarSettingsTable(12, 1) = 0
        avarSettingsTable(12, 2) = 12
        avarSettingsTable(12, 3) = False
        avarSettingsTable(12, 4) = "txt12"
    avarSettingsTable(13, 0) = "lbl13"
        avarSettingsTable(13, 1) = 0
        avarSettingsTable(13, 2) = 12
        avarSettingsTable(13, 3) = False
        avarSettingsTable(13, 4) = "txt13"
    avarSettingsTable(14, 0) = "lbl14"
        avarSettingsTable(14, 1) = 0
        avarSettingsTable(14, 2) = 13
        avarSettingsTable(14, 3) = False
        avarSettingsTable(14, 4) = "txt14"
    avarSettingsTable(15, 0) = "lbl15"
        avarSettingsTable(15, 1) = 0
        avarSettingsTable(15, 2) = 14
        avarSettingsTable(15, 3) = False
        avarSettingsTable(15, 4) = "txt15"
    
    Dim intColumn As Integer
    Dim intRow As Integer
        
    Dim inti As Integer
    For inti = LBound(avarSettingsTable, 1) To 2
        Set lblLabel = CreateControl(strFormName, acLabel, acDetail, avarSettingsTable(inti, 4))
        lblLabel.Name = avarSettingsTable(inti, 0) ' set name
        lblLabel.Visible = avarSettingsTable(inti, 3) ' set visibility
        
        intColumn = avarSettingsTable(inti, 1)
        intRow = avarSettingsTable(inti, 2)
        Set lblLabel = basSupport.PositionObjectInTable(lblLabel, aintTableSettings, intColumn, intRow) ' set position
    Next
    
End Sub

' contains field settings
Private Sub CaptionAndSource(ByVal strFormName As String, ByVal intNumberOfRows As Integer)

    Dim astrCaptionAndSettings(3, 3) As String
    astrCaptionAndSettings(0, 0) = "Label.Name"
        astrCaptionAndSettings(0, 1) = "Label.Caption"
        astrCaptionAndSettings(0, 2) = "Textbox.Name"
        astrCaptionAndSettings(0, 3) = "Textbox.ControlSource"
    astrCaptionAndSettings(1, 0) = "lbl00"
        astrCaptionAndSettings(1, 1) = "BWI Alias"
        astrCaptionAndSettings(1, 2) = "txt00"
        astrCaptionAndSettings(1, 3) = "BWIKey"
    astrCaptionAndSettings(2, 0) = "lbl01"
        astrCaptionAndSettings(2, 1) = "Einzelauftrag"
        astrCaptionAndSettings(2, 2) = "txt01"
        astrCaptionAndSettings(2, 3) = "EAkurzKey"
    astrCaptionAndSettings(3, 0) = "lbl02"
        astrCaptionAndSettings(3, 1) = "Bemerkung"
        astrCaptionAndSettings(3, 2) = "txt02"
        astrCaptionAndSettings(3, 3) = "Bemerkung"
    
    ' avarField(3, 0) = "MengengeruestLink"
        ' avarField(3, 1) = False
        ' avarField(3, 2) = "Mengengeruest"
        ' avarField(3, 3) = "txt2"
    ' avarField(4, 0) = "LeistungsbeschreibungLink"
        ' avarField(4, 1) = False
        ' avarField(4, 2) = "Leistungsbeschreibung"
        ' avarField(4, 3) = "txt3"
    ' avarField(5, 0) = "Verfuegung"
        ' avarField(5, 1) = False
        ' avarField(5, 2) = "Verfuegung"
        ' avarField(5, 3) = "txt0"
        ' avarField(5, 3) = Null
    ' avarField(7, 0) = "BeauftragtDatum"
        ' avarField(7, 1) = False
        ' avarField(7, 2) = "Beauftragt"
        ' avarField(7, 3) = "txt0"
        ' avarField(7, 3) = Null
    ' avarField(8, 0) = "AbgebrochenDatum"
        ' avarField(8, 1) = False
        ' avarField(8, 2) = "Abgebrochen"
        ' avarField(8, 3) = "txt0"
        ' avarField(8, 3) = Null
    ' avarField(9, 0) = "MitzeichnungI21Datum"
        ' avarField(9, 1) = False
        ' avarField(9, 2) = "Mitzeichnung I2.1"
        ' avarField(9, 3) = "txt0"
        ' avarField(9, 3) = Null
    ' avarField(10, 0) = "MitzeichnungI25Datum"
        ' avarField(10, 1) = False
        ' avarField(10, 2) = "Mitzeichnung I2.5"
        ' avarField(10, 3) = "txt0"
        ' avarField(10, 3) = Null
    ' avarField(11, 0) = "AngebotDatum"
        ' avarField(11, 1) = False
        ' avarField(11, 2) = "Angeboten"
        ' avarField(11, 3) = Null
    ' avarField(12, 0) = "AbgenommenDatum"
        ' avarField(12, 1) = False
        ' avarField(12, 2) = "Abgenommen"
        ' avarField(12, 3) = "tx0"
        ' avarField(12, 3) = Null
    ' avarField(13, 0) = "AftrBeginn"
        ' avarField(13, 1) = False
        ' avarField(13, 2) = "Auftragsbeginn"
        ' avarField(13, 3) = "txt0"
        ' avarField(13, 3) = Null
    ' avarField(14, 0) = "AftrEnde"
        ' avarField(14, 1) = False
        ' avarField(14, 2) = "Auftragsende"
        ' avarField(14, 3) = "txt0"
        ' avarField(14, 3) = Null
    ' avarField(15, 0) = "StorniertDatum"
        ' avarField(15, 1) = False
        ' avarField(15, 2) = "Storniert"
        ' avarField(15, 3) = "txt0"
        ' avarField(15, 3) = Null
    ' avarField(16, 0) = "AngebotBrutto"
        ' avarField(16, 1) = False
        ' avarField(16, 2) = "Betrag (Brutto)"
        ' avarField(16, 3) = "txt0"
        ' avarField(16, 3) = Null
    
    ' set caption and controlSource
    Dim inti As Integer
    For inti = LBound(astrCaptionAndSettings, 1) + 1 To intNumberOfRows
        Forms(strFormName).Controls(astrCaptionAndSettings(inti, 0)).Caption = astrCaptionAndSettings(inti, 1) ' set caption
        Forms(strFormName).Controls(astrCaptionAndSettings(inti, 2)).ControlSource = astrCaptionAndSettings(inti, 3) ' set ControlSource
    Next
    
End Sub

' load recordset to destination form
Public Function SelectAngebot()
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.SelectAngebot ausfuehren"
    End If
    
    ' destination form name setting
    Dim strDestFormName As String
    strDestFormName = "frmAngebotSuchen"
    
    ' check if destination form is loaded
    If Not (CurrentProject.AllForms(strDestFormName).IsLoaded) Then
        Debug.Print "basAngebotSuchenSub.SelectRecordset: " & strDestFormName _
            & " nicht geladen, Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    ' declare reference attribute
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strDestFormName).Controls("frbSubForm").Controls("BWIKey")
    
    ' initiate class Angebot
    Dim angebot As clsAngebot
    Set angebot = New clsAngebot
    
    ' set selected Recordset
    angebot.SelectRecordset varRecordsetName
    
    Dim intNumberOfRows As Integer
    intNumberOfRows = 6
    
    Dim astrTextBoxValues() As String
    astrTextBoxValues = basAngebotSuchen.CaptionAndValueSettings(intNumberOfRows)
    
    ' set textboxes and labels
    Dim inti As Integer
    
    ' assign values to textboxes
    For inti = LBound(astrTextBoxValues, 1) + 1 To intNumberOfRows ' skip titles
        Forms.Item(strDestFormName).Controls.Item(astrTextBoxValues(inti, 2)) = CallByName(angebot, astrTextBoxValues(inti, 3), VbGet)
    Next
    
ExitProc:
    Set angebot = Nothing
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

' delete form
' 1. check if form exists
' 2. close if form is loaded
' 3. delete form
Private Sub ClearForm(ByVal strFormName As String)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.ClearForm ausfuehren"
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
                    Debug.Print "basAngebotSuchenSub.ClearForm: " & strFormName & " ist geoeffnet, form schlie?en"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' event message
            If gconVerbatim = True Then
                Debug.Print "basAngebotSuchenSub.ClearForm: " & strFormName & " existiert bereits, form loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
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
                    Debug.Print "basSupport.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
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
        Debug.Print "basBuild.DeleteQuery executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

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
        
            intNumberOfColumns = 11
            intNumberOfRows = 2
            intColumnWidth = 2500
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
Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAngebotSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basAngebotSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
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
