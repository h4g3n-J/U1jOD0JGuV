Attribute VB_Name = "basAngebotSuchenSub"
Option Compare Database
Option Explicit

' build form
Public Sub BuildAngebotSuchenSub()
    
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.BuildAngebotSuchenSub ausfuehren"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchenSub"
    
    ' clear existing form
    basAngebotSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim frmForm As Form
    
    ' create form
    Set frmForm = CreateForm
    
    ' set recordsetSource
    frmForm.RecordSource = "qryAngebotAuswahl"
    
    ' set OnCurrent methode
    frmForm.OnCurrent = "=SelectRecordsetAngebot()"
    
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
        basAngebotSuchenSub.CreateTextbox frmForm.Name, aintGridSettings, intNumberOfRows
    
        ' create labels
        basAngebotSuchenSub.CreateLabel frmForm.Name, aintGridSettings, intNumberOfRows
    
    ' set Caption and ControlSource
    basAngebotSuchenSub.CaptionAndSource frmForm.Name, intNumberOfRows
    
    Dim inti As Integer
    
    ' set form properties
        frmForm.AllowDatasheetView = True
        frmForm.AllowFormView = False
        frmForm.DefaultView = 2 ' 2 is for datasheet
    
    ' restore form size
    DoCmd.Restore
    
    ' save temporary form name in strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = frmForm.Name
    
    ' close and save form
    DoCmd.Close acForm, strFormNameTemp, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
        
End Sub

Private Sub CreateTextbox(ByVal strFormName As String, aintTableSettings() As Integer, ByVal intNumberOfRows As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.CreateTextbox ausfuehren"
    End If
    
    ' declare textbox
    Dim txtTextbox As Textbox
    
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
        Set txtTextbox = PositionObjectInTable(txtTextbox, aintTableSettings, intColumn, intRow) ' set position
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
Public Function SelectRecordsetAngebot()
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.SelectRecordsetAngebot ausfuehren"
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
    varRecordsetName = Forms.Item(strDestFormName).Controls("frbSubFrm").Controls("BWIKey")
    
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
