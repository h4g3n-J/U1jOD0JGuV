Attribute VB_Name = "basAngebotSuchenSub"
Option Compare Database
Option Explicit

' builds form
Public Sub BuildFormAngebotSuchenSub()
    
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.BuildFormAngebotSuchenSub ausfuehren"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchenSub"
    
    ' clear form
    basSupport.ClearForm strFormName
    
    ' declare object 'frm'
    Dim frm As Form
    
    ' create frm
    Set frm = CreateForm
    frm.RecordSource = "qryAngebotAuswahl"
    
    ' set OnCurrent methode
    ' frm.OnCurrent = "=SelectRecordset()"
    frm.OnCurrent = "=SelectRecordsetAngebot()"
    
    ' define cell width
    Dim intCellWidth As Integer
    intCellWidth = 4 * 567
    
    ' define cell heigth
    Dim intCellHeigth As Integer
    intCellHeigth = 1 * 567
    
    ' define padding left
    Dim intPaddingLeft As Integer
    intPaddingLeft = 1 * 567
    
    'define padding top
    Dim intPaddingTop As Integer
    intPaddingTop = 1 * 567
        
    ' define cell padding
    Dim intCellPaddingLeft As Integer
    intCellPaddingLeft = 567
    
    ' define cell padding
    Dim intCellPaddingTop As Integer
    intCellPaddingTop = 567
        
    ' get object settings
    Dim avarField As Variant
    avarField = basAngebotSuchenSub.ObjectSettings
    
    ' create control objects
    Dim intPositionHorizontal As Integer
    Dim intPositionVertical As Integer
    
    Dim inti As Integer
    
    ' avoid empty spaces caused by invisible fields
    Dim intj As Integer
    intj = 0
    
    ' declare textbox and label
    Dim ctlText As Control
    Dim ctlLabel As Control
    
    ' skip column name
    For inti = LBound(avarField, 1) + 1 To UBound(avarField, 1)
        
        ' compute horizontal and vertical position of the cells
        intPositionHorizontal = intPaddingLeft + (intCellWidth + intCellPaddingLeft)
        intPositionVertical = intPaddingTop + (intCellHeigth + intCellPaddingTop) * intj
        
        ' skip entry when visible = False
        If avarField(inti, 1) Then
            
            ' create textboxes
            ' expression.CreateControl (FormName, ControlType, Section, Parent, ColumnName, Left, Top, Width, Height)
            Set ctlText = CreateControl(frm.Name, acTextBox, acDetail, , , intPositionHorizontal, intPositionVertical, intCellWidth, intCellHeigth)
            
            ' set textbox name
            ctlText.Name = "txt" & intj
            
            ' link textbox to field
            ctlText.ControlSource = avarField(inti, 0)
            
            ' create labels
            Set ctlLabel = CreateControl(frm.Name, acLabel, acDetail, ctlText.Name, , intPaddingLeft, intPositionVertical, intCellWidth, intCellHeigth)
            ctlLabel.Name = "lbl" & intj
            
            ' set label caption
            ctlLabel.Caption = avarField(inti, 2)
            
            intj = intj + 1
        End If
        
    Next
    
    ' set form properties
        ' set defaultView to datasheet (2)
        frm.AllowDatasheetView = True
        frm.AllowFormView = False
        frm.DefaultView = 2
    
    ' restore form size
    DoCmd.Restore
    
    ' save temporary form name in strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = frm.Name
    
    ' close and save form
    DoCmd.Close acForm, strFormNameTemp, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
        
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basBuild.frmAngebotSuchenSub: " & strFormName & " erstellt"
    End If
End Sub

' form layout settings
' Private Function ObjectSettings(ByVal intItem As Integer, intProperty As Integer) As Variant
Private Function ObjectSettings() As Variant
    
    If gconVerbatim Then
        Debug.Print "basAngebotSuchenSub.ObjectSettings ausfuehren"
    End If
    
    Dim avarField(16, 3) As Variant
    avarField(0, 0) = "Feldname"
        avarField(0, 1) = "Visible"
        avarField(0, 2) = "Label"
        avarField(0, 3) = "Feldname"
    avarField(1, 0) = "BWIKey"
        avarField(1, 1) = True
        avarField(1, 2) = "BWI Alias"
        avarField(1, 3) = "txt0"
    avarField(2, 0) = "EAkurzKey"
        avarField(2, 1) = True
        avarField(2, 2) = "Einzelauftrag"
        avarField(2, 3) = "txt1"
    avarField(3, 0) = "MengengeruestLink"
        avarField(3, 1) = False
        avarField(3, 2) = "Mengengeruest"
        avarField(3, 3) = "txt2"
    avarField(4, 0) = "LeistungsbeschreibungLink"
        avarField(4, 1) = False
        avarField(4, 2) = "Leistungsbeschreibung"
        avarField(4, 3) = "txt3"
    avarField(5, 0) = "Verfuegung"
        avarField(5, 1) = False
        avarField(5, 2) = "Verfuegung"
        ' avarField(5, 3) = "txt0"
        avarField(5, 3) = Null
    avarField(6, 0) = "Bemerkung"
        avarField(6, 1) = True
        avarField(6, 2) = "Bemerkung"
        ' avarField(6, 3) = "txt0"
        avarField(6, 3) = Null
    avarField(7, 0) = "BeauftragtDatum"
        avarField(7, 1) = False
        avarField(7, 2) = "Beauftragt"
        ' avarField(7, 3) = "txt0"
        avarField(7, 3) = Null
    avarField(8, 0) = "AbgebrochenDatum"
        avarField(8, 1) = False
        avarField(8, 2) = "Abgebrochen"
        ' avarField(8, 3) = "txt0"
        avarField(8, 3) = Null
    avarField(9, 0) = "MitzeichnungI21Datum"
        avarField(9, 1) = False
        avarField(9, 2) = "Mitzeichnung I2.1"
        ' avarField(9, 3) = "txt0"
        avarField(9, 3) = Null
    avarField(10, 0) = "MitzeichnungI25Datum"
        avarField(10, 1) = False
        avarField(10, 2) = "Mitzeichnung I2.5"
        ' avarField(10, 3) = "txt0"
        avarField(10, 3) = Null
    avarField(11, 0) = "AngebotDatum"
        avarField(11, 1) = False
        ' avarField(11, 2) = "Angeboten"
        avarField(11, 3) = Null
    avarField(12, 0) = "AbgenommenDatum"
        avarField(12, 1) = False
        avarField(12, 2) = "Abgenommen"
        ' avarField(12, 3) = "tx0"
        avarField(12, 3) = Null
    avarField(13, 0) = "AftrBeginn"
        avarField(13, 1) = False
        avarField(13, 2) = "Auftragsbeginn"
        ' avarField(13, 3) = "txt0"
        avarField(13, 3) = Null
    avarField(14, 0) = "AftrEnde"
        avarField(14, 1) = False
        avarField(14, 2) = "Auftragsende"
        ' avarField(14, 3) = "txt0"
        avarField(14, 3) = Null
    avarField(15, 0) = "StorniertDatum"
        avarField(15, 1) = False
        avarField(15, 2) = "Storniert"
        ' avarField(15, 3) = "txt0"
        avarField(15, 3) = Null
    avarField(16, 0) = "AngebotBrutto"
        avarField(16, 1) = False
        avarField(16, 2) = "Betrag (Brutto)"
        ' avarField(16, 3) = "txt0"
        avarField(16, 3) = Null
        
    ' ObjectSettings = avarField(intItem, intProperty)
    ObjectSettings = avarField
End Function

' load recordset to destination form
Public Function SelectRecordsetAngebot()
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
    
    ' declare RecorsetName
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strDestFormName).Controls("frbSubFrm").Controls("BWIKey")
    
    ' initiate class Angebot
    Dim angebot As clsAngebot
    Set angebot = New clsAngebot
    
    ' set selected Recordset as Angebot
    angebot.SelectRecordset varRecordsetName
    
    Dim avarTextBoxAndLabelSettings As Variant
    avarTextBoxAndLabelSettings = basAngebotSuchenSub.ObjectSettings
    
    ' set textboxes and labels
    Dim inti As Integer
    
    ' assign values to textboxes
    ' skip titles
    For inti = LBound(avarTextBoxAndLabelSettings, 1) + 1 To UBound(avarTextBoxAndLabelSettings, 1)
        ' check if value is null
        ' opening frmSearchMain will open frmAuftragSuchenSub at a time
        ' when varTextboxesAndLabels is not set yet -> IsEmpty
        If Not IsEmpty(avarTextBoxAndLabelSettings(inti, 0)) And Not IsNull(avarTextBoxAndLabelSettings(inti, 0)) And Not IsNull(avarTextBoxAndLabelSettings(inti, 3)) Then
            Forms.Item(strDestFormName).Controls.Item(avarTextBoxAndLabelSettings(inti, 3)) = CallByName(angebot, avarTextBoxAndLabelSettings(inti, 0), VbGet)
        End If
    Next
    
ExitProc:
    Set angebot = Nothing
End Function
