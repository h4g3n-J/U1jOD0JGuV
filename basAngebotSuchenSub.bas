Attribute VB_Name = "basAngebotSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildFormAngebotSuchenSub()
    basAngebotSuchenSub.frmAngebotSuchenSub
End Sub


Private Sub frmAngebotSuchenSub()
    Dim strFormName As String
    strFormName = "frmAngebotSuchenSub"
    
    ' clear form
    basSupport.ClearForm strFormName
    
    ' declare object frm
    Dim frm As Form
    
    ' create new form
    Set frm = CreateForm
    frm.RecordSource = "qryAngebotAuswahl"
    
    Dim intCellWidth As Integer
    intCellWidth = 4 * 567
    
    Dim intCellHeigth As Integer
    intCellHeigth = 1 * 567
    
    Dim intPaddingLeft As Integer
    intPaddingLeft = 1 * 567
        
    Dim intPaddingTop As Integer
    intPaddingTop = 1 * 567
    
    Dim intCellPaddingLeft As Integer
    intCellPaddingLeft = 567
    
    Dim intCellPaddingTop As Integer
    intCellPaddingTop = 567
        
    Dim avarField(16, 2) As Variant
    avarField(0, 0) = "Feldname"
        avarField(0, 1) = "Visible"
        avarField(0, 2) = "Label"
    avarField(1, 0) = "BWIKey"
        avarField(1, 1) = True
        avarField(1, 2) = "BWI Alias"
    avarField(2, 0) = "EAkurzKey"
        avarField(2, 1) = True
        avarField(2, 2) = "Einzelauftrag"
    avarField(3, 0) = "MengengeruestLink"
        avarField(3, 1) = True
        avarField(3, 2) = "Mengengeruest"
    avarField(4, 0) = "LeistungsbeschreibungLink"
        avarField(4, 1) = True
        avarField(4, 2) = "Leistungsbeschreibung"
    avarField(5, 0) = "Verfuegung"
        avarField(5, 1) = False
        avarField(5, 2) = "Verfuegung"
    avarField(6, 0) = "Bemerkung"
        avarField(6, 1) = True
        avarField(6, 2) = "Bemerkung"
    avarField(7, 0) = "BeauftragtDatum"
        avarField(7, 1) = True
        avarField(7, 2) = "Beauftragt"
    avarField(8, 0) = "AbgebrochenDatum"
        avarField(8, 1) = True
        avarField(8, 2) = "Abgebrochen"
    avarField(9, 0) = "MitzeichnungI21Datum"
        avarField(9, 1) = True
        avarField(9, 2) = "Mitzeichnung I2.1"
    avarField(10, 0) = "MitzeichnungI25Datum"
        avarField(10, 1) = True
        avarField(10, 2) = "Mitzeichnung I2.5"
    avarField(11, 0) = "AngebotDatum"
        avarField(11, 1) = True
        avarField(11, 2) = "Angeboten"
    avarField(12, 0) = "AbgenommenDatum"
        avarField(12, 1) = True
        avarField(12, 2) = "Abgenommen"
    avarField(13, 0) = "AftrBeginn"
        avarField(13, 1) = True
        avarField(13, 2) = "Auftragsbeginn"
    avarField(14, 0) = "AftrEnde"
        avarField(14, 1) = True
        avarField(14, 2) = "Auftragsende"
    avarField(15, 0) = "StorniertDatum"
        avarField(15, 1) = True
        avarField(15, 2) = "Storniert"
    avarField(16, 0) = "AngebotBrutto"
        avarField(16, 1) = True
        avarField(16, 2) = "Betrag (Brutto)"
    
    ' create control objects
    Dim intPositionHorizontal As Integer
    Dim intPositionVertical As Integer
    Dim inti As Integer
    Dim intj As Integer
    intj = 0
    Dim ctlText As Control
    Dim ctlLabel As Control
    ' skip column name
    For inti = LBound(avarField, 1) + 1 To UBound(avarField, 1)
        
        intPositionHorizontal = intPaddingLeft + (intCellWidth + intCellPaddingLeft)
        intPositionVertical = intPaddingTop + (intCellHeigth + intCellPaddingTop) * intj
        
        ' if visibile = True
        If avarField(inti, 1) Then
            ' expression.CreateControl (FormName, ControlType, Section, Parent, ColumnName, Left, Top, Width, Height)
            
            ' create textboxes
            Set ctlText = CreateControl(frm.Name, acTextBox, acDetail, , , intPositionHorizontal, intPositionVertical, intCellWidth, intCellHeigth)
            ctlText.Name = "txt" & intj
            ctlText.ControlSource = avarField(inti, 0)
            ' create labels
            Set ctlLabel = CreateControl(frm.Name, acLabel, acDetail, ctlText.Name, , intPaddingLeft, intPositionVertical, intCellWidth, intCellHeigth)
            ctlLabel.Name = "lbl" & intj
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

