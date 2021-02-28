Attribute VB_Name = "basSearchMain"
' basSearchMain

Option Compare Database
Option Explicit

' adjust Auftrag Focus
Private Function AuftragFocus() As String
    AuftragFocus = "txt0"
End Function

' adjust Auftrag textboxes and labels
Private Function AuftragTextboxesAndLabels() As Variant
    Dim varSetup(9, 11) As Variant
    varSetup(0, 0) = "label name"
        varSetup(1, 0) = "label caption"
        varSetup(2, 0) = "textbox name"
        varSetup(3, 0) = "textbox value"
        varSetup(4, 0) = "textbox border style"
        varSetup(5, 0) = "textbox ishyperlink"
        varSetup(6, 0) = "textbox locked"
        varSetup(7, 0) = "textbox format"
        varSetup(8, 0) = "textbox visible"
        varSetup(9, 0) = "textbox defaultValue"
    varSetup(0, 1) = "lbl0"
        varSetup(1, 1) = "ID"
        varSetup(2, 1) = "txt0"
        varSetup(3, 1) = "AftrID"
        varSetup(4, 1) = 0
        varSetup(5, 1) = False
        varSetup(6, 1) = True
        varSetup(7, 1) = ""
        varSetup(8, 1) = True
        varSetup(9, 1) = Null
    varSetup(0, 2) = "lbl1"
        varSetup(1, 2) = "Titel"
        varSetup(2, 2) = "txt1"
        varSetup(3, 2) = "AftrTitel"
        varSetup(4, 2) = 1
        varSetup(5, 2) = False
        varSetup(6, 2) = False
        varSetup(7, 2) = ""
        varSetup(8, 2) = True
        varSetup(9, 2) = Null
    varSetup(0, 3) = "lbl2"
        varSetup(1, 3) = "ICD Status"
        varSetup(2, 3) = "txt2"
        varSetup(3, 3) = "StatusKey"
        varSetup(4, 3) = 0
        varSetup(5, 3) = False
        varSetup(6, 3) = True
        varSetup(7, 3) = ""
        varSetup(8, 3) = True
        varSetup(9, 3) = Null
    varSetup(0, 4) = "lbl3"
        varSetup(1, 4) = "Owner"
        varSetup(2, 4) = "txt3"
        varSetup(3, 4) = "OwnerKey"
        varSetup(4, 4) = 0
        varSetup(5, 4) = False
        varSetup(6, 4) = True
        varSetup(7, 4) = ""
        varSetup(8, 4) = True
        varSetup(9, 4) = Null
    varSetup(0, 5) = "lbl4"
        varSetup(1, 5) = "Priorität"
        varSetup(2, 5) = "txt4"
        varSetup(3, 5) = "PrioritaetKey"
        varSetup(4, 5) = 0
        varSetup(5, 5) = False
        varSetup(6, 5) = True
        varSetup(7, 5) = ""
        varSetup(8, 5) = True
        varSetup(9, 5) = Null
    varSetup(0, 6) = "lbl5"
        varSetup(1, 6) = "Parent"
        varSetup(2, 6) = "txt5"
        varSetup(3, 6) = "ParentKey"
        varSetup(4, 6) = 0
        varSetup(5, 6) = False
        varSetup(6, 6) = True
        varSetup(7, 6) = ""
        varSetup(8, 6) = True
        varSetup(9, 6) = Null
    varSetup(0, 7) = "lbl6"
        varSetup(1, 7) = "Bemerkung"
        varSetup(2, 7) = "txt6"
        varSetup(3, 7) = "Bemerkung"
        varSetup(4, 7) = 1
        varSetup(5, 7) = False
        varSetup(6, 7) = False
        varSetup(7, 7) = ""
        varSetup(8, 7) = True
        varSetup(9, 7) = Null
    varSetup(0, 8) = "lbl7"
        varSetup(1, 8) = "Beginn (Soll)"
        varSetup(2, 8) = "txt7"
        varSetup(3, 8) = "BeginnSoll"
        varSetup(4, 8) = 0
        varSetup(5, 8) = False
        varSetup(6, 8) = True
        varSetup(7, 8) = "Short Date"
        varSetup(8, 8) = True
        varSetup(9, 8) = Null
    varSetup(0, 9) = "lbl8"
        varSetup(1, 9) = "Ende (Soll)"
        varSetup(2, 9) = "txt8"
        varSetup(3, 9) = "EndeSoll"
        varSetup(4, 9) = 0
        varSetup(5, 9) = False
        varSetup(6, 9) = True
        varSetup(7, 9) = "Short Date"
        varSetup(8, 9) = True
        varSetup(9, 9) = Null
    varSetup(0, 10) = "lbl9"
        varSetup(1, 10) = "Kunde"
        varSetup(2, 10) = "txt9"
        varSetup(3, 10) = "Kunde"
        varSetup(4, 10) = 1
        varSetup(5, 10) = False
        varSetup(6, 10) = False
        varSetup(7, 10) = ""
        varSetup(8, 10) = True
        varSetup(9, 10) = Null
    varSetup(0, 11) = Null
        varSetup(1, 11) = Null
        varSetup(2, 11) = "txt10"
        varSetup(3, 11) = Null
        varSetup(4, 11) = 1
        varSetup(5, 11) = False
        varSetup(6, 11) = False
        varSetup(7, 11) = ""
        varSetup(8, 11) = True
        varSetup(9, 11) = ""
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.AuftragTextboxesAndLabels ausgefuehrt"
    End If
    
    AuftragTextboxesAndLabels = varSetup
        
End Function

' adjust Auftrag commandbuttons
Private Function AuftragCommandButtons() As Variant
    Dim varSetup(2, 12) As Variant
    varSetup(0, 0) = "object Name"
        varSetup(1, 0) = "object caption"
        varSetup(2, 0) = "object visible"
    varSetup(0, 1) = "cmd0"
        varSetup(1, 1) = "Hauptmenü"
        varSetup(2, 1) = True
    varSetup(0, 2) = "cmd1"
        varSetup(1, 2) = "Suchen"
        varSetup(2, 2) = True
    varSetup(0, 3) = "cmd2"
        varSetup(1, 3) = "Auftrag erstellen"
        varSetup(2, 3) = True
    varSetup(0, 4) = "cmd3"
        varSetup(1, 4) = "Auftrag erteilen"
        varSetup(2, 4) = False
    varSetup(0, 5) = "cmd4"
        varSetup(1, 5) = "Leistung abnehmen"
        varSetup(2, 5) = False
    varSetup(0, 6) = "cmd5"
        varSetup(1, 6) = "Rechnung erfassen"
        varSetup(2, 6) = False
    varSetup(0, 7) = "cmd6"
        varSetup(1, 7) = "Auftrag stornieren"
        varSetup(2, 7) = False
    varSetup(0, 8) = "cmd7"
        varSetup(1, 8) = "Auftragdetails anzeigen"
        varSetup(2, 8) = False
    varSetup(0, 9) = "cmd8"
        varSetup(1, 9) = "Liefergegenstand anzeigen"
        varSetup(2, 9) = False
    varSetup(0, 10) = "cmd9"
        varSetup(1, 10) = "Speichern"
        varSetup(2, 10) = True
    varSetup(0, 11) = "cmd10"
        varSetup(1, 11) = "leer"
        varSetup(2, 11) = False
    varSetup(0, 12) = "cmd11"
        varSetup(1, 12) = "Speichern"
        varSetup(2, 12) = False
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.AuftragCommandButtons ausgefuehrt"
    End If
        
    AuftragCommandButtons = varSetup
End Function

' adjust Auftrag commandbuttons
Private Function AuftragSubForm() As Variant
    Dim varSetup(1, 1) As Variant
    varSetup(0, 0) = "object Name"
        varSetup(1, 0) = "source"
    varSetup(0, 1) = "frb1"
        varSetup(1, 1) = "frmAuftragSuchenSub"
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.AuftragSubForm ausgefuehrt"
    End If
        
    AuftragSubForm = varSetup
End Function

Private Sub LiefergegenstandSuchenConfig()
    ' set focus
    mstrFocus = "txt0"
    
    ' 0 = label name
    ' 1 = label caption
    ' 2 = textbox name
    ' 3 = textbox value
    ' 4 = textbox border style
    ' 5 = textbox ishyperlink
    ' 6 = textbox locked
    ' 7 = textbox format
    ' 8 = textbox visible
    ' 9 = textbox defaultValue
    avarTextBoxAndLabelConfig(0, 0) = "lbl0"
        avarTextBoxAndLabelConfig(1, 0) = "ID"
        avarTextBoxAndLabelConfig(2, 0) = "txt0"
        avarTextBoxAndLabelConfig(3, 0) = "AftrID"
        avarTextBoxAndLabelConfig(4, 0) = 0
        avarTextBoxAndLabelConfig(5, 0) = False
        avarTextBoxAndLabelConfig(6, 0) = True
        avarTextBoxAndLabelConfig(7, 0) = ""
        avarTextBoxAndLabelConfig(8, 0) = True
        avarTextBoxAndLabelConfig(9, 0) = Null
    avarTextBoxAndLabelConfig(0, 1) = "lbl1"
        avarTextBoxAndLabelConfig(1, 1) = "Titel"
        avarTextBoxAndLabelConfig(2, 1) = "txt1"
        avarTextBoxAndLabelConfig(3, 1) = "AftrTitel"
        avarTextBoxAndLabelConfig(4, 1) = 1
        avarTextBoxAndLabelConfig(5, 1) = False
        avarTextBoxAndLabelConfig(6, 1) = False
        avarTextBoxAndLabelConfig(7, 1) = ""
        avarTextBoxAndLabelConfig(8, 1) = True
        avarTextBoxAndLabelConfig(9, 1) = Null
    avarTextBoxAndLabelConfig(0, 2) = "lbl2"
        avarTextBoxAndLabelConfig(1, 2) = "ICD Status"
        avarTextBoxAndLabelConfig(2, 2) = "txt2"
        avarTextBoxAndLabelConfig(3, 2) = "StatusKey"
        avarTextBoxAndLabelConfig(4, 2) = 0
        avarTextBoxAndLabelConfig(5, 2) = False
        avarTextBoxAndLabelConfig(6, 2) = True
        avarTextBoxAndLabelConfig(7, 2) = ""
        avarTextBoxAndLabelConfig(8, 2) = True
        avarTextBoxAndLabelConfig(9, 2) = Null
    avarTextBoxAndLabelConfig(0, 3) = "lbl3"
        avarTextBoxAndLabelConfig(1, 3) = "Owner"
        avarTextBoxAndLabelConfig(2, 3) = "txt3"
        avarTextBoxAndLabelConfig(3, 3) = "OwnerKey"
        avarTextBoxAndLabelConfig(4, 3) = 0
        avarTextBoxAndLabelConfig(5, 3) = False
        avarTextBoxAndLabelConfig(6, 3) = True
        avarTextBoxAndLabelConfig(7, 3) = ""
        avarTextBoxAndLabelConfig(8, 3) = True
        avarTextBoxAndLabelConfig(9, 3) = Null
    avarTextBoxAndLabelConfig(0, 4) = "lbl4"
        avarTextBoxAndLabelConfig(1, 4) = "Priorität"
        avarTextBoxAndLabelConfig(2, 4) = "txt4"
        avarTextBoxAndLabelConfig(3, 4) = "PrioritaetKey"
        avarTextBoxAndLabelConfig(4, 4) = 0
        avarTextBoxAndLabelConfig(5, 4) = False
        avarTextBoxAndLabelConfig(6, 4) = True
        avarTextBoxAndLabelConfig(7, 4) = ""
        avarTextBoxAndLabelConfig(8, 4) = True
        avarTextBoxAndLabelConfig(9, 4) = Null
    avarTextBoxAndLabelConfig(0, 5) = "lbl5"
        avarTextBoxAndLabelConfig(1, 5) = "Parent"
        avarTextBoxAndLabelConfig(2, 5) = "txt5"
        avarTextBoxAndLabelConfig(3, 5) = "ParentKey"
        avarTextBoxAndLabelConfig(4, 5) = 0
        avarTextBoxAndLabelConfig(5, 5) = False
        avarTextBoxAndLabelConfig(6, 5) = True
        avarTextBoxAndLabelConfig(7, 5) = ""
        avarTextBoxAndLabelConfig(8, 5) = True
        avarTextBoxAndLabelConfig(9, 5) = Null
    avarTextBoxAndLabelConfig(0, 6) = "lbl6"
        avarTextBoxAndLabelConfig(1, 6) = "Bemerkung"
        avarTextBoxAndLabelConfig(2, 6) = "txt6"
        avarTextBoxAndLabelConfig(3, 6) = "Bemerkung"
        avarTextBoxAndLabelConfig(4, 6) = 1
        avarTextBoxAndLabelConfig(5, 6) = False
        avarTextBoxAndLabelConfig(6, 6) = False
        avarTextBoxAndLabelConfig(7, 6) = ""
        avarTextBoxAndLabelConfig(8, 6) = True
        avarTextBoxAndLabelConfig(9, 6) = Null
    avarTextBoxAndLabelConfig(0, 7) = "lbl7"
        avarTextBoxAndLabelConfig(1, 7) = "Beginn (Soll)"
        avarTextBoxAndLabelConfig(2, 7) = "txt7"
        avarTextBoxAndLabelConfig(3, 7) = "BeginnSoll"
        avarTextBoxAndLabelConfig(4, 7) = 0
        avarTextBoxAndLabelConfig(5, 7) = False
        avarTextBoxAndLabelConfig(6, 7) = True
        avarTextBoxAndLabelConfig(7, 7) = "Short Date"
        avarTextBoxAndLabelConfig(8, 7) = True
        avarTextBoxAndLabelConfig(9, 7) = Null
    avarTextBoxAndLabelConfig(0, 8) = "lbl8"
        avarTextBoxAndLabelConfig(1, 8) = "Ende (Soll)"
        avarTextBoxAndLabelConfig(2, 8) = "txt8"
        avarTextBoxAndLabelConfig(3, 8) = "EndeSoll"
        avarTextBoxAndLabelConfig(4, 8) = 0
        avarTextBoxAndLabelConfig(5, 8) = False
        avarTextBoxAndLabelConfig(6, 8) = True
        avarTextBoxAndLabelConfig(7, 8) = "Short Date"
        avarTextBoxAndLabelConfig(8, 8) = True
        avarTextBoxAndLabelConfig(9, 8) = Null
    avarTextBoxAndLabelConfig(0, 9) = "lbl9"
        avarTextBoxAndLabelConfig(1, 9) = "Kunde"
        avarTextBoxAndLabelConfig(2, 9) = "txt9"
        avarTextBoxAndLabelConfig(3, 9) = "Kunde"
        avarTextBoxAndLabelConfig(4, 9) = 1
        avarTextBoxAndLabelConfig(5, 9) = False
        avarTextBoxAndLabelConfig(6, 9) = False
        avarTextBoxAndLabelConfig(7, 9) = ""
        avarTextBoxAndLabelConfig(8, 9) = True
        avarTextBoxAndLabelConfig(9, 9) = Null
    avarTextBoxAndLabelConfig(0, 10) = Null
        avarTextBoxAndLabelConfig(1, 10) = Null
        avarTextBoxAndLabelConfig(2, 10) = "txt10"
        avarTextBoxAndLabelConfig(3, 10) = Null
        avarTextBoxAndLabelConfig(4, 10) = 1
        avarTextBoxAndLabelConfig(5, 10) = False
        avarTextBoxAndLabelConfig(6, 10) = False
        avarTextBoxAndLabelConfig(7, 10) = ""
        avarTextBoxAndLabelConfig(8, 10) = True
        avarTextBoxAndLabelConfig(9, 10) = ""
        
        ' 0 = object Name
        ' 1 = object caption
        ' 2 = object visible
    avarCommandButtonConfig(0, 0) = "cmd0"
        avarCommandButtonConfig(1, 0) = "Hauptmenü"
        avarCommandButtonConfig(2, 0) = True
    avarCommandButtonConfig(0, 1) = "cmd1"
        avarCommandButtonConfig(1, 1) = "Suchen"
        avarCommandButtonConfig(2, 1) = True
    avarCommandButtonConfig(0, 2) = "cmd2"
        ' avarCommandButtonConfig(1, 2) = "Angebot erfassen"
        avarCommandButtonConfig(1, 2) = "Auftrag erstellen"
        avarCommandButtonConfig(2, 2) = True
    avarCommandButtonConfig(0, 3) = "cmd3"
        avarCommandButtonConfig(1, 3) = "Auftrag erteilen"
        avarCommandButtonConfig(2, 3) = False
    avarCommandButtonConfig(0, 4) = "cmd4"
        avarCommandButtonConfig(1, 4) = "Leistung abnehmen"
        avarCommandButtonConfig(2, 4) = False
    avarCommandButtonConfig(0, 5) = "cmd5"
        avarCommandButtonConfig(1, 5) = "Rechnung erfassen"
        avarCommandButtonConfig(2, 5) = False
    avarCommandButtonConfig(0, 6) = "cmd6"
        avarCommandButtonConfig(1, 6) = "Auftrag stornieren"
        avarCommandButtonConfig(2, 6) = False
    avarCommandButtonConfig(0, 7) = "cmd7"
        avarCommandButtonConfig(1, 7) = "Auftragdetails anzeigen"
        avarCommandButtonConfig(2, 7) = False
    avarCommandButtonConfig(0, 8) = "cmd8"
        avarCommandButtonConfig(1, 8) = "Liefergegenstand anzeigen"
        avarCommandButtonConfig(2, 8) = False
    avarCommandButtonConfig(0, 9) = "cmd9"
        avarCommandButtonConfig(1, 9) = "Speichern"
        avarCommandButtonConfig(2, 9) = True
    avarCommandButtonConfig(0, 10) = "cmd10"
        avarCommandButtonConfig(1, 10) = "leer"
        avarCommandButtonConfig(2, 10) = False
    avarCommandButtonConfig(0, 11) = "cmd11"
        avarCommandButtonConfig(1, 11) = "Speichern"
        avarCommandButtonConfig(2, 11) = False
        
    ' subform configuration
    ' 0 = object Name
    ' 1 = source Object
    avarSubFormConfig(0, 0) = "frb1"
        avarSubFormConfig(1, 0) = "frmAuftragSuchenSub"
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.FormConfiguration: avarTextBoxAndLabelConfig, avarCommandButtonConfig and avarSubFormConfig initiated"
    End If
End Sub

' open frmSearchMain and set textboxes and labels
' feasible value: AuftragSuchen, AngebotSuchen

Public Sub OpenFormSearchMain(ByVal strMode As String)
    DoCmd.OpenForm "frmSearchMain", acNormal
    
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.OpenFormSearchMain ausfuehren"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' select form mode
    Select Case strMode
        Case "AuftragSuchen"
            FormConfiguration
        Case "AngebotSuchen"
            LiefergegenstandSuchenConfig
        Case Else
            Debug.Print "basSearchMain.OpenForm: Uebergebener " & _
                "Parameter 'strMode' nicht im Wertevorrat enthalten"
    End Select
    
    ' set focus
    Dim strFocus As String
    strFocus = basSearchMain.AuftragFocus
    Forms.Item(strFormName).Controls.Item(strFocus).SetFocus
    
    ' set labels and textboxes
    Dim varLabelsAndTextboxes As Variant
    varLabelsAndTextboxes = basSearchMain.AuftragTextboxesAndLabels
    
    ' set commandbuttons
    Dim varCommandButtons As Variant
    varCommandButtons = basSearchMain.AuftragCommandButtons
    
    ' set subform
    Dim varSubForm As Variant
    varSubForm = basSearchMain.AuftragSubForm
    
    ' set textboxes and labels
    Dim inti As Integer
    ' skip columnnames
    For inti = LBound(varLabelsAndTextboxes, 2) + 1 To UBound(varLabelsAndTextboxes, 2)
        ' set label caption
            If Not IsNull(varLabelsAndTextboxes(1, inti)) Then
                Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(0, inti)).Caption = varLabelsAndTextboxes(1, inti)
            End If
        ' set textbox border style
        Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).BorderStyle = varLabelsAndTextboxes(4, inti)
        ' set textbox ishyperlink
        Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).IsHyperlink = varLabelsAndTextboxes(5, inti)
        ' set textbox locked
        Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).Locked = varLabelsAndTextboxes(6, inti)
        ' set textbox format
        Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).Format = varLabelsAndTextboxes(7, inti)
        ' set textbox visible
        Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).Visible = varLabelsAndTextboxes(8, inti)
        ' set textbox defaultValue
            If Not IsNull(varLabelsAndTextboxes(9, inti)) Then
                Forms.Item(strFormName).Controls.Item(varLabelsAndTextboxes(2, inti)).DefaultValue = varLabelsAndTextboxes(9, inti)
            End If
    Next
        
    ' set command buttons
    ' skip column names
    For inti = LBound(varCommandButtons, 2) + 1 To UBound(varCommandButtons, 2)
        ' set caption
        Forms.Item(strFormName).Controls.Item(varCommandButtons(0, inti)).Caption = varCommandButtons(1, inti)
        ' set visibility
        Forms.Item(strFormName).Controls.Item(varCommandButtons(0, inti)).Visible = varCommandButtons(2, inti)
    Next
                
    ' set subform
    ' skip column names
    For inti = LBound(varSubForm, 2) + 1 To UBound(varSubForm, 2)
        Forms.Item(strFormName).Controls.Item(varSubForm(0, inti)).SourceObject = varSubForm(1, inti)
    Next
    
End Sub

' check if form is loaded
' assign values to textboxes via avarTextBoxAndLabelConfig
' Public Sub ShowRecordset(ByVal strRecordsetName As String)
Public Sub ShowRecordset(ByVal varRecordsetName As Variant)
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.ShowRecordset ausfuehren" & vbCrLf & _
            "basSearchMain.ShowRecordset: varRecordsetName = " & varRecordsetName
    End If
    
    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
        
    ' initiate class Auftrag
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' set Auftrag to selected Recordset
    Auftrag.SelectRecordset varRecordsetName
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' error handler, case strFormName is not loaded
    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        Debug.Print "basSearchMain.ShowRecordset: Formular " & strFormName & _
            " ist nicht geoeffnet, Prozedur abgebrochen"
        Exit Sub
    End If
    
    Dim varTextBoxesAndLabels As Variant
    varTextBoxesAndLabels = basSearchMain.AuftragTextboxesAndLabels
    
    ' assign values to textboxes
    ' skip column names
    Dim inti As Integer
    For inti = LBound(varTextBoxesAndLabels, 2) + 1 To UBound(varTextBoxesAndLabels, 2)
        ' handler in case field value is null
        ' IsEmpty is neccesary because opening frmSearchMain will open frmAuftragSuchenSub,
        ' at a time when varTextboxesAndLabels is not set yet
        If Not IsEmpty(varTextBoxesAndLabels(3, inti)) And Not IsNull(varTextBoxesAndLabels(3, inti)) Then
            Forms.Item(strFormName).Controls.Item(varTextBoxesAndLabels(2, inti)) = CallByName(Auftrag, varTextBoxesAndLabels(3, inti), VbGet)
        End If
    Next
End Sub

Public Sub SaveAuftrag()
    ' initiate object
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    Dim varTextBoxesAndLabels As Variant
    varTextBoxesAndLabels = basSearchMain.AuftragTextboxesAndLabels
    
    ' write textbox values to class properties, skip ID
    Dim inti As Integer
    For inti = LBound(varTextBoxesAndLabels, 2) + 1 To UBound(varTextBoxesAndLabels, 2)
        ' null value handler
        If Not IsNull(varTextBoxesAndLabels(3, inti)) Then
            CallByName Auftrag, varTextBoxesAndLabels(3, inti), VbLet, Forms.Item("frmSearchMain").Controls(varTextBoxesAndLabels(2, inti))
        End If
    Next
    
    ' save Auftrag
    Auftrag.SaveRecordset Forms.Item("frmSearchMain").Controls(varTextBoxesAndLabels(2, 1))
End Sub

Public Sub SearchAuftrag()
    
    ' set name of resulting query
    Dim strResultQueryName As String
    strResultQueryName = "qryAuftragAuswahl"
    
    ' set name of origin
    Dim strQuerySourceName As String
    strQuerySourceName = "qryAuftrag"
    
    ' initiate
    Dim varSearchTerm As Variant
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' set textbox name
    Dim strTextboxName As String
    strTextboxName = "txt10"
    
    varSearchTerm = Forms.Item(strFormName).Controls.Item(strTextboxName)
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' set query definition list
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    ' initiate recordset
    Dim RecordSet As Object
    
    ' set behavior, when input is empty
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
    
    ' close resulting query if loaded
    If SysCmd(acSysCmdGetObjectState, acQuery, strResultQueryName) = 1 Then
        DoCmd.Close acQuery, strResultQueryName, acSaveYes
        ' verbatim message
        If gconVerbatim = True Then
            Debug.Print "basSearchMain.SearchAuftrag: " & strResultQueryName & " geschlossen"
        End If
    End If
    
    ' delete query
    For Each RecordSet In dbsCurrentDB.QueryDefs
        If RecordSet.Name = strResultQueryName Then
            DoCmd.DeleteObject acQuery, strResultQueryName
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "basSearchMain.SearchAuftrag: " & strResultQueryName & " geloescht"
            End If
        End If
    Next RecordSet
    
    ' create SQL-code
    With qdfQuery
        .SQL = " SELECT " & strQuerySourceName & ".AftrID, " & strQuerySourceName & ".AftrTitel, " & strQuerySourceName & ".ParentKey, " & strQuerySourceName & ".Bemerkung" _
                & " FROM " & strQuerySourceName & "" _
                & " WHERE " & strQuerySourceName & ".AftrID LIKE '*" & varSearchTerm & "*' OR " & strQuerySourceName & ".AftrTitel LIKE '*" & varSearchTerm & "*' OR " & strQuerySourceName & ".ParentKey LIKE '*" & varSearchTerm & "*' OR " & strQuerySourceName & ".ParentKey LIKE '*" & varSearchTerm & "*'" _
                & " ;"
        .Name = strResultQueryName
    End With
    
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.SearchAuftrag: ausgefuehrt, varSearchTerm: " & varSearchTerm
    End If
    
ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
        qdfQuery.Close
        Set qdfQuery = Nothing
End Sub


