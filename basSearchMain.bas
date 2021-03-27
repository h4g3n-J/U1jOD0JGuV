Attribute VB_Name = "basSearchMain"
' basSearchMain

Option Compare Database
Option Explicit

' adjust Auftrag Focus
Private Function FocusSettings() As String
    FocusSettings = "txt0"
End Function

' adjust Auftrag textboxes and labels
Private Function TextboxAndLabelSettings() As Variant
    Dim varSettings(11, 9) As Variant
    varSettings(0, 0) = "label name"
        varSettings(0, 1) = "label caption"
        varSettings(0, 2) = "textbox name"
        varSettings(0, 3) = "recordset property / value"
        varSettings(0, 4) = "textbox border style"
        varSettings(0, 5) = "textbox ishyperlink"
        varSettings(0, 6) = "textbox locked"
        varSettings(0, 7) = "textbox format"
        varSettings(0, 8) = "textbox visible"
        varSettings(0, 9) = "textbox defaultValue"
    varSettings(1, 0) = "lbl0"
        varSettings(1, 1) = "ID"
        varSettings(1, 2) = "txt0"
        varSettings(1, 3) = "AftrID"
        varSettings(1, 4) = 0
        varSettings(1, 5) = False
        varSettings(1, 6) = True
        varSettings(1, 7) = ""
        varSettings(1, 8) = True
        varSettings(1, 9) = Null
    varSettings(2, 0) = "lbl1"
        varSettings(2, 1) = "Titel"
        varSettings(2, 2) = "txt1"
        varSettings(2, 3) = "AftrTitel"
        varSettings(2, 4) = 1
        varSettings(2, 5) = False
        varSettings(2, 6) = False
        varSettings(2, 7) = ""
        varSettings(2, 8) = True
        varSettings(2, 9) = Null
    varSettings(3, 0) = "lbl2"
        varSettings(3, 1) = "ICD Status"
        varSettings(3, 2) = "txt2"
        varSettings(3, 3) = "StatusKey"
        varSettings(3, 4) = 0
        varSettings(3, 5) = False
        varSettings(3, 6) = True
        varSettings(3, 7) = ""
        varSettings(3, 8) = True
        varSettings(3, 9) = Null
    varSettings(4, 0) = "lbl3"
        varSettings(4, 1) = "Owner"
        varSettings(4, 2) = "txt3"
        varSettings(4, 3) = "OwnerKey"
        varSettings(4, 4) = 0
        varSettings(4, 5) = False
        varSettings(4, 6) = True
        varSettings(4, 7) = ""
        varSettings(4, 8) = True
        varSettings(4, 9) = Null
    varSettings(5, 0) = "lbl4"
        varSettings(5, 1) = "Priorität"
        varSettings(5, 2) = "txt4"
        varSettings(5, 3) = "PrioritaetKey"
        varSettings(5, 4) = 0
        varSettings(5, 5) = False
        varSettings(5, 6) = True
        varSettings(5, 7) = ""
        varSettings(5, 8) = True
        varSettings(5, 9) = Null
    varSettings(6, 0) = "lbl5"
        varSettings(6, 1) = "Parent"
        varSettings(6, 2) = "txt5"
        varSettings(6, 3) = "ParentKey"
        varSettings(6, 4) = 0
        varSettings(6, 5) = False
        varSettings(6, 6) = True
        varSettings(6, 7) = ""
        varSettings(6, 8) = True
        varSettings(6, 9) = Null
    varSettings(7, 0) = "lbl6"
        varSettings(7, 1) = "Bemerkung"
        varSettings(7, 2) = "txt6"
        varSettings(7, 3) = "Bemerkung"
        varSettings(7, 4) = 1
        varSettings(7, 5) = False
        varSettings(7, 6) = False
        varSettings(7, 7) = ""
        varSettings(7, 8) = True
        varSettings(7, 9) = Null
    varSettings(8, 0) = "lbl7"
        varSettings(8, 1) = "Beginn (Soll)"
        varSettings(8, 2) = "txt7"
        varSettings(8, 3) = "BeginnSoll"
        varSettings(8, 4) = 0
        varSettings(8, 5) = False
        varSettings(8, 6) = True
        varSettings(8, 7) = "Short Date"
        varSettings(8, 8) = True
        varSettings(8, 9) = Null
    varSettings(9, 0) = "lbl8"
        varSettings(9, 1) = "Ende (Soll)"
        varSettings(9, 2) = "txt8"
        varSettings(9, 3) = "EndeSoll"
        varSettings(9, 4) = 0
        varSettings(9, 5) = False
        varSettings(9, 6) = True
        varSettings(9, 7) = "Short Date"
        varSettings(9, 8) = True
        varSettings(9, 9) = Null
    varSettings(10, 0) = "lbl9"
        varSettings(10, 1) = "Kunde"
        varSettings(10, 2) = "txt9"
        varSettings(10, 3) = "Kunde"
        varSettings(10, 4) = 1
        varSettings(10, 5) = False
        varSettings(10, 6) = False
        varSettings(10, 7) = ""
        varSettings(10, 8) = True
        varSettings(10, 9) = Null
    varSettings(11, 0) = Null
        varSettings(11, 1) = Null
        varSettings(11, 2) = "txt10"
        varSettings(11, 3) = Null
        varSettings(11, 4) = 1
        varSettings(11, 5) = False
        varSettings(11, 6) = False
        varSettings(11, 7) = ""
        varSettings(11, 8) = True
        varSettings(11, 9) = ""
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.TextboxAndLabelSettings ausgefuehrt"
    End If
    
    TextboxAndLabelSettings = varSettings
        
End Function

' adjust Auftrag commandbuttons
Private Function CommandButtonSettings() As Variant
    Dim varSettings(12, 2) As Variant
    varSettings(0, 0) = "object Name"
        varSettings(0, 1) = "object caption"
        varSettings(0, 2) = "object visible"
    varSettings(1, 0) = "cmd0"
        varSettings(1, 1) = "Hauptmenü"
        varSettings(1, 2) = True
    varSettings(2, 0) = "cmd1"
        varSettings(2, 1) = "Suchen"
        varSettings(2, 2) = True
    varSettings(3, 0) = "cmd2"
        varSettings(3, 1) = "Auftrag erstellen"
        varSettings(3, 2) = True
    varSettings(4, 0) = "cmd3"
        varSettings(4, 1) = "Auftrag erteilen"
        varSettings(4, 2) = False
    varSettings(5, 0) = "cmd4"
        varSettings(5, 1) = "Leistung abnehmen"
        varSettings(5, 2) = False
    varSettings(6, 0) = "cmd5"
        varSettings(6, 1) = "Rechnung erfassen"
        varSettings(6, 2) = False
    varSettings(7, 0) = "cmd6"
        varSettings(7, 1) = "Auftrag stornieren"
        varSettings(7, 2) = False
    varSettings(8, 0) = "cmd7"
        varSettings(8, 1) = "Auftragdetails anzeigen"
        varSettings(8, 2) = False
    varSettings(9, 0) = "cmd8"
        varSettings(9, 1) = "Liefergegenstand anzeigen"
        varSettings(9, 2) = False
    varSettings(10, 0) = "cmd9"
        varSettings(10, 1) = "Speichern"
        varSettings(10, 2) = True
    varSettings(11, 0) = "cmd10"
        varSettings(11, 1) = "leer"
        varSettings(11, 2) = False
    varSettings(12, 0) = "cmd11"
        varSettings(12, 1) = "Speichern"
        varSettings(12, 2) = False
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.CommandButtonSettings ausgefuehrt"
    End If
        
    CommandButtonSettings = varSettings
End Function

' adjust Auftrag commandbuttons
Private Function SubFormSettings() As Variant
    Dim varSettings(1, 1) As Variant
    varSettings(0, 0) = "object Name"
        varSettings(0, 1) = "source"
    varSettings(1, 0) = "frb1"
        varSettings(1, 1) = "frmAuftragSuchenSub"
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.SubFormSettings ausgefuehrt"
    End If
        
    SubFormSettings = varSettings
End Function

Private Sub LiefergegenstandSuchenSettings()
    ' set focus
    mstrFocus = "txt0"
    
    Dim varSettings As Variant
    varSettings(0, 0) = "label name"
        varSettings(0, 1) = "label caption"
        varSettings(0, 2) = "textbox name"
        varSettings(0, 3) = "textbox value"
        varSettings(0, 4) = "textbox border style"
        varSettings(0, 5) = "textbox ishyperlink"
        varSettings(0, 6) = "textbox locked"
        varSettings(0, 7) = "textbox format"
        varSettings(0, 8) = "textbox visible"
        varSettings(0, 9) = "textbox defaultValue"
    varSettings(1, 0) = "lbl0"
        varSettings(1, 1) = "ID"
        varSettings(1, 2) = "txt0"
        varSettings(1, 3) = "AftrID"
        varSettings(1, 4) = 0
        varSettings(1, 5) = False
        varSettings(1, 6) = True
        varSettings(1, 7) = ""
        varSettings(1, 8) = True
        varSettings(1, 9) = Null
    varSettings(2, 0) = "lbl1"
        varSettings(2, 1) = "Titel"
        varSettings(2, 2) = "txt1"
        varSettings(2, 3) = "AftrTitel"
        varSettings(2, 4) = 1
        varSettings(2, 5) = False
        varSettings(2, 6) = False
        varSettings(2, 7) = ""
        varSettings(2, 8) = True
        varSettings(2, 9) = Null
    varSettings(3, 0) = "lbl2"
        varSettings(3, 1) = "ICD Status"
        varSettings(3, 2) = "txt2"
        varSettings(3, 3) = "StatusKey"
        varSettings(3, 4) = 0
        varSettings(3, 5) = False
        varSettings(3, 6) = True
        varSettings(3, 7) = ""
        varSettings(3, 8) = True
        varSettings(3, 9) = Null
    varSettings(4, 0) = "lbl3"
        varSettings(4, 1) = "Owner"
        varSettings(4, 2) = "txt3"
        varSettings(4, 3) = "OwnerKey"
        varSettings(4, 4) = 0
        varSettings(4, 5) = False
        varSettings(4, 6) = True
        varSettings(4, 7) = ""
        varSettings(4, 8) = True
        varSettings(4, 9) = Null
    varSettings(5, 0) = "lbl4"
        varSettings(5, 1) = "Priorität"
        varSettings(5, 2) = "txt4"
        varSettings(5, 3) = "PrioritaetKey"
        varSettings(5, 4) = 0
        varSettings(5, 5) = False
        varSettings(5, 6) = True
        varSettings(5, 7) = ""
        varSettings(5, 8) = True
        varSettings(5, 9) = Null
    varSettings(6, 0) = "lbl5"
        varSettings(6, 1) = "Parent"
        varSettings(6, 2) = "txt5"
        varSettings(6, 3) = "ParentKey"
        varSettings(6, 4) = 0
        varSettings(6, 5) = False
        varSettings(6, 6) = True
        varSettings(6, 7) = ""
        varSettings(6, 8) = True
        varSettings(6, 9) = Null
    varSettings(7, 0) = "lbl6"
        varSettings(7, 1) = "Bemerkung"
        varSettings(7, 2) = "txt6"
        varSettings(7, 3) = "Bemerkung"
        varSettings(7, 4) = 1
        varSettings(7, 5) = False
        varSettings(7, 6) = False
        varSettings(7, 7) = ""
        varSettings(7, 8) = True
        varSettings(7, 9) = Null
    varSettings(8, 0) = "lbl7"
        varSettings(8, 1) = "Beginn (Soll)"
        varSettings(8, 2) = "txt7"
        varSettings(8, 3) = "BeginnSoll"
        varSettings(8, 4) = 0
        varSettings(8, 5) = False
        varSettings(8, 6) = True
        varSettings(8, 7) = "Short Date"
        varSettings(8, 8) = True
        varSettings(8, 9) = Null
    varSettings(9, 0) = "lbl8"
        varSettings(9, 1) = "Ende (Soll)"
        varSettings(9, 2) = "txt8"
        varSettings(9, 3) = "EndeSoll"
        varSettings(9, 4) = 0
        varSettings(9, 5) = False
        varSettings(9, 6) = True
        varSettings(9, 7) = "Short Date"
        varSettings(9, 8) = True
        varSettings(9, 9) = Null
    varSettings(10, 0) = "lbl9"
        varSettings(10, 1) = "Kunde"
        varSettings(10, 2) = "txt9"
        varSettings(10, 3) = "Kunde"
        varSettings(10, 4) = 1
        varSettings(10, 5) = False
        varSettings(10, 6) = False
        varSettings(10, 7) = ""
        varSettings(10, 8) = True
        varSettings(10, 9) = Null
    varSettings(11, 0) = Null
        varSettings(11, 1) = Null
        varSettings(11, 2) = "txt10"
        varSettings(11, 3) = Null
        varSettings(11, 4) = 1
        varSettings(11, 5) = False
        varSettings(11, 6) = False
        varSettings(11, 7) = ""
        varSettings(11, 8) = True
        varSettings(11, 9) = ""
        
    avarCommandButtonConfig(0, 0) = "object Name"
        avarCommandButtonConfig(0, 1) = "object caption"
        avarCommandButtonConfig(0, 2) = "object visible"
    avarCommandButtonConfig(1, 0) = "cmd0"
        avarCommandButtonConfig(1, 1) = "Hauptmenü"
        avarCommandButtonConfig(1, 2) = True
    avarCommandButtonConfig(2, 0) = "cmd1"
        avarCommandButtonConfig(2, 1) = "Suchen"
        avarCommandButtonConfig(2, 2) = True
    avarCommandButtonConfig(3, 0) = "cmd2"
        ' avarCommandButtonConfig(3, 1) = "Angebot erfassen"
        avarCommandButtonConfig(3, 1) = "Auftrag erstellen"
        avarCommandButtonConfig(3, 2) = True
    avarCommandButtonConfig(4, 0) = "cmd3"
        avarCommandButtonConfig(4, 1) = "Auftrag erteilen"
        avarCommandButtonConfig(4, 2) = False
    avarCommandButtonConfig(5, 0) = "cmd4"
        avarCommandButtonConfig(5, 1) = "Leistung abnehmen"
        avarCommandButtonConfig(5, 2) = False
    avarCommandButtonConfig(6, 0) = "cmd5"
        avarCommandButtonConfig(6, 1) = "Rechnung erfassen"
        avarCommandButtonConfig(6, 2) = False
    avarCommandButtonConfig(7, 0) = "cmd6"
        avarCommandButtonConfig(7, 1) = "Auftrag stornieren"
        avarCommandButtonConfig(7, 2) = False
    avarCommandButtonConfig(8, 0) = "cmd7"
        avarCommandButtonConfig(8, 1) = "Auftragdetails anzeigen"
        avarCommandButtonConfig(8, 2) = False
    avarCommandButtonConfig(9, 0) = "cmd8"
        avarCommandButtonConfig(9, 1) = "Liefergegenstand anzeigen"
        avarCommandButtonConfig(9, 2) = False
    avarCommandButtonConfig(10, 0) = "cmd9"
        avarCommandButtonConfig(10, 1) = "Speichern"
        avarCommandButtonConfig(10, 2) = True
    avarCommandButtonConfig(11, 0) = "cmd10"
        avarCommandButtonConfig(11, 1) = "leer"
        avarCommandButtonConfig(11, 2) = False
    avarCommandButtonConfig(12, 0) = "cmd11"
        avarCommandButtonConfig(12, 1) = "Speichern"
        avarCommandButtonConfig(12, 2) = False
        
    ' subform configuration
    avarSubFormConfig(0, 0) = "object Name"
        avarSubFormConfig(0, 1) = "source Object"
    avarSubFormConfig(1, 0) = "frb1"
        avarSubFormConfig(1, 1) = "frmAuftragSuchenSub"
        
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.FormConfiguration: varSettings, avarCommandButtonConfig und avarSubFormConfig ausgefuehrt"
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
    
    ' declare variables for
    ' labels and textboxes
        Dim varTextBoxAndLabelSettings As Variant
    ' commandbuttons
        Dim varCommandButtonSettings As Variant
    ' subform
        Dim varSubFormSettings As Variant
    ' focus
        Dim strFocusSettings As String
    
    Select Case strMode
        Case "AuftragSuchen"
            ' commit labels and textboxes
            varTextBoxAndLabelSettings = basSearchMain.TextboxAndLabelSettings
            ' commit command buttons
            varCommandButtonSettings = basSearchMain.CommandButtonSettings
            ' commit subform
            varSubFormSettings = basSearchMain.SubFormSettings
            ' commit focus
            strFocusSettings = basSearchMain.FocusSettings
        Case "AngebotSuchen"
            LiefergegenstandSuchenSettings
        Case Else
            Debug.Print "basSearchMain.OpenForm: Uebergebener " & _
                "Parameter 'strMode' nicht im Wertevorrat enthalten"
    End Select
    
    ' set focus
    Forms.Item(strFormName).Controls.Item(strFocusSettings).SetFocus
    
    ' set textboxes and labels
    Dim inti As Integer
    ' skip column names
    For inti = LBound(varTextBoxAndLabelSettings, 1) + 1 To UBound(varTextBoxAndLabelSettings, 1)
        ' set label caption
            If Not IsNull(varTextBoxAndLabelSettings(inti, 1)) Then
                Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 0)).Caption = varTextBoxAndLabelSettings(inti, 1)
            End If
        ' set textbox border style
        Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).BorderStyle = varTextBoxAndLabelSettings(inti, 4)
        ' set textbox ishyperlink
        Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).IsHyperlink = varTextBoxAndLabelSettings(inti, 5)
        ' set textbox locked
        Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).Locked = varTextBoxAndLabelSettings(inti, 6)
        ' set textbox format
        Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).Format = varTextBoxAndLabelSettings(inti, 7)
        ' set textbox visible
        Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).Visible = varTextBoxAndLabelSettings(inti, 8)
        ' set textbox defaultValue
            If Not IsNull(varTextBoxAndLabelSettings(inti, 9)) Then
                Forms.Item(strFormName).Controls.Item(varTextBoxAndLabelSettings(inti, 2)).DefaultValue = varTextBoxAndLabelSettings(inti, 9)
            End If
    Next
        
    ' set command buttons
    ' skip column names
    For inti = LBound(varCommandButtonSettings, 1) + 1 To UBound(varCommandButtonSettings, 1)
        ' set caption
        Forms.Item(strFormName).Controls.Item(varCommandButtonSettings(inti, 0)).Caption = varCommandButtonSettings(inti, 1)
        ' set visibility
        Forms.Item(strFormName).Controls.Item(varCommandButtonSettings(inti, 0)).Visible = varCommandButtonSettings(inti, 2)
    Next
                
    ' set subform
    ' skip column names
    For inti = LBound(varSubFormSettings, 1) + 1 To UBound(varSubFormSettings, 1)
        Forms.Item(strFormName).Controls.Item(varSubFormSettings(inti, 0)).SourceObject = varSubFormSettings(inti, 1)
    Next
    
End Sub

Public Sub SaveAuftrag()
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basSearchMain.SaveAuftrag ausfuehren"
    End If
    
    ' initiate auftrag
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' load textbox and label settings
    Dim varTextBoxAndLabelSettings As Variant
    varTextBoxAndLabelSettings = basSearchMain.TextboxAndLabelSettings
    
    ' write textbox values to class properties, skip ID
    Dim inti As Integer
    ' move through textboxes, write detected values to recordset
    For inti = LBound(varTextBoxAndLabelSettings, 1) + 1 To UBound(varTextBoxAndLabelSettings, 1)
        ' null value handler
        If Not IsNull(varTextBoxAndLabelSettings(inti, 3)) Then
            ' select textbox by name [varTextBoxAndLabelSettings(inti, 2)] then _
            ' write value to recordset property [varTextBoxAndLabelSettings(inti, 3)]
            CallByName Auftrag, varTextBoxAndLabelSettings(inti, 3), VbLet, Forms.Item("frmSearchMain").Controls(varTextBoxAndLabelSettings(inti, 2))
        End If
    Next
    
    ' save Auftrag
    Auftrag.SaveRecordset Forms.Item("frmSearchMain").Controls(varTextBoxAndLabelSettings(1, 2))
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


