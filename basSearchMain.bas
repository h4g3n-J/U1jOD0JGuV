Attribute VB_Name = "basSearchMain"
' basSearchMain

Option Compare Database
Option Explicit

Private avarTextBoxAndLabelConfig(8, 10) As Variant
Private avarCommandButtonConfig(2, 11) As Variant
Private avarSubFormConfig(1, 0) As Variant
Private mstrFocus As String

Private Sub FormConfiguration()
        
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
    avarTextBoxAndLabelConfig(0, 0) = "lbl0"
        avarTextBoxAndLabelConfig(1, 0) = "ID"
        avarTextBoxAndLabelConfig(2, 0) = "txt0"
        avarTextBoxAndLabelConfig(3, 0) = "AftrID"
        avarTextBoxAndLabelConfig(4, 0) = 0
        avarTextBoxAndLabelConfig(5, 0) = False
        avarTextBoxAndLabelConfig(6, 0) = True
        avarTextBoxAndLabelConfig(7, 0) = ""
        avarTextBoxAndLabelConfig(8, 0) = True
    avarTextBoxAndLabelConfig(0, 1) = "lbl1"
        avarTextBoxAndLabelConfig(1, 1) = "Titel"
        avarTextBoxAndLabelConfig(2, 1) = "txt1"
        avarTextBoxAndLabelConfig(3, 1) = "AftrTitel"
        avarTextBoxAndLabelConfig(4, 1) = 1
        avarTextBoxAndLabelConfig(5, 1) = False
        avarTextBoxAndLabelConfig(6, 1) = False
        avarTextBoxAndLabelConfig(7, 1) = ""
        avarTextBoxAndLabelConfig(8, 1) = True
    avarTextBoxAndLabelConfig(0, 2) = "lbl2"
        avarTextBoxAndLabelConfig(1, 2) = "ICD Status"
        avarTextBoxAndLabelConfig(2, 2) = "txt2"
        avarTextBoxAndLabelConfig(3, 2) = "StatusKey"
        avarTextBoxAndLabelConfig(4, 2) = 0
        avarTextBoxAndLabelConfig(5, 2) = False
        avarTextBoxAndLabelConfig(6, 2) = True
        avarTextBoxAndLabelConfig(7, 2) = ""
        avarTextBoxAndLabelConfig(8, 2) = True
    avarTextBoxAndLabelConfig(0, 3) = "lbl3"
        avarTextBoxAndLabelConfig(1, 3) = "Owner"
        avarTextBoxAndLabelConfig(2, 3) = "txt3"
        avarTextBoxAndLabelConfig(3, 3) = "OwnerKey"
        avarTextBoxAndLabelConfig(4, 3) = 0
        avarTextBoxAndLabelConfig(5, 3) = False
        avarTextBoxAndLabelConfig(6, 3) = True
        avarTextBoxAndLabelConfig(7, 3) = ""
        avarTextBoxAndLabelConfig(8, 3) = True
    avarTextBoxAndLabelConfig(0, 4) = "lbl4"
        avarTextBoxAndLabelConfig(1, 4) = "Priorität"
        avarTextBoxAndLabelConfig(2, 4) = "txt4"
        avarTextBoxAndLabelConfig(3, 4) = "PrioritaetKey"
        avarTextBoxAndLabelConfig(4, 4) = 0
        avarTextBoxAndLabelConfig(5, 4) = False
        avarTextBoxAndLabelConfig(6, 4) = True
        avarTextBoxAndLabelConfig(7, 4) = ""
        avarTextBoxAndLabelConfig(8, 4) = True
    avarTextBoxAndLabelConfig(0, 5) = "lbl5"
        avarTextBoxAndLabelConfig(1, 5) = "Parent"
        avarTextBoxAndLabelConfig(2, 5) = "txt5"
        avarTextBoxAndLabelConfig(3, 5) = "ParentKey"
        avarTextBoxAndLabelConfig(4, 5) = 0
        avarTextBoxAndLabelConfig(5, 5) = False
        avarTextBoxAndLabelConfig(6, 5) = True
        avarTextBoxAndLabelConfig(7, 5) = ""
        avarTextBoxAndLabelConfig(8, 5) = True
    avarTextBoxAndLabelConfig(0, 6) = "lbl6"
        avarTextBoxAndLabelConfig(1, 6) = "Bemerkung"
        avarTextBoxAndLabelConfig(2, 6) = "txt6"
        avarTextBoxAndLabelConfig(3, 6) = "Bemerkung"
        avarTextBoxAndLabelConfig(4, 6) = 1
        avarTextBoxAndLabelConfig(5, 6) = False
        avarTextBoxAndLabelConfig(6, 6) = False
        avarTextBoxAndLabelConfig(7, 6) = ""
        avarTextBoxAndLabelConfig(8, 6) = True
    avarTextBoxAndLabelConfig(0, 7) = "lbl7"
        avarTextBoxAndLabelConfig(1, 7) = "Beginn (Soll)"
        avarTextBoxAndLabelConfig(2, 7) = "txt7"
        avarTextBoxAndLabelConfig(3, 7) = "BeginnSoll"
        avarTextBoxAndLabelConfig(4, 7) = 0
        avarTextBoxAndLabelConfig(5, 7) = False
        avarTextBoxAndLabelConfig(6, 7) = True
        avarTextBoxAndLabelConfig(7, 7) = "Short Date"
        avarTextBoxAndLabelConfig(8, 7) = True
    avarTextBoxAndLabelConfig(0, 8) = "lbl8"
        avarTextBoxAndLabelConfig(1, 8) = "Ende (Soll)"
        avarTextBoxAndLabelConfig(2, 8) = "txt8"
        avarTextBoxAndLabelConfig(3, 8) = "EndeSoll"
        avarTextBoxAndLabelConfig(4, 8) = 0
        avarTextBoxAndLabelConfig(5, 8) = False
        avarTextBoxAndLabelConfig(6, 8) = True
        avarTextBoxAndLabelConfig(7, 8) = "Short Date"
        avarTextBoxAndLabelConfig(8, 8) = True
    avarTextBoxAndLabelConfig(0, 9) = "lbl9"
        avarTextBoxAndLabelConfig(1, 9) = "Kunde"
        avarTextBoxAndLabelConfig(2, 9) = "txt9"
        avarTextBoxAndLabelConfig(3, 9) = "Kunde"
        avarTextBoxAndLabelConfig(4, 9) = 1
        avarTextBoxAndLabelConfig(5, 9) = False
        avarTextBoxAndLabelConfig(6, 9) = False
        avarTextBoxAndLabelConfig(7, 9) = ""
        avarTextBoxAndLabelConfig(8, 9) = True
    avarTextBoxAndLabelConfig(0, 10) = Null
        avarTextBoxAndLabelConfig(1, 10) = Null
        avarTextBoxAndLabelConfig(2, 10) = "txt10"
        avarTextBoxAndLabelConfig(3, 10) = Null
        avarTextBoxAndLabelConfig(4, 10) = 1
        avarTextBoxAndLabelConfig(5, 10) = False
        avarTextBoxAndLabelConfig(6, 10) = False
        avarTextBoxAndLabelConfig(7, 10) = ""
        avarTextBoxAndLabelConfig(8, 10) = True
        
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
Public Sub OpenFormAuftrag()
    DoCmd.OpenForm "frmSearchMain", acNormal
    
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.OpenFormAuftrag ausfuehren"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' initialize FormConfiguratoin
    FormConfiguration
    
    ' set focus
    Forms.Item(strFormName).Controls.Item(mstrFocus).SetFocus
    
    ' set labels and textboxes
    Dim inti As Integer
    For inti = LBound(avarTextBoxAndLabelConfig, 2) To UBound(avarTextBoxAndLabelConfig, 2)
        ' set label caption
            If Not IsNull(avarTextBoxAndLabelConfig(1, inti)) Then
                Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(0, inti)).Caption = avarTextBoxAndLabelConfig(1, inti)
            End If
        ' set textbox border style
        Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)).BorderStyle = avarTextBoxAndLabelConfig(4, inti)
        ' set textbox ishyperlink
        Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)).IsHyperlink = avarTextBoxAndLabelConfig(5, inti)
        ' set textbox locked
        Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)).Locked = avarTextBoxAndLabelConfig(6, inti)
        ' set textbox format
        Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)).Format = avarTextBoxAndLabelConfig(7, inti)
        ' set textbox visible
        Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)).Visible = avarTextBoxAndLabelConfig(8, inti)
    Next
        
    ' set command buttons
    For inti = LBound(avarCommandButtonConfig, 2) To UBound(avarCommandButtonConfig, 2)
        ' set caption
        Forms.Item(strFormName).Controls.Item(avarCommandButtonConfig(0, inti)).Caption = avarCommandButtonConfig(1, inti)
        ' set visibility
        Forms.Item(strFormName).Controls.Item(avarCommandButtonConfig(0, inti)).Visible = avarCommandButtonConfig(2, inti)
    Next
                
    ' set subform
    For inti = LBound(avarSubFormConfig, 2) To UBound(avarSubFormConfig, 2)
        Forms.Item(strFormName).Controls.Item(avarSubFormConfig(0, inti)).SourceObject = avarSubFormConfig(1, inti)
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
    
    ' assign values to textboxes
    Dim inti As Integer
    For inti = LBound(avarTextBoxAndLabelConfig, 2) To UBound(avarTextBoxAndLabelConfig, 2)
        ' handler in case field value is null
        ' IsEmpty is neccesary because opening frmSearchMain will open frmAuftragSuchenSub,
        ' at a time when avarTextBoxAndLabelConfig is not set yet
        If Not IsEmpty(avarTextBoxAndLabelConfig(3, inti)) And Not IsNull(avarTextBoxAndLabelConfig(3, inti)) Then
            Forms.Item(strFormName).Controls.Item(avarTextBoxAndLabelConfig(2, inti)) = CallByName(Auftrag, avarTextBoxAndLabelConfig(3, inti), VbGet)
        End If
    Next
End Sub

Public Sub SaveAuftrag()
    ' Propertys schreiben
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' write textbox values to class properties, skipp ID
    Dim inti As Integer
    For inti = LBound(avarTextBoxAndLabelConfig, 2) + 1 To UBound(avarTextBoxAndLabelConfig, 2)
        ' null value handler
        If Not IsNull(avarTextBoxAndLabelConfig(3, inti)) Then
            CallByName Auftrag, avarTextBoxAndLabelConfig(3, inti), VbLet, Forms.Item("frmSearchMain").Controls(avarTextBoxAndLabelConfig(2, inti))
        End If
    Next
    
    ' save Auftrag
    Auftrag.SaveRecordset Forms.Item("frmSearchMain").Controls(avarTextBoxAndLabelConfig(2, 0))
End Sub

Public Sub SearchAuftrag()
    
    ' set name of resulting query
    Dim strResultQueryName As String
    strResultQueryName = "qryAuftragAuswahl"
    
    ' set name of origin query
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
        .SQL = " SELECT " & strQuerySourceName & ".*" _
                & " FROM " & strQuerySourceName & "" _
                & " WHERE " & strQuerySourceName & ".AftrID LIKE '*" & varSearchTerm & "*' OR " & strQuerySourceName & ".AftrTitel LIKE '*" & varSearchTerm & "*' OR " & strQuerySourceName & ".ParentKey LIKE '*" & varSearchTerm & "*'" _
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


