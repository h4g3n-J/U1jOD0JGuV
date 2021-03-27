Attribute VB_Name = "basAuftragSuchenSub"
Option Compare Database
Option Explicit

' check if target form is loaded
' assign values to textboxes via avarTextBoxAndLabelConfig
Public Sub SelectRecordsetAuftrag()
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basAuftragSuchenSub.SelectRecordsetAuftrag ausfuehren"
    End If
    
    ' declare displaying form name
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    ' error handler, case target form is not loaded
    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        Debug.Print "basAuftrag.SelectRecordsetAuftrag " & strFormName & _
            " nicht geoeffnet, Prozedur abgebrochen"
        Exit Sub
    End If
    
    ' declare recordsetName, get recordset Name
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls("frb1").Controls("AftrID")
        
    ' initiate class Auftrag
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' load selected Recordset
    Auftrag.SelectRecordset varRecordsetName
    
    ' load textbox settings
    Dim varTextBoxesAndLabels As Variant
    varTextBoxesAndLabels = basAuftragSuchenSub.TextboxAndLabelSettings
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' assign recordset to textboxes
    ' skip column names
    Dim inti As Integer
    For inti = LBound(varTextBoxesAndLabels, 1) + 1 To UBound(varTextBoxesAndLabels, 1)
        ' handler in case field value is null
        ' IsEmpty is neccesary because opening frmSearchMain will open frmAuftragSuchenSub,
        ' at a time when varTextboxesAndLabels is not set yet
        If Not IsEmpty(varTextBoxesAndLabels(inti, 3)) And Not IsNull(varTextBoxesAndLabels(inti, 3)) Then
            Forms.Item(strFormName).Controls.Item(varTextBoxesAndLabels(inti, 2)) = CallByName(Auftrag, varTextBoxesAndLabels(inti, 3), VbGet)
            
            ' local verbatim message
            If bolLocalVerbatim Then
                Debug.Print "basAuftrag.SelectRecordsetAuftrag: " & varTextBoxesAndLabels(inti, 3) & "= " & CallByName(Auftrag, varTextBoxesAndLabels(inti, 3), VbGet)
            End If
        End If
    Next

End Sub

' adjust Auftrag textboxes and labels
Private Function TextboxAndLabelSettings() As Variant
    Dim varSetup(11, 9) As Variant
    varSetup(0, 0) = "label name"
        varSetup(0, 1) = "label caption"
        varSetup(0, 2) = "textbox name"
        varSetup(0, 3) = "textbox value"
        varSetup(0, 4) = "textbox border style"
        varSetup(0, 5) = "textbox ishyperlink"
        varSetup(0, 6) = "textbox locked"
        varSetup(0, 7) = "textbox format"
        varSetup(0, 8) = "textbox visible"
        varSetup(0, 9) = "textbox defaultValue"
    varSetup(1, 0) = "lbl0"
        varSetup(1, 1) = "ID"
        varSetup(1, 2) = "txt0"
        varSetup(1, 3) = "AftrID"
        varSetup(1, 4) = 0
        varSetup(1, 5) = False
        varSetup(1, 6) = True
        varSetup(1, 7) = ""
        varSetup(1, 8) = True
        varSetup(1, 9) = Null
    varSetup(2, 0) = "lbl1"
        varSetup(2, 1) = "Titel"
        varSetup(2, 2) = "txt1"
        varSetup(2, 3) = "AftrTitel"
        varSetup(2, 4) = 1
        varSetup(2, 5) = False
        varSetup(2, 6) = False
        varSetup(2, 7) = ""
        varSetup(2, 8) = True
        varSetup(2, 9) = Null
    varSetup(3, 0) = "lbl2"
        varSetup(3, 1) = "ICD Status"
        varSetup(3, 2) = "txt2"
        varSetup(3, 3) = "StatusKey"
        varSetup(3, 4) = 0
        varSetup(3, 5) = False
        varSetup(3, 6) = True
        varSetup(3, 7) = ""
        varSetup(3, 8) = True
        varSetup(3, 9) = Null
    varSetup(4, 0) = "lbl3"
        varSetup(4, 1) = "Owner"
        varSetup(4, 2) = "txt3"
        varSetup(4, 3) = "OwnerKey"
        varSetup(4, 4) = 0
        varSetup(4, 5) = False
        varSetup(4, 6) = True
        varSetup(4, 7) = ""
        varSetup(4, 8) = True
        varSetup(4, 9) = Null
    varSetup(5, 0) = "lbl4"
        varSetup(5, 1) = "Priorität"
        varSetup(5, 2) = "txt4"
        varSetup(5, 3) = "PrioritaetKey"
        varSetup(5, 4) = 0
        varSetup(5, 5) = False
        varSetup(5, 6) = True
        varSetup(5, 7) = ""
        varSetup(5, 8) = True
        varSetup(5, 9) = Null
    varSetup(6, 0) = "lbl5"
        varSetup(6, 1) = "Parent"
        varSetup(6, 2) = "txt5"
        varSetup(6, 3) = "ParentKey"
        varSetup(6, 4) = 0
        varSetup(6, 5) = False
        varSetup(6, 6) = True
        varSetup(6, 7) = ""
        varSetup(6, 8) = True
        varSetup(6, 9) = Null
    varSetup(7, 0) = "lbl6"
        varSetup(7, 1) = "Bemerkung"
        varSetup(7, 2) = "txt6"
        varSetup(7, 3) = "Bemerkung"
        varSetup(7, 4) = 1
        varSetup(7, 5) = False
        varSetup(7, 6) = False
        varSetup(7, 7) = ""
        varSetup(7, 8) = True
        varSetup(7, 9) = Null
    varSetup(8, 0) = "lbl7"
        varSetup(8, 1) = "Beginn (Soll)"
        varSetup(8, 2) = "txt7"
        varSetup(8, 3) = "BeginnSoll"
        varSetup(8, 4) = 0
        varSetup(8, 5) = False
        varSetup(8, 6) = True
        varSetup(8, 7) = "Short Date"
        varSetup(8, 8) = True
        varSetup(8, 9) = Null
    varSetup(9, 0) = "lbl8"
        varSetup(9, 1) = "Ende (Soll)"
        varSetup(9, 2) = "txt8"
        varSetup(9, 3) = "EndeSoll"
        varSetup(9, 4) = 0
        varSetup(9, 5) = False
        varSetup(9, 6) = True
        varSetup(9, 7) = "Short Date"
        varSetup(9, 8) = True
        varSetup(9, 9) = Null
    varSetup(10, 0) = "lbl9"
        varSetup(10, 1) = "Kunde"
        varSetup(10, 2) = "txt9"
        varSetup(10, 3) = "Kunde"
        varSetup(10, 4) = 1
        varSetup(10, 5) = False
        varSetup(10, 6) = False
        varSetup(10, 7) = ""
        varSetup(10, 8) = True
        varSetup(10, 9) = Null
    varSetup(11, 0) = Null
        varSetup(11, 1) = Null
        varSetup(11, 2) = "txt10"
        varSetup(11, 3) = Null
        varSetup(11, 4) = 1
        varSetup(11, 5) = False
        varSetup(11, 6) = False
        varSetup(11, 7) = ""
        varSetup(11, 8) = True
        varSetup(11, 9) = ""
        
    If gconVerbatim = True Then
        Debug.Print "basAuftragSuchenSub.TextboxAndLabelSettings ausgefuehrt"
    End If
    
    TextboxAndLabelSettings = varSetup
        
End Function

