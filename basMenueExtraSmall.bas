Attribute VB_Name = "basMenueExtraSmall"
' basMenueExtraSmall

Option Compare Database
Option Explicit

Public Sub FormOeffnenAuftragErstellen()
    DoCmd.OpenForm "frmMenueExtraSmall", acNormal
    
    ' textbox setup
    Dim avarLayoutMenueExtraSmall(5, 5) As Variant
    
    ' 0 = label name
    ' 1 = label caption
    ' 2 = label isVisible
    ' 3 = textbox name
    ' 4 = textbox caption
    ' 5 = textbox visible
    avarLayoutMenueExtraSmall(0, 0) = Null
        avarLayoutMenueExtraSmall(1, 0) = Null
        avarLayoutMenueExtraSmall(2, 0) = False
        avarLayoutMenueExtraSmall(3, 0) = "txt0"
        avarLayoutMenueExtraSmall(4, 0) = Null
        avarLayoutMenueExtraSmall(5, 0) = False
    avarLayoutMenueExtraSmall(0, 1) = Null
        avarLayoutMenueExtraSmall(1, 1) = Null
        avarLayoutMenueExtraSmall(2, 1) = False
        avarLayoutMenueExtraSmall(3, 1) = "txt1"
        avarLayoutMenueExtraSmall(4, 1) = Null
        avarLayoutMenueExtraSmall(5, 1) = False
    avarLayoutMenueExtraSmall(0, 2) = "lbl2"
        avarLayoutMenueExtraSmall(1, 2) = "ID"
        avarLayoutMenueExtraSmall(2, 2) = True
        avarLayoutMenueExtraSmall(3, 2) = "txt2"
        avarLayoutMenueExtraSmall(4, 2) = Null
        avarLayoutMenueExtraSmall(5, 2) = True
    avarLayoutMenueExtraSmall(0, 3) = "lbl3"
        avarLayoutMenueExtraSmall(1, 3) = "Titel"
        avarLayoutMenueExtraSmall(2, 3) = True
        avarLayoutMenueExtraSmall(3, 3) = "txt3"
        avarLayoutMenueExtraSmall(4, 3) = Null
        avarLayoutMenueExtraSmall(5, 3) = True
    avarLayoutMenueExtraSmall(0, 4) = "lbl4"
        avarLayoutMenueExtraSmall(1, 4) = Null
        avarLayoutMenueExtraSmall(2, 4) = True
        avarLayoutMenueExtraSmall(3, 4) = "txt4"
        avarLayoutMenueExtraSmall(4, 4) = Null
        avarLayoutMenueExtraSmall(5, 4) = False
    avarLayoutMenueExtraSmall(0, 5) = "lbl5"
        avarLayoutMenueExtraSmall(1, 5) = Null
        avarLayoutMenueExtraSmall(2, 5) = True
        avarLayoutMenueExtraSmall(3, 5) = "txt5"
        avarLayoutMenueExtraSmall(4, 5) = Null
        avarLayoutMenueExtraSmall(5, 5) = False
        
        Forms.Item("frmMenueExtraSmall").Controls("txt2").SetFocus
        
    ' checkbox setup
    Dim avarCheckbox(4, 0) As Variant
    
    ' 0 = object name
    ' 1 = object visible
    ' 2 = label name
    ' 3 = label caption
    ' 4 = label visible
    avarCheckbox(0, 0) = "chk0"
        avarCheckbox(1, 0) = False
        avarCheckbox(2, 0) = "lbl6"
        avarCheckbox(3, 0) = Null
        avarCheckbox(4, 0) = False
        
    ' conrols setup
    Dim avarControl(2, 2) As Variant
    
    ' 0 = object name
    ' 1 = object caption
    ' 2 = object visible
    avarControl(0, 0) = "cmd0"
        avarControl(1, 0) = "Schlieﬂen"
        avarControl(2, 0) = True
    avarControl(0, 1) = "cmd1"
        avarControl(1, 1) = "Erzeugen"
        avarControl(2, 1) = True
    avarControl(0, 2) = "cmd2"
        avarControl(1, 2) = Null
        avarControl(2, 2) = False
        
    Dim inti As Integer
    For inti = LBound(avarLayoutMenueExtraSmall, 2) To UBound(avarLayoutMenueExtraSmall, 2)
        
        ' set label
        If Not (IsNull(avarLayoutMenueExtraSmall(0, inti)) Or IsNull(avarLayoutMenueExtraSmall(1, inti))) Then
            ' set label caption
            Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(0, inti)).Caption = avarLayoutMenueExtraSmall(1, inti)
            ' set label visibility
            Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(0, inti)).Visible = avarLayoutMenueExtraSmall(2, inti)
        End If
        
        ' set textbox, set textbox visibility
        Forms.Item("frmMenueExtraSmall").Controls(avarLayoutMenueExtraSmall(3, inti)).Visible = avarLayoutMenueExtraSmall(5, inti)
    Next
    
    ' set checkbox, set checkbox label
    For inti = LBound(avarCheckbox, 2) To UBound(avarCheckbox, 2)
        ' set checkbox visible
        Forms.Item("frmMenueExtraSmall").Controls(avarCheckbox(0, inti)).Visible = avarCheckbox(1, inti)
        ' set label visible
        Forms.Item("frmMenueExtraSmall").Controls(avarCheckbox(2, inti)).Visible = avarCheckbox(4, inti)
    Next
    
    ' set control
    For inti = LBound(avarControl, 2) To UBound(avarControl, 2)
        ' error handler
        If Not IsNull(avarControl(1, inti)) Then
            ' set control caption
            Forms.Item("frmMenueExtraSmall").Controls(avarControl(0, inti)).Caption = avarControl(1, inti)
        End If
        ' set control visible
        Forms.Item("frmMenueExtraSmall").Controls(avarControl(0, inti)).Visible = avarControl(2, inti)
    Next
    
End Sub

Public Sub FormularSchlieﬂen()
    If gconVerbatim = True Then
        Debug.Print "basMenueExtraSmall.FormularSchlieﬂen ausfuehren"
    End If
    
    DoCmd.Close acForm, "frmMenueExtraSmall", acSaveYes
End Sub

Public Sub AuftragErstellen()
    If gconVerbatim = True Then
        Debug.Print "basMenueExtraSmall.AuftragErstellen ausfuehren"
    End If
    
    ' error state
    Dim bolError As Boolean
    bolError = False
    
    ' create class Auftrag
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' create recordset
    ' ExitProc if recordset name is empty or recordset name already exist
    bolError = Auftrag.AddRecordset(Forms.Item("frmMenueExtraSmall").txt2)
    
        If bolError Then
            GoTo ExitProc
        End If
        
    ' set property "Bemerkung"
    Auftrag.Bemerkung = Forms.Item("frmMenueExtraSmall").txt3
    ' save recordset
    Auftrag.SaveRecordset (Forms.Item("frmMenueExtraSmall").txt2)
    
ExitProc:
End Sub

