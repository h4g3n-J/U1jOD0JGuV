Attribute VB_Name = "basSearchMain"
' basSearchMain

Option Compare Database
Option Explicit

Private avarTextBoxAndLabelConfig(5, 9) As Variant
Private avarCommandButtonConfig(2, 11) As Variant
Private avarSubFormConfig(1, 0) As Variant

Private Sub FormConfiguration()
        
    ' 0 = label name
    ' 1 = label caption
    ' 2 = textbox name
    ' 3 = textbox value
    ' 4 = textbox border style
    ' 5 = textbox ishyperlink
    avarTextBoxAndLabelConfig(0, 0) = "lbl0"
        avarTextBoxAndLabelConfig(1, 0) = "ID"
        avarTextBoxAndLabelConfig(2, 0) = "txt0"
        avarTextBoxAndLabelConfig(3, 0) = "AftrID"
        avarTextBoxAndLabelConfig(4, 0) = 0
        avarTextBoxAndLabelConfig(5, 0) = False
    avarTextBoxAndLabelConfig(0, 1) = "lbl1"
        avarTextBoxAndLabelConfig(1, 1) = "Titel"
        avarTextBoxAndLabelConfig(2, 1) = "txt1"
        avarTextBoxAndLabelConfig(3, 1) = "AftrTitel"
        avarTextBoxAndLabelConfig(4, 1) = 1
        avarTextBoxAndLabelConfig(5, 1) = False
    avarTextBoxAndLabelConfig(0, 2) = "lbl2"
        avarTextBoxAndLabelConfig(1, 2) = "BWI Alias"
        avarTextBoxAndLabelConfig(2, 2) = "txt2"
        avarTextBoxAndLabelConfig(3, 2) = Null ' "BWIAlias"
        avarTextBoxAndLabelConfig(4, 2) = 1
        avarTextBoxAndLabelConfig(5, 2) = False
    avarTextBoxAndLabelConfig(0, 3) = "lbl3"
        avarTextBoxAndLabelConfig(1, 3) = "ICD Status"
        avarTextBoxAndLabelConfig(2, 3) = "txt3"
        avarTextBoxAndLabelConfig(3, 3) = "StatusKey"
        avarTextBoxAndLabelConfig(4, 3) = 0
        avarTextBoxAndLabelConfig(5, 3) = False
    avarTextBoxAndLabelConfig(0, 4) = "lbl4"
        avarTextBoxAndLabelConfig(1, 4) = "Auftrag Status"
        avarTextBoxAndLabelConfig(2, 4) = "txt4"
        avarTextBoxAndLabelConfig(3, 4) = Null
        avarTextBoxAndLabelConfig(4, 4) = 0
        avarTextBoxAndLabelConfig(5, 4) = False
    avarTextBoxAndLabelConfig(0, 5) = "lbl5"
        avarTextBoxAndLabelConfig(1, 5) = "Parent"
        avarTextBoxAndLabelConfig(2, 5) = "txt5"
        avarTextBoxAndLabelConfig(3, 5) = "ParentKey"
        avarTextBoxAndLabelConfig(4, 5) = 1
        avarTextBoxAndLabelConfig(5, 5) = False
    avarTextBoxAndLabelConfig(0, 6) = "lbl6"
        avarTextBoxAndLabelConfig(1, 6) = "Leistungsbeschreibung" & vbCrLf & "(Link)"
        avarTextBoxAndLabelConfig(2, 6) = "txt6"
        avarTextBoxAndLabelConfig(3, 6) = Null
        avarTextBoxAndLabelConfig(4, 6) = 1
        avarTextBoxAndLabelConfig(5, 6) = True
    avarTextBoxAndLabelConfig(0, 7) = "lbl7"
        avarTextBoxAndLabelConfig(1, 7) = "Mengengerüst" & vbCrLf & "(Link)"
        avarTextBoxAndLabelConfig(2, 7) = "txt7"
        avarTextBoxAndLabelConfig(3, 7) = Null
        avarTextBoxAndLabelConfig(4, 7) = 1
        avarTextBoxAndLabelConfig(5, 7) = True
    avarTextBoxAndLabelConfig(0, 8) = "lbl8"
        avarTextBoxAndLabelConfig(1, 8) = "Rechnung (Link)"
        avarTextBoxAndLabelConfig(2, 8) = "txt8"
        avarTextBoxAndLabelConfig(3, 8) = Null
        avarTextBoxAndLabelConfig(4, 8) = 1
        avarTextBoxAndLabelConfig(5, 8) = True
    avarTextBoxAndLabelConfig(0, 9) = "lbl9"
        avarTextBoxAndLabelConfig(1, 9) = "Bemerkung"
        avarTextBoxAndLabelConfig(2, 9) = "txt9"
        avarTextBoxAndLabelConfig(3, 9) = "Bemerkung"
        avarTextBoxAndLabelConfig(4, 9) = 1
        avarTextBoxAndLabelConfig(5, 9) = False
        
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
        avarCommandButtonConfig(2, 3) = True
    avarCommandButtonConfig(0, 4) = "cmd4"
        avarCommandButtonConfig(1, 4) = "Leistung abnehmen"
        avarCommandButtonConfig(2, 4) = True
    avarCommandButtonConfig(0, 5) = "cmd5"
        avarCommandButtonConfig(1, 5) = "Rechnung erfassen"
        avarCommandButtonConfig(2, 5) = True
    avarCommandButtonConfig(0, 6) = "cmd6"
        avarCommandButtonConfig(1, 6) = "Auftrag stornieren"
        avarCommandButtonConfig(2, 6) = True
    avarCommandButtonConfig(0, 7) = "cmd7"
        avarCommandButtonConfig(1, 7) = "Auftragdetails anzeigen"
        avarCommandButtonConfig(2, 7) = True
    avarCommandButtonConfig(0, 8) = "cmd8"
        avarCommandButtonConfig(1, 8) = "Liefergegenstand anzeigen"
        avarCommandButtonConfig(2, 8) = True
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
    
    ' initialize FormConfiguratoin
    FormConfiguration
    
    ' set labels and textboxes
    Dim inti As Integer
    For inti = LBound(avarTextBoxAndLabelConfig, 2) To UBound(avarTextBoxAndLabelConfig, 2)
        ' set label caption
        Forms.Item("frmSearchMain").Controls.Item(avarTextBoxAndLabelConfig(0, inti)).Caption = avarTextBoxAndLabelConfig(1, inti)
        ' set border style
        Forms.Item("frmSearchMain").Controls.Item(avarTextBoxAndLabelConfig(2, inti)).BorderStyle = avarTextBoxAndLabelConfig(4, inti)
        ' set ishyperlink
        Forms.Item("frmSearchMain").Controls.Item(avarTextBoxAndLabelConfig(2, inti)).IsHyperlink = avarTextBoxAndLabelConfig(5, inti)
    Next
        
    ' set command buttons
    For inti = LBound(avarCommandButtonConfig, 2) To UBound(avarCommandButtonConfig, 2)
        ' set caption
        Forms.Item("frmSearchMain").Controls.Item(avarCommandButtonConfig(0, inti)).Caption = avarCommandButtonConfig(1, inti)
        ' set visibility
        Forms.Item("frmSearchMain").Controls.Item(avarCommandButtonConfig(0, inti)).Visible = avarCommandButtonConfig(2, inti)
    Next
                
    ' set subform
    For inti = LBound(avarSubFormConfig, 2) To UBound(avarSubFormConfig, 2)
        Forms.Item("frmSearchMain").Controls.Item(avarSubFormConfig(0, inti)).SourceObject = avarSubFormConfig(1, inti)
    Next
    
End Sub

Public Sub ShowRecordset(ByVal strRecordsetName As String)
    If gconVerbatim = True Then
        Debug.Print "basSearchMain.ShowRecordset: strRecordsetName = " & strRecordsetName
    End If
        
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' check if necessary -> is necessary
    Auftrag.SelectRecordset strRecordsetName
    
    Dim strFormName As String
    strFormName = "frmSearchMain"
    
    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        Debug.Print "basSearchMain.ShowRecordset: Formular " & strFormName & _
            " ist nicht geoeffnet, Prozedur abgebrochen"
        Exit Sub
    End If
    
    Dim inti As Integer
    For inti = LBound(avarTextBoxAndLabelConfig, 2) To UBound(avarTextBoxAndLabelConfig)
        ' handler in case field value is null
        If Not IsNull(avarTextBoxAndLabelConfig(3, inti)) Then
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


