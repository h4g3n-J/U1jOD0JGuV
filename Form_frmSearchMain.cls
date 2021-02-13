VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSearchMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form_frmSearchMain

Option Compare Database
Option Explicit

Private mbolVerbatim As Boolean
Private mstrMode As String

' Formular schließen
Private Sub cmd0_Click()
    
    DoCmd.Close acForm, "frmSearchMain", acSaveYes
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "frmSearchMain.cmd0_click ausgefuehrt"
    End If
End Sub

' search
Private Sub cmd1_Click()
    ' search
    basSearchMain.SearchAuftrag
    ' reload form frmSearchMain
    basSearchMain.OpenFormAuftrag
End Sub

' Button 'Auftrag erfassen'
Private Sub cmd2_Click()
    ' Debug.Print "Form_frmSearchMain: " _
        ' & "frmSearchMain.cmd2_click ausführen"
        
    ' If gobjSearchMain Is Nothing Then
        ' Debug.Print "Form_frmSearchMain.cmd2_Click: " _
                ' & "gobjSearchMain nicht instanziiert"
        ' Exit Sub
    ' End If
    
    ' gobjSearchMain.cmd2
    basMenueExtraSmall.FormOeffnenAuftragErstellen
End Sub

Private Sub cmd3_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd3_click ausführen"
    
    If gobjSearchMain Is Nothing Then
        Debug.Print "Form_frmSearchMain.cmd3_Click: " _
                & "gobjSearchMain nicht instanziiert"
        Exit Sub
    End If
    
    gobjSearchMain.cmd3
End Sub

Private Sub cmd4_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd4_click ausführen"
        
    If gobjSearchMain Is Nothing Then
        Debug.Print "Form_frmSearchMain.cmd4_Click: " _
                & "gobjSearchMain nicht instanziiert"
        Exit Sub
    End If
        
    gobjSearchMain.cmd4
End Sub

Private Sub cmd5_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd5_click ausführen"
    If gobjSearchMain Is Nothing Then
        Debug.Print "Form_frmSearchMain.cmd5_Click: " _
                & "gobjSearchMain nicht instanziiert"
        Exit Sub
    End If
        
    gobjSearchMain.cmd5
End Sub

' 1. Auftrag stornieren
Private Sub cmd6_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd6_click ausführen"
    If gobjSearchMain Is Nothing Then
        Exit Sub
    End If
    gobjSearchMain.cmd6
End Sub

' 1. Auftrag Details
' 2. Liefergegenstand erstellen
Private Sub cmd7_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd7_click ausführen"
    If gobjSearchMain Is Nothing Then
        Exit Sub
    End If
    gobjSearchMain.cmd7
End Sub

Private Sub cmd8_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd8_click ausführen"
    If gobjSearchMain Is Nothing Then
        Exit Sub
    End If
    gobjSearchMain.cmd8
End Sub

' Button 'Speichern'
Private Sub cmd9_Click()
    If gconVerbatim = True Then
        Debug.Print "Form_frmSearchMain: cmd9_click ausführen"
    End If
    
    basSearchMain.SaveAuftrag
End Sub

Private Sub cmd10_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd10_click ausführen"
    If gobjSearchMain Is Nothing Then
        Debug.Print "gobjSearchMain exisitert nicht, " _
            & "Form_frmSearchMain.cmd10_Click abgebrochen"
        Exit Sub
    End If
    
    gobjSearchMain.cmd10
End Sub

Private Sub cmd11_Click()
    Debug.Print "Form_frmSearchMain: " _
        & "frmSearchMain.cmd11_click ausführen"
    If gobjSearchMain Is Nothing Then
        Debug.Print "gobjSearchMain exisitert nicht, " _
            & "Form_frmSearchMain.cmd11_Click abgebrochen"
        Exit Sub
    End If
    
    gobjSearchMain.cmd11
End Sub

Public Property Let Mode(ByVal varModus As Variant)
    mstrMode = basSupport.CheckDataType(varModus, "string", mbolVerbatim)
    
    If mbolVerbatim = True Then
    Debug.Print "Form_frmSearchMain.Mode: mstrMode = " & mstrMode
    End If
    
End Property

Public Property Get Mode() As Variant
    Mode = basSupport.CheckDataType(mstrMode, "string", mbolVerbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolVerbatim = basSupport.CheckDataType(varVerbatim, "boolean", True)
End Property

Public Property Get Verbatim() As Variant
    Verbatim = basSupport.CheckDataType(mbolVerbatim, "boolean", True)
End Property


' Default Modus des Formulars
Public Sub default()

    ' überprüft, ob das Formular frmSearchMain bereits geöffnet ist,
    ' wenn ja Formular schließen
    If CurrentProject.AllForms("frmSearchMain").IsLoaded Then
        DoCmd.Close acForm, "frmSearchMain", acSaveYes
    End If
    
    ' Formular frmSearchMain öffnen
    DoCmd.OpenForm "frmSearchMain", acNormal
    
    With Forms!frmSearchMain
        !lbl0.Caption = "lbl0"
        !lbl1.Caption = "lbl1"
        !lbl2.Caption = "lbl2"
        !lbl3.Caption = "lbl3"
        !lbl4.Caption = "lbl4"
        !lbl5.Caption = "lbl5"
        !lbl6.Caption = "lbl6"
        !lbl7.Caption = "lbl7"
        !lbl8.Caption = "lbl8"
        !lbl9.Caption = "lbl9"
        
        !txt0.DefaultValue = "'txt0'"
        !txt1.DefaultValue = "'txt1'"
        !txt2.DefaultValue = "'txt2'"
        !txt3.DefaultValue = "'txt3'"
        !txt4.DefaultValue = "'txt4'"
        !txt5.DefaultValue = "'txt5'"
        !txt6.DefaultValue = "'txt6'"
        !txt7.DefaultValue = "'txt7'"
        !txt8.DefaultValue = "'txt8'"
        !txt9.DefaultValue = "'txt9'"
        !txt10.DefaultValue = "'txt10'"
        
        !cmd0.Caption = "cmd0"
        !cmd1.Caption = "cmd1"
        !cmd2.Caption = "cmd2"
        !cmd3.Caption = "cmd3"
        !cmd4.Caption = "cmd4"
        !cmd5.Caption = "cmd5"
        !cmd6.Caption = "cmd6"
        !cmd7.Caption = "cmd7"
        !cmd8.Caption = "cmd8"
        !cmd9.Caption = "cmd9"
        
        ' Unterformular modifizieren
        '!frb1.SourceObject
    End With
End Sub

Private Sub AuftragBearbeiten()
    
    If gconVerbatim = True Then
        Debug.Print "clsMenueSearchMain.AuftragBearbeiten ausfuehren"
    End If
    
    With Forms!frmSearchMain
        !lbl0.Caption = "ID"
        !lbl1.Caption = "Titel"
        !lbl2.Caption = "BWI Alias"
        !lbl3.Caption = "ICD Status"
        !lbl4.Caption = "Auftrag Status"
        !lbl5.Caption = "Parent"
        !lbl6.Caption = "Leistungsbeschreibung" & vbCrLf & "(Link)"
        !lbl7.Caption = "Mengengerüst" & vbCrLf & "(Link)"
        !lbl8.Caption = "Rechnung (Link)"
        !lbl9.Caption = "Bemerkung"
        
        '!txt0.DefaultValue = "'txt0'"
        !txt1.BorderStyle = 1
        !txt2.BorderStyle = 1
        !txt5.BorderStyle = 1
        !txt6.BorderStyle = 1
        !txt6.IsHyperlink = True
        !txt7.BorderStyle = 1
        !txt7.IsHyperlink = True
        !txt8.BorderStyle = 1
        !txt8.IsHyperlink = True
        !txt9.BorderStyle = 1
        '!txt10.DefaultValue = "'txt10'"
        
        !cmd0.Caption = "Hauptmenü"
        !cmd1.Caption = "Suchen"
        !cmd1.Visible = True
        !cmd2.Caption = "Angebot erfassen"
        !cmd2.Visible = True
        !cmd3.Caption = "Auftrag erteilen"
        !cmd3.Visible = True
        !cmd4.Caption = "Leistung abnehmen"
        !cmd4.Visible = True
        !cmd5.Caption = "Rechnung erfassen"
        !cmd5.Visible = True
        !cmd6.Caption = "Auftrag stornieren"
        !cmd6.Visible = True
        !cmd7.Caption = "Auftragsdetails anzeigen"
        !cmd7.Visible = True
        !cmd8.Caption = "Liefergegenstand anzeigen"
        !cmd8.Visible = True
        !cmd9.Caption = "speichern"
        !cmd9.Visible = False
        !cmd9.Caption = "leer"
        !cmd10.Visible = False
        !cmd11.Caption = "speichern"
        !cmd11.Visible = True
        
        ' Unterformular modifizieren
        !frb1.SourceObject = "frmAuftragSuchenSub"
    End With
End Sub

Private Sub Form_Current()
    
    If gconVerbatim = True Then
        Debug.Print "Form_frmSearchMain.Form_Current ausfuehren"
    End If
    
    If gconVerbatim = True Then
        ' Debug.Print "Form_frmSearchMain.Form_Current: mstrModus = " + mstrMode
    End If
                
    ' Select Case mstrMode
         ' Case "Auftrag"
            ' AuftragBearbeiten
        ' Case "LiefergegenstandSuchen"
            ' LiefergegenstandOeffnen
    ' End Select
End Sub
