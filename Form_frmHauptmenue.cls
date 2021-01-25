VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHauptmenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form_frmHauptmenue

Option Compare Database
Option Explicit

Private Sub cmd0_Click()
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "Form_frmHauptmenue: cmd0_Click() ausfuehren"
    End If
        
    ' open frmSearchMain in AuftragSuchen mode
    basAuftrag.OpenFormAuftrag
    
End Sub

Private Sub cmd1_Click()
    Debug.Print "Form_frmHauptmenue.cmd1_click ausführen"
    basMain.LiefergegenstandSuchen
End Sub

' # Alt #

Private Sub cmdLiefergegenstand_Click()
    'Debug.Print "'frmHauptmenue' schließen"
        DoCmd.Close acForm, "frmHauptmenue", acSaveYes
    'Debug.Print "'frmLiefergegenstandSuchen'"
        DoCmd.OpenForm "frmLiefergegenstandSuchen", acNormal
End Sub

Private Sub cmdEinzelauftrag_Click()
    Debug.Print "'frmHauptmenue' schließen"
        DoCmd.Close acForm, "frmHauptmenue", acSaveYes
    Debug.Print "'frmEinzelauftrag' öffnen"
        DoCmd.OpenForm "frmEinzelauftrag", acNormal
End Sub

Private Sub cmdKontinuierlicheLeistungen_Click()
    Debug.Print "'frmHauptmenue' schließen"
        DoCmd.Close acForm, "frmHauptmenue", acSaveYes
    Debug.Print "'frmKontinuierlicheLeistung' anzeigen"
        DoCmd.OpenForm "frmKontinuierlicheLeistung", acNormal
End Sub

Private Sub cmdForecast_Click()
    Form_OeffneSchreibgesch "frmVorhabenZuAuftrag", "frmHauptmenue"
End Sub

Private Sub Form_Current()
    
    If gconVerbatim = True Then
        Debug.Print "Form_frmHauptmenue.Form_Current ausfuehren."
    End If
    
    ' Me!cmd0.Caption = "Ticket Suchen"
    ' Me!cmd1.Caption = "Liefergegenstand Suchen"
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    If gconVerbatim = True Then
        Debug.Print "Form_frmHauptmenue.Form_Open ausfuehren."
    End If
    
    Me!cmd0.Caption = "Ticket Suchen"
    Me!cmd1.Caption = "Liefergegenstand Suchen"
End Sub
