VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenueLarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mstrModus As String
Private mbolVerbatim As Boolean

Public Property Let Modus(ByVal strModus As String)
    mstrModus = strModus
End Property

Public Property Get Modus() As String
    Modus = mstrModus
End Property

Private Sub Class_Initialize()
    If mbolVerbatim Is True Then
        Debug.Print "clsMenueLarge instanziiert"
    End If
End Sub

Private Sub Class_Terminate()
    If mbolVerbatim Is True Then
        Debug.Print "clsMenueLarge terminiert"
    End If
End Sub

Public Sub Default()

End Sub

Public Sub Speichern()

End Sub

Public Sub Oeffnen()

End Sub

Private Sub LiefergegenstandDetailsOeffnen()
    DoCmd.OpenForm "frmMenueLarge", acNormal
    With Forms!frmMenueSmall
        !lbl2.Caption = "Auftrag ID"
        !lbl3.Caption = "Bezeichnung"
        !lbl4.Caption = "Herstellerkennzeichen"
        !lbl5.Caption = "Anzahl"
        !lbl6.Caption = "Seriennummer"
        !lbl7.Caption = "Lieferdatum"
        !lbl8.Caption = "Lieferschein"
        !lbl9.Caption = "Lieferschein (Link)"
        !lbl10.Caption = "UAN Angebot (Link)"
        !lbl11.Caption = "Angebot (netto)"
        !lbl12.Caption = "Preis (brutto)"
        !lbl13.Caption = "Bemerkung"
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Caption = "Umlaufreserve"
        
        !txt2.SetFocus
        !txt0 = gobjLiefergegenstand.LiefergegenstandID
        !txt1.Visible = False
        !txt2 = gobjLiefergegenstand.AftrID
        !txt3 = gobjLiefergegenstand.bezeichnung
        !txt4 = gobjLiefergegenstand.herstellerkennzeichen
        !txt5 = gobjLiefergegenstand.anzahl
        !txt6 = gobjLiefergegenstand.seriennummer
        !txt7 = gobjLiefergegenstand.lieferdatum
        !txt8 = gobjLiefergegenstand.lieferschein
        !txt9 = gobjLiefergegenstand.LieferscheinLink
        !txt10 = gobjLiefergegenstand.UnterauftragnehmerAngebotLink
        !txt11 = gobjLiefergegenstand.angebotNetto
        !txt12 = gobjLiefergegenstand.preisBrutto
        !txt13 = gobjLiefergegenstand.Bemerkung
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
        
        !chk0 = gobjLiefergegenstand.IstReserve
                
        !cmd2.Caption = "Speichern und nächster Datensatz"
        !cmd2.Visible = True
    End With
End Sub


