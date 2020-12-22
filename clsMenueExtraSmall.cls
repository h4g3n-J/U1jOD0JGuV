VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenueExtraSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsMenueExtraSmall

Option Compare Database
Option Explicit

Private mstrModus As String
Private mbolVerbatim As Boolean

Private Sub Class_Initialize()
    Debug.Print "Klasse clsMenueExtraSmall instanziiert"
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsMenueExtraSmall terminiert"
End Sub

Private Sub Default()

End Sub

Public Sub Speichern()
 
    Select Case mstrModus
        Case "AuftragStornieren"
            AuftragStornierenSpeichern
        Case "LiefergegenstandErstellen"
            LiefergegenstandErstellenSpeichern
        Case "KostenErfassen"
            KostenErfassenSpeichern
    End Select

End Sub

Private Sub AuftragStornierenSpeichern()
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenueExtraSmall.AuftragStornierenSpeichern: " _
            & "gobjAuftrag nicht instanziiert"
        Exit Sub
    End If
    With Forms!frmMenueExtraSmall
        gobjAuftrag.storniert = !txt4
        gobjAuftrag.Bemerkung = !txt5
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub LiefergegenstandErstellenSpeichern()
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenueExtraSmall.LiefergegenstandErstellenSpeichern: " _
            & "gobjLiefergegenstand nicht instanziiert"
        Exit Sub
    End If
        
    gobjLiefergegenstand.Default
    
    With Forms!frmMenueExtraSmall
        gobjLiefergegenstand.AftrID = !txt2
        gobjLiefergegenstand.bezeichnung = !txt3
        gobjLiefergegenstand.anzahl = !txt4
        gobjLiefergegenstand.Bemerkung = !txt5
        gobjLiefergegenstand.IstReserve = !chk0
    End With
    
    gobjLiefergegenstand.DatensatzHinzufuegen
    ' gobjLiefergegenstand.Speichern
End Sub

Private Sub KostenErfassenSpeichern()
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenueExtraSmall.KostenErfassenSpeichern: " _
            & "gobjLiefergegenstand nicht instanziiert"
        Exit Sub
    End If
    
    With Forms!frmMenueExtraSmall
        gobjLiefergegenstand.angebotNetto = !txt2
        gobjLiefergegenstand.UnterauftragnehmerAngebotLink = !txt3
        gobjLiefergegenstand.preisBrutto = !txt4
        gobjLiefergegenstand.Bemerkung = !txt5
    End With
    
    gobjLiefergegenstand.Speichern
End Sub

Public Property Let Modus(ByVal strModus As Variant)
    mstrModus = PruefeString(strModus)
End Property

Public Property Get Modus() As Variant
    Modus = PruefeString(mstrModus)
End Property

' wertet des Attribut Modus aus und öffnet das MenueExtraSmall
' im gewählten Modus
Public Sub Oeffnen()
    Select Case mstrModus
        Case "AuftragStornierenOeffnen"
            AuftragStornierenOeffnen
        Case "LiefergegenstandErstellen"
            LiefergegenstandErstellenOeffnen
        Case "KostenErfassen"
            KostenErfassenOeffnen
    End Select
End Sub

Private Sub AuftragStornierenOeffnen()
    DoCmd.OpenForm "frmMenueExtraSmall", acNormal
    
    With Forms!frmMenueExtraSmall
        !lbl2.Caption = "BWI Alias"
        !lbl3.Caption = "Status"
        !lbl4.Caption = "Storniert (Datum)"
        !lbl5.Caption = "Bemerkung"
        
        If Not gobjAuftrag Is Nothing Then
            !txt0 = gobjAuftrag.ID
            !txt0.BorderStyle = 0
            !txt1 = gobjAuftrag.Titel
            !txt1.BorderStyle = 0
            !txt2 = gobjAuftrag.BWIAlias
            !txt2.BorderStyle = 0
            !txt3 = gobjAuftrag.AuftragStatus
            !txt3.BorderStyle = 0
            !txt4 = gobjAuftrag.storniert
            !txt4.Format = "short date"
            !txt5 = gobjAuftrag.Bemerkung
        End If
        
        !cmd0.Caption = "Schließen"
        !cmd0.Visible = True
        !cmd1.Caption = "Speichern"
        !cmd1.Visible = True
        !cmd2.Caption = "Leer"
        !cmd2.Visible = False
    End With
End Sub

Private Sub LiefergegenstandErstellenOeffnen()
    DoCmd.OpenForm "frmMenueExtraSmall", acNormal
    
    With Forms!frmMenueExtraSmall
        !lbl2.Caption = "Auftrag ID"
        !lbl3.Caption = "Bezeichnung"
        !lbl4.Caption = "Anzahl"
        !lbl5.Caption = "Bemerkung"
        !lbl6.Caption = "Umlaufreserve"
        
        !txt2.SetFocus
        !txt0.BorderStyle = 0
        !txt1.Visible = False
        !chk0 = False
        
        ' !cmd2.Caption = "Speichern" & vbCrLf _
            & "&" & vbCrLf _
            & "nächster Datensatz"
        !cmd2.Visible = False
    End With
End Sub

Private Sub KostenErfassenOeffnen()
    DoCmd.OpenForm "frmMenueExtraSmall", acNormal
        
    With Forms!frmMenueExtraSmall
        !lbl2.Caption = "Angebot (Netto)"
        !lbl3.Caption = "UAN Angebot (Link)"
        !lbl4.Caption = "Preis (Brutto)"
        !lbl5.Caption = "Bemerkung"
        !lbl6.Visible = False
    
        If Not gobjLiefergegenstand Is Nothing Then
            !txt0 = gobjLiefergegenstand.LiefergegenstandID
            !txt1 = gobjLiefergegenstand.bezeichnung
            !txt2 = gobjLiefergegenstand.angebotNetto
            !txt3 = gobjLiefergegenstand.UnterauftragnehmerAngebotLink
            !txt4 = gobjLiefergegenstand.preisBrutto
            !txt5 = gobjLiefergegenstand.Bemerkung
        End If
            
        !txt2.SetFocus
        !txt0.BorderStyle = 0
        !txt1.BorderStyle = 0
        !txt2.Format = "Euro"
        !txt3.IsHyperlink = True
        !txt4.Format = "Euro"
        
        !chk0.Visible = False
        
        !cmd2.Visible = False
    End With
End Sub

' Die Pruefe-Prozeduren sollen ermöglichen, dass null Werte in den
' Recordset geschrieben und von dort ausgelesen als solche ausgelesen
' werden können, gleichzeitig sollen sie verhindern, dass falsche
' Datentypen eingegeben werden können

' Prüft, ob der übergebene Wert vom Typ String ist
' und überführt ihn in diesen Typ
' Private Function PruefeString(ByVal varInput As Variant) As String
Private Function PruefeString(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
    
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeString: " _
                & "Property ist null oder leer"
        End If
        
        PruefeString = varInput
        Exit Function
    End If
    
    PruefeString = CStr(varInput)
End Function
    

