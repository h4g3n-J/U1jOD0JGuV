VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenueSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsMenueSmall

Option Compare Database
Option Explicit

Private mstrModus As String
Private mbolVerbatim As Boolean

Public Sub Default()
    DoCmd.OpenForm "frmMenueSmall", acNormal
    
    With Forms!frmMenueSmall
        !lbl2.Caption = "lbl2"
        !lbl3.Caption = "lbl3"
        !lbl4.Caption = "lbl4"
        !lbl5.Caption = "lbl5"
        !lbl6.Caption = "lbl6"
        !lbl7.Caption = "lbl7"
        !lbl8.Caption = "lbl8"
        !lbl9.Caption = "lbl9"
        !lbl10.Caption = "lbl10"
        !lbl11.Caption = "lbl11"
        !lbl12.Caption = "lbl12"
        !lbl13.Caption = "lbl13"
        !lbl14.Caption = "lbl14"
        !lbl15.Caption = "lbl15"
        !lbl16.Caption = "lbl16"
        !lbl17.Caption = "lbl17"
        !lbl18.Caption = "lbl18"
        
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
        !txt11.DefaultValue = "'txt11'"
        !txt12.DefaultValue = "'txt12'"
        !txt13.DefaultValue = "'txt13'"
        !txt14.DefaultValue = "'txt14'"
        !txt15.DefaultValue = "'txt15'"
        !txt16.DefaultValue = "'txt16'"
        !txt17.DefaultValue = "'txt17'"
        
        !cmd0.Caption = "cmd0"
        !cmd0.Visible = True
        !cmd1.Caption = "cmd1"
        !cmd1.Visible = True
        !cmd2.Caption = "cmd2"
        !cmd2.Visible = True
        
        !chk0.Visible = True
    End With
End Sub


Public Sub Speichern()

    Select Case mstrModus
        Case "AngebotErfassen"
            AngebotErfassenSpeichern
        Case "AuftragErteilen"
            AuftragErteilenSpeichern
        Case "LeistungAbnehmen"
            LeistungAbnehmenSpeichern
        Case "RechnungErfassen"
            RechnungErfassenSpeichern
        Case "AuslieferungErfassen"
            AuslieferungErfassenSpeichern
        ' Case "UmlaufreserveVerteilen"
            ' UmlaufreserveVerteilenSpeichern
        Case "UmlaufreserveZuweisen"
            UmlaufreserveZuweisenSpeichern
        Case "UmlaufreserveAuslieferungErfassen"
            UmlaufreserveAuslieferungErfassenSpeichern
    End Select

End Sub

Private Sub AngebotErfassenSpeichern()
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenue.AngebotErfassenSpeichern: " _
            & "gobjAuftrag nicht instanziiert"
    End If
    With Forms!frmMenueSmall
        gobjAuftrag.BWIAlias = !txt5
        gobjAuftrag.EA = !txt6
        gobjAuftrag.AngebotBrutto = !txt7
        gobjAuftrag.AngebotDatum = !txt8
        gobjAuftrag.LeistungsbeschreibungLink = !txt9
        gobjAuftrag.Mengengeruest = !txt10
        gobjAuftrag.Bemerkung = !txt11
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub AuftragErteilenSpeichern()
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenue.AuftragErteilenSpeichern: " _
            & "gobjAuftrag nicht instanziiert"
    End If
    With Forms!frmMenueSmall
        gobjAuftrag.EA = !txt5
        gobjAuftrag.AuftragBeginn = !txt6
        gobjAuftrag.AuftragEnde = !txt7
        gobjAuftrag.AngebotBrutto = !txt8
        gobjAuftrag.beauftragt = !txt9
        gobjAuftrag.Bemerkung = !txt10
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub LeistungAbnehmenSpeichern()
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenue.LeistungAbnehmenSpeichern: " _
            & "gobjAuftrag nicht instanziiert"
    End If
    With Forms!frmMenueSmall
        gobjAuftrag.abgenommen = !txt9
        gobjAuftrag.Bemerkung = !txt10
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub RechnungErfassenSpeichern()
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenue.RechnungErfassenSpeichern: " _
            & "gobjAuftrag nicht instanziiert"
    End If
    With Forms!frmMenueSmall
        gobjAuftrag.RechnungNummer = !txt11
        gobjAuftrag.RechnungBrutto = !txt12
        gobjAuftrag.TechnischRichtig = !txt13
        gobjAuftrag.Leistungserfassungsblatt = !txt14
        gobjAuftrag.RechnungLink = !txt15
        gobjAuftrag.Bemerkung = !txt17
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub AuslieferungErfassenSpeichern()
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenue.AuslieferungErfassenSpeichern: " _
            & "gobjLiefergegenstand nicht instanziiert"
    End If
    With Forms!frmMenueSmall
        gobjLiefergegenstand.lieferschein = !txt2
        gobjLiefergegenstand.bezeichnung = !txt3
        gobjLiefergegenstand.seriennummer = !txt4
        gobjLiefergegenstand.herstellerkennzeichen = !txt5
        gobjLiefergegenstand.anzahl = !txt6
        gobjLiefergegenstand.lieferdatum = !txt7
        gobjLiefergegenstand.LiefergegenstandLink = !txt8
        gobjLiefergegenstand.LieferscheinLink = !txt9
        gobjLiefergegenstand.Bemerkung = !txt10
        gobjLiefergegenstand.IstReserve = !chk0
    End With
    gobjLiefergegenstand.Speichern
End Sub

Private Sub UmlaufreserveZuweisenSpeichern()
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenueSmall.UmlaufreserveVerteilenSpeichern: " _
            & "gobjLiefergegenstand nicht instanziiert"
        Exit Sub
    End If
    With Forms!frmMenueSmall
        gobjLiefergegenstand.zielAftrID = !txt11
        gobjLiefergegenstand.zielLinkLieferschein = !txt12
        gobjLiefergegenstand.Bemerkung = !txt13
    End With
    gobjLiefergegenstand.Speichern
End Sub

Private Sub UmlaufreserveAuslieferungErfassenSpeichern()
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenueSmall.UmlaufreserveAuslieferungErfassenSpeichern: " _
            & "gobjLiefergegenstand nicht instanziiert"
        Exit Sub
    End If
    With Forms!frmMenueSmall
        gobjLiefergegenstand.zielLieferschein = !txt10
        gobjLiefergegenstand.zielLieferdatum = !txt11
        gobjLiefergegenstand.zielLinkLieferschein = !txt12
        gobjLiefergegenstand.Bemerkung = !txt13
    End With
    gobjLiefergegenstand.Speichern
End Sub

Public Property Let Modus(ByVal strModus As String)
    mstrModus = strModus
End Property

Public Property Get Modus() As String
    Modus = mstrModus
End Property

' wertet des Attribut Modus aus und öffnet das MenueSmall
' im gewählten Modus
Public Sub Oeffnen()
    Select Case mstrModus
        Case "AngebotErfassen"
            AngebotErfassenOeffnen
        Case "AuftragErteilen"
            AuftragErteilenOeffnen
        Case "LeistungAbnehmen"
            LeistungAbnehmenOeffnen
        Case "RechnungErfassen"
            RechnungErfassenOeffnen
        Case "AuftragDetails"
            AuftragDetailsOeffnen
        Case "AuslieferungErfassen"
            AuslieferungErfassenOeffnen
        ' Case "UmlaufreserveVerteilen"
            ' UmlaufreserveVerteilenOeffnen
        Case "UmlaufreserveZuweisen"
            UmlaufreserveZuweisenOeffnen
        Case "UmlaufreserveAuslieferungErfassen"
            UmlaufreserveAuslieferungErfassenOeffnen
    End Select
End Sub

Private Sub AngebotErfassenOeffnen()
    DoCmd.OpenForm "frmMenueSmall", acNormal
    
    With Forms!frmMenueSmall
        !lbl2.Visible = False
        !lbl3.Caption = "ICD Beginn(SOLL)"
        !lbl4.Caption = "ICD Ende(SOLL)"
        !lbl5.Caption = "BWI Referenz"
        !lbl6.Caption = "Zugehöriger Einzelauftrag"
        !lbl7.Caption = "Angebot Betrag (Brutto)"
        !lbl8.Caption = "Angebot eingegangen am: "
        !lbl9.Caption = "Leistungsbeschreibung (Link)"
        !lbl10.Caption = "Mengengerüst (Link)"
        !lbl11.Caption = "Bemerkung"
        !lbl12.Visible = False
        !lbl13.Visible = False
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Visible = False
        
        If Not gobjAuftrag Is Nothing Then
            !txt0 = gobjAuftrag.ID
            !txt0.BorderStyle = 0
            !txt1 = gobjAuftrag.Titel
            !txt1.BorderStyle = 0
            !txt2.Visible = False
            !txt3 = gobjAuftrag.BeginnSoll
            !txt3.Format = "Short Date"
            !txt3.BorderStyle = 0
            !txt4 = gobjAuftrag.EndeSoll
            !txt4.Format = "short Date"
            !txt4.BorderStyle = 0
            !txt5 = gobjAuftrag.BWIAlias
            !txt5.BorderStyle = 1
            !txt6 = gobjAuftrag.EA
            !txt6.BorderStyle = 1
            !txt7 = gobjAuftrag.AngebotBrutto
            !txt7.BorderStyle = 1
            !txt7.Format = "Euro"
            !txt8 = gobjAuftrag.AngebotDatum
            !txt8.BorderStyle = 1
            !txt8.Format = "Short Date"
            !txt9 = gobjAuftrag.LeistungsbeschreibungLink
            !txt9.BorderStyle = 1
            !txt9.IsHyperlink = True
            !txt10 = gobjAuftrag.Mengengeruest
            !txt10.BorderStyle = 1
            !txt10.IsHyperlink = True
            !txt11 = gobjAuftrag.Bemerkung
            !txt11.BorderStyle = 1
        End If
        
        !txt12.Visible = False
        !txt13.Visible = False
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
        
        !chk0.Visible = False
        
        !cmd0.Caption = "Schließen"
        !cmd0.Visible = True
        !cmd1.Caption = "Speichern"
        !cmd1.Visible = True
        !cmd2.Caption = "Leer"
        !cmd2.Visible = False
    End With
End Sub

Private Sub AuftragErteilenOeffnen()
    DoCmd.OpenForm "frmMenueSmall", acNormal
    
    With Forms!frmMenueSmall
        !lbl2.Visible = False
        !lbl3.Caption = "BWI Key"
        !lbl4.Caption = "Leistungsbeschreibung"
        !lbl5.Caption = "Einzelauftrag"
        !lbl6.Caption = "Auftragsbeginn (Soll)"
        !lbl7.Caption = "Auftragsende (Soll)"
        !lbl8.Caption = "Angebot Betrag (Brutto)"
        !lbl9.Caption = "beauftragt am: "
        !lbl10.Caption = "Bemerkung"
        !lbl11.Caption = "Mitzeichnung I2.5"
        !lbl12.Caption = "Mitzeichnung I2.1"
        !lbl13.Visible = False
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Visible = False
        
        If Not gobjAuftrag Is Nothing Then
            !txt0 = gobjAuftrag.ID
            !txt0.BorderStyle = 0
            !txt1 = gobjAuftrag.Titel
            !txt1.BorderStyle = 0
            !txt2.Visible = False
            !txt3 = gobjAuftrag.BWIAlias
            !txt3.BorderStyle = 0
            !txt4 = gobjAuftrag.LeistungsbeschreibungLink
            !txt4.BorderStyle = 0
            !txt4.IsHyperlink = True
            !txt5 = gobjAuftrag.EA
            !txt6 = gobjAuftrag.AuftragBeginn
            !txt6.BorderStyle = 1
            !txt6.Format = "short Date"
            !txt7 = gobjAuftrag.AuftragEnde
            !txt7.BorderStyle = 1
            !txt7.Format = "short Date"
            !txt8 = gobjAuftrag.AngebotBrutto
            !txt8.BorderStyle = 1
            !txt8.Format = "Euro"
            !txt9 = gobjAuftrag.beauftragt
            !txt9.BorderStyle = 1
            !txt9.Format = "short Date"
            !txt10 = gobjAuftrag.Bemerkung
            !txt10.BorderStyle = 1
            !txt11.Visible = False
            !txt12.Visible = False
            !txt13.Visible = False
            !txt14.Visible = False
            !txt15.Visible = False
            !txt16.Visible = False
            !txt17.Visible = False
        End If
        
        !chk0.Visible = False
        
        !cmd2.Visible = False
    End With
End Sub

Private Sub LeistungAbnehmenOeffnen()
'Default
    DoCmd.OpenForm "frmMenueSmall", acNormal
    With Forms!frmMenueSmall
        !lbl2.Caption = "BWI Key"
        !lbl3.Caption = "Leistungsbeschreibung"
        !lbl4.Caption = "Mengengerüst"
        !lbl5.Caption = "Angebot Betrag (Brutto)"
        !lbl6.Caption = "Auftragsbeginn (Soll)"
        !lbl7.Caption = "Auftragsende (Soll)"
        !lbl8.Visible = False
        !lbl9.Caption = "Abgenommen"
        !lbl10.Caption = "Bemerkung"
        !lbl11.Visible = False
        !lbl12.Visible = False
        !lbl13.Visible = False
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Visible = False
        
        !txt0 = gobjAuftrag.ID
        !txt0.BorderStyle = 0
        !txt1 = gobjAuftrag.Titel
        !txt1.BorderStyle = 0
        !txt2 = gobjAuftrag.BWIAlias
        !txt2.BorderStyle = 0
        !txt3 = gobjAuftrag.LeistungsbeschreibungLink
        !txt3.BorderStyle = 0
        !txt3.IsHyperlink = True
        !txt4 = gobjAuftrag.Mengengeruest
        !txt4.BorderStyle = 0
        !txt4.IsHyperlink = True
        !txt5 = gobjAuftrag.AngebotBrutto
        !txt5.BorderStyle = 0
        !txt5.Format = "Euro"
        !txt6 = gobjAuftrag.AuftragBeginn
        !txt6.BorderStyle = 0
        !txt6.Format = "short Date"
        !txt7 = gobjAuftrag.AuftragEnde
        !txt7.BorderStyle = 0
        !txt7.Format = "short Date"
        !txt8.Visible = False
        !txt9 = gobjAuftrag.abgenommen
        !txt9.BorderStyle = 1
        !txt9.Format = "short Date"
        !txt10 = gobjAuftrag.Bemerkung
        !txt10.BorderStyle = 1
        !txt11.Visible = False
        !txt12.Visible = False
        !txt13.Visible = False
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
        
        !chk0.Visible = False
        
        !cmd2.Caption = "Liefergegenstand anzeigen"
    End With
End Sub

Private Sub RechnungErfassenOeffnen()
    'Default
    DoCmd.OpenForm "frmMenueSmall", acNormal
    With Forms!frmMenueSmall
        !lbl2.Caption = "Einzelauftrag"
        !lbl3.Caption = "BWI Alias"
        !lbl4.Visible = False
        !lbl5.Caption = "Angebot Betrag (Brutto)"
        !lbl6.Caption = "Mengengerüst"
        !lbl7.Caption = "Leistungsbeschreibung"
        !lbl8.Caption = "Auftragsbeginn (Soll)"
        !lbl9.Caption = "Auftragsende (Soll)"
        !lbl10.Caption = "Abgenommen (Datum)"
        !lbl11.Caption = "Rechnung Nummer"
        !lbl12.Caption = "Rechnung (Brutto)"
        !lbl13.Caption = "Technisch Richtig (Datum)"
        !lbl14.Caption = "Leistungserfassungsblatt"
        !lbl15.Caption = "Rechnung (Link)"
        !lbl16.Caption = "Controlling Status"
        !lbl17.Caption = "Bemerkung"
        !lbl18.Visible = False
        
        If Not gobjAuftrag Is Nothing Then
            !txt0 = gobjAuftrag.ID
            !txt0.BorderStyle = 0
            !txt1 = gobjAuftrag.Titel
            !txt1.BorderStyle = 0
            !txt2 = gobjAuftrag.EA
            !txt2.BorderStyle = 0
            !txt3 = gobjAuftrag.BWIAlias
            !txt3.BorderStyle = 0
            !txt4.Visible = False
            !txt5 = gobjAuftrag.AngebotBrutto
            !txt5.BorderStyle = 0
            !txt5.Format = "Euro"
            !txt6 = gobjAuftrag.Mengengeruest
            !txt6.BorderStyle = 0
            !txt6.IsHyperlink = True
            !txt7 = gobjAuftrag.LeistungsbeschreibungLink
            !txt7.BorderStyle = 0
            !txt7.IsHyperlink = True
            !txt8 = gobjAuftrag.AuftragBeginn
            !txt8.BorderStyle = 0
            !txt8.Format = "Short Date"
            !txt9 = gobjAuftrag.AuftragEnde
            !txt9.BorderStyle = 0
            !txt9.Format = "Short Date"
            !txt10 = gobjAuftrag.abgenommen
            !txt10.Format = "Short Date"
            !txt10.BorderStyle = 0
            !txt11 = gobjAuftrag.RechnungNummer
            !txt12 = gobjAuftrag.RechnungBrutto
            !txt12.Format = "Euro"
            !txt13 = gobjAuftrag.TechnischRichtig
            !txt13.Format = "Short Date"
            !txt14 = gobjAuftrag.Leistungserfassungsblatt
            !txt15 = gobjAuftrag.RechnungLink
            !txt15.IsHyperlink = True
            !txt16 = gobjAuftrag.AuftragStatus
            !txt16.BorderStyle = 0
            !txt17 = gobjAuftrag.Bemerkung
        End If
        
        !chk0.Visible = False
        
        !cmd0.Caption = "Schließen"
        !cmd1.Caption = "Speichern"
        !cmd2.Visible = False
    End With
End Sub

Private Sub AuftragDetailsOeffnen()
    'Default
    
    DoCmd.OpenForm "frmMenueSmall", acNormal
    With Forms!frmMenueSmall
        !lbl2.Visible = False
        !lbl3.Caption = "BWI Alias"
        !lbl4.Caption = "Leistungsbeschreibung"
        !lbl5.Caption = "Angebot (Brutto)"
        !lbl6.Caption = "Auftragsbeginn"
        !lbl7.Caption = "Auftragsende"
        !lbl8.Caption = "Leistungserfassungsblatt"
        !lbl9.Caption = "Rechnung Nummer"
        !lbl10.Caption = "Rechnung (Link)"
        !lbl11.Caption = "Status"
        !lbl12.Caption = "Angeboten (Datum)"
        !lbl13.Caption = "Beauftragt (Datum)"
        !lbl14.Caption = "Abgenommen (Datum)"
        !lbl15.Caption = "Storniert (Datum)"
        !lbl16.Caption = "Abgebrochen (Datum)"
        !lbl17.Caption = "Bemerkung"
        !lbl18.Visible = False
    
        If Not gobjAuftrag Is Nothing Then
            !txt0 = gobjAuftrag.ID
            !txt0.BorderStyle = 0
            !txt1 = gobjAuftrag.Titel
            !txt1.BorderStyle = 0
            !txt2.Visible = False
            !txt3 = gobjAuftrag.BWIAlias
            !txt3.BorderStyle = 0
            !txt4 = gobjAuftrag.LeistungsbeschreibungLink
            !txt4.BorderStyle = 0
            !txt4.IsHyperlink = True
            !txt5 = gobjAuftrag.AngebotBrutto
            !txt5.BorderStyle = 0
            !txt5.Format = "Euro"
            !txt6 = gobjAuftrag.AuftragBeginn
            !txt6.BorderStyle = 0
            !txt6.Format = "short Date"
            !txt7 = gobjAuftrag.AuftragEnde
            !txt7.BorderStyle = 0
            !txt7.Format = "short Date"
            !txt8 = gobjAuftrag.Leistungserfassungsblatt
            !txt8.BorderStyle = 0
            !txt9 = gobjAuftrag.RechnungNummer
            !txt9.BorderStyle = 0
            !txt10 = gobjAuftrag.RechnungLink
            !txt10.BorderStyle = 0
            !txt10.IsHyperlink = True
            !txt11 = gobjAuftrag.AuftragStatus
            !txt11.BorderStyle = 0
            !txt12 = gobjAuftrag.AngebotDatum
            !txt12.BorderStyle = 0
            !txt12.Format = "short Date"
            !txt13 = gobjAuftrag.beauftragt
            !txt13.BorderStyle = 0
            !txt13.Format = "short Date"
            !txt14 = gobjAuftrag.abgenommen
            !txt14.BorderStyle = 0
            !txt14.Format = "short Date"
            !txt15 = gobjAuftrag.storniert
            !txt15.BorderStyle = 0
            !txt15.Format = "short Date"
            !txt16 = gobjAuftrag.abgebrochen
            !txt16.BorderStyle = 0
            !txt16.Format = "short Date"
            !txt17 = gobjAuftrag.Bemerkung
            !txt17.BorderStyle = 0
            !txt17.Format = "short Date"
        End If
        
        !chk0.Visible = False
        
        !cmd0.Caption = "Speichern"
        !cmd1.Visible = False
        !cmd2.Visible = False
    End With
End Sub

Private Sub AuslieferungErfassenOeffnen()
    DoCmd.OpenForm "frmMenueSmall", acNormal
        
    With Forms!frmMenueSmall
        !lbl2.Caption = "Lieferschein Nummer"
        !lbl3.Caption = "Bezeichnung"
        !lbl4.Caption = "Seriennummer"
        !lbl5.Caption = "Herstellerkennzeichen"
        !lbl6.Caption = "Anzahl"
        !lbl7.Caption = "Lieferdatum"
        !lbl8.Caption = "Liefergegenstand (Link)" & vbCrLf _
            & "(nur bei digitalen Liefergegenständen)"
        !lbl9.Caption = "Lieferschein (Link)"
        !lbl10.Caption = "Bemerkung"
        !lbl11.Visible = False
        !lbl12.Visible = False
        !lbl13.Visible = False
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Caption = "Umlaufreserve"
        
        If gobjLiefergegenstand Is Nothing Then
            Debug.Print "clsMenuSmall.AuslieferungErfassenOeffnen: " _
                & "gobjLiefergegenstand nicht instanziiert"
            Exit Sub
        End If
        
        !txt2.SetFocus
        !txt0 = gobjLiefergegenstand.LiefergegenstandID
        !txt0.BorderStyle = 0
        !txt1 = gobjLiefergegenstand.AftrID
        !txt1.BorderStyle = 0
        !txt2 = gobjLiefergegenstand.lieferschein
        !txt3 = gobjLiefergegenstand.bezeichnung
        !txt4 = gobjLiefergegenstand.seriennummer
        !txt5 = gobjLiefergegenstand.herstellerkennzeichen
        !txt6 = gobjLiefergegenstand.anzahl
        !txt7 = gobjLiefergegenstand.lieferdatum
        !txt7.Format = "Short Date"
        !txt8 = gobjLiefergegenstand.LiefergegenstandLink
        !txt8.IsHyperlink = True
        !txt9 = gobjLiefergegenstand.LieferscheinLink
        !txt9.IsHyperlink = True
        !txt10 = gobjLiefergegenstand.Bemerkung
        !txt11.Visible = False
        !txt12.Visible = False
        !txt13.Visible = False
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
                
        !chk0 = gobjLiefergegenstand.IstReserve
    End With
End Sub

Private Sub UmlaufreserveZuweisenOeffnen()
    DoCmd.OpenForm "frmMenueSmall", acNormal
    
    With Forms!frmMenueSmall
        !lbl2.Caption = "BWI Alias"
        !lbl3.Caption = "Titel"
        !lbl4.Caption = "Leistungsbeschreibung (Link)"
        !lbl5.Caption = "Bezeichnung"
        !lbl6.Caption = "Herstellerkennzeichen"
        !lbl7.Caption = "Seriennummer"
        !lbl8.Caption = "Anzahl"
        !lbl9.Caption = "Lieferdatum"
        !lbl10.Caption = "Lieferschein (Link)"
        !lbl11.Caption = "(Ziel-)Auftrag ID"
        !lbl12.Caption = "(Ziel-)Lieferschein (Link)"
        !lbl13.Caption = "Bemerkung"
        !lbl14.Visible = False
        !lbl15.Visible = False
        !lbl16.Visible = False
        !lbl17.Visible = False
        !lbl18.Caption = "Umlaufreserve"
        
        !txt11.SetFocus
        
        !txt0.BorderStyle = 0
        !txt1.BorderStyle = 0
        !txt2.BorderStyle = 0
        !txt3.BorderStyle = 0
        !txt4.BorderStyle = 0
        !txt4.IsHyperlink = True
        !txt5.BorderStyle = 0
        !txt6.BorderStyle = 0
        !txt7.BorderStyle = 0
        !txt8.BorderStyle = 0
        !txt9.BorderStyle = 0
        !txt9.Format = "Short Date"
        !txt10.BorderStyle = 0
        !txt10.IsHyperlink = True
        !txt12.IsHyperlink = True
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
        
        !cmd2.Visible = False
        
        If gobjLiefergegenstand Is Nothing Then
            Debug.Print "clsMenuSmall.UmlaufreserveVerteilenOeffnen: " _
                & "gobjLiefergegenstand nicht instanziiert"
            Exit Sub
        End If
    
        !txt0 = gobjLiefergegenstand.LiefergegenstandID
        !txt1 = gobjLiefergegenstand.AftrID
        !txt2 = gobjLiefergegenstand.BWIKey
        !txt3 = gobjLiefergegenstand.AftrTitel
        !txt4 = gobjLiefergegenstand.LeistungsbeschreibungLink
        !txt5 = gobjLiefergegenstand.bezeichnung
        !txt6 = gobjLiefergegenstand.herstellerkennzeichen
        !txt7 = gobjLiefergegenstand.seriennummer
        !txt8 = gobjLiefergegenstand.anzahl
        !txt9 = gobjLiefergegenstand.lieferdatum
        !txt10 = gobjLiefergegenstand.LieferscheinLink
        !txt11 = gobjLiefergegenstand.zielAftrID
        !txt12 = gobjLiefergegenstand.zielLinkLieferschein
        !txt13 = gobjLiefergegenstand.Bemerkung
        
        !chk0 = gobjLiefergegenstand.IstReserve
    End With
End Sub

Private Sub UmlaufreserveAuslieferungErfassenOeffnen()
    DoCmd.OpenForm "frmMenueSmall", acNormal
    
    With Forms!frmMenueSmall
        !lbl2.Caption = "BWI Alias"
        !lbl3.Caption = "Auftrag Titel"
        !lbl4.Caption = "Bezeichnung"
        !lbl5.Caption = "Anzahl"
        !lbl6.Caption = "Seriennummer"
        !lbl7.Caption = "Lieferdatum"
        !lbl8.Caption = "Lieferschein (Link)"
        !lbl9.Caption = "(Ziel-)Auftrag"
        !lbl10.Caption = "(Ziel-)Lieferschein"
        !lbl11.Caption = "(Ziel-)Lieferdatum"
        !lbl12.Caption = "(Ziel-)Lieferschein (Link)"
        !lbl13.Caption = "Bemerkung"
        !lbl18.Caption = "Umlaufreserve"
        
        !txt10.SetFocus
        
        !txt0.BorderStyle = 0
        !txt1.BorderStyle = 0
        !txt2.BorderStyle = 0
        !txt3.BorderStyle = 0
        !txt4.BorderStyle = 0
        !txt5.BorderStyle = 0
        !txt6.BorderStyle = 0
        !txt7.BorderStyle = 0
        !txt8.BorderStyle = 0
        !txt9.BorderStyle = 0
        !txt7.Format = "Short Date"
        !txt8.IsHyperlink = True
        !txt11.Format = "Short Date"
        !txt12.IsHyperlink = True
        
        If Not gobjLiefergegenstand Is Nothing Then
            !txt0 = gobjLiefergegenstand.LiefergegenstandID
            !txt1 = gobjLiefergegenstand.AftrID
            !txt2 = gobjLiefergegenstand.BWIKey
            !txt3 = gobjLiefergegenstand.AftrTitel
            !txt4 = gobjLiefergegenstand.bezeichnung
            !txt5 = gobjLiefergegenstand.anzahl
            !txt6 = gobjLiefergegenstand.seriennummer
            !txt7 = gobjLiefergegenstand.lieferdatum
            !txt8 = gobjLiefergegenstand.LieferscheinLink
            !txt9 = gobjLiefergegenstand.zielAftrID
            !txt10 = gobjLiefergegenstand.zielLieferschein
            !txt11 = gobjLiefergegenstand.zielLieferdatum
            !txt12 = gobjLiefergegenstand.zielLinkLieferschein
            !txt13 = gobjLiefergegenstand.Bemerkung
            
            !chk0 = gobjLiefergegenstand.IstReserve
        End If
        
        !txt14.Visible = False
        !txt15.Visible = False
        !txt16.Visible = False
        !txt17.Visible = False
    End With
End Sub
