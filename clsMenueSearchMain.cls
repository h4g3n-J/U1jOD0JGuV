VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenueSearchMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsMenueSearchMain
' Formatiert das Formular SearchMain und ist für die Darstellung
' der Attribute im Formular zuständig

Option Compare Database
Option Explicit

Private mstrModus As String
Private mbolVerbatim As Boolean

' Default Modus des Formulars
Public Sub Default()

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

' wertet des Attribut Modus aus und öffnet das MenueSearchMain
' im gewählten Modus
Public Sub Oeffnen()
    ' Debug.Print "clsMenueSearchMain.Oeffnen ausführen, mstModus: " _
                + mstrModus
    Select Case mstrModus
        Case "AuftragSuchen"
            AuftragOeffnen
        Case "LiefergegenstandSuchen"
            LiefergegenstandOeffnen
    End Select
End Sub

Private Sub AuftragOeffnen()
    Debug.Print "clsMenueSearchMain: AuftragOeffnen ausführen"
    
    ' überprüft, ob das Formular frmSearchMain bereits geöffnet ist,
    ' wenn ja Formular schließen
    If CurrentProject.AllForms("frmSearchMain").IsLoaded Then
        Debug.Print "frmSearchMain ist bereits geöffnet, " _
            + "schließe frmSearchMain"
        DoCmd.Close acForm, "frmSearchMain", acSaveYes
    End If
    
    ' Formular frmSearchMain öffnen
    DoCmd.OpenForm "frmSearchMain", acNormal
    
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

Private Sub LiefergegenstandOeffnen()
    Default
    
    ' überprüft, ob das Formular frmSearchMain bereits geöffnet ist,
    ' wenn ja Formular schließen
    If CurrentProject.AllForms("frmSearchMain").IsLoaded Then
        DoCmd.Close acForm, "frmSearchMain", acSaveYes
    End If
    
    ' Formular frmSearchMain öffnen
    DoCmd.OpenForm "frmSearchMain", acNormal
    
    With Forms!frmSearchMain
        ' !lbl0.Caption = "ID"
        ' !lbl1.Caption = "Bezeichnung"
        ' !lbl2.Caption = "Seriennummer"
        ' !lbl3.Caption = "Herstellerkennzeichen"
        ' !lbl4.Caption = "Anzahl"
        ' !lbl5.Caption = "BWI Alias"
        ' !lbl6.Caption = "Parent"
        ' !lbl7.Caption = "Lieferdatum"
        ' !lbl8.Caption = "Lieferschein"
        ' !lbl9.Caption = "Lieferschein (Link)"

        !lbl0.Caption = "Bezeichnung"
        !lbl1.Caption = "Seriennummer"
        !lbl2.Caption = "Anzahl"
        !lbl3.Caption = "Auftrag ID"
        !lbl4.Caption = "BWI Alias"
        !lbl5.Caption = "Lieferschein (Link)"
        !lbl6.Caption = "Ziel Auftrag ID"
        !lbl7.Caption = "Ziel Lieferschein (Link)"
        !lbl8.Caption = "Liefergegenstand (Link)"
        !lbl9.Caption = "Bemerkung"
        
        !txt0.BorderStyle = 1
        !txt1.BorderStyle = 1
        !txt2.BorderStyle = 1
        !txt3.BorderStyle = 1
        !txt4.BorderStyle = 0
        !txt5.BorderStyle = 1
        !txt5.IsHyperlink = True
        !txt6.BorderStyle = 1
        !txt7.BorderStyle = 1
        !txt7.IsHyperlink = True
        !txt8.BorderStyle = 1
        !txt8.Visible = True
        !txt8.IsHyperlink = True
        !txt9.BorderStyle = 1
        
        !cmd0.Caption = "Hauptmenü"
        !cmd1.Caption = "Suchen"
        !cmd1.Visible = True
        !cmd2.Caption = "Gegenstand erstellen"
        !cmd2.Visible = True
        !cmd3.Caption = "Auslieferung erfassen"
        !cmd3.Visible = True
        !cmd4.Caption = "Kosten erfassen"
        !cmd4.Visible = True
        !cmd5.Caption = "Umlaufreserve zuteilen"
        !cmd5.Visible = True
        !cmd6.Caption = "Umlaufreserve Auslieferung erfassen"
        !cmd6.Visible = True
        !cmd7.Caption = "Datensatz hinzufügen"
        !cmd7.Visible = False
        !cmd8.Caption = "Auftrag anzeigen"
        !cmd8.Visible = False
        !cmd9.Caption = "leer"
        !cmd9.Visible = False
        !cmd10.Caption = "leer"
        !cmd10.Visible = False
        !cmd11.Caption = "Speichern"
        !cmd11.Visible = True
        
        ' Unterformular modifizieren
        !frb1.SourceObject = "frmLiefergegenstandSuchenSub"
    End With
End Sub

Public Sub DatensatzAnzeigen(ByVal varID As Variant)
    Select Case mstrModus
        Case "AuftragSuchen"
            AuftragAnzeigen (varID)
        Case "LiefergegenstandSuchen"
            LiefergegenstandAnzeigen (varID)
    End Select
End Sub

Private Sub AuftragAnzeigen(ByVal varID As Variant)
    Debug.Print "clsMenueSearchMain: AuftragAnzeigen ausführen"
    
    ' Überprüfen, ob das Formular frmSearchMain geöffnet ist
    If Not CurrentProject.AllForms("frmSearchMain").IsLoaded Then
        Debug.Print "clsMenueSearchMain.AuftragAnzeigen: " _
            & "frmSeachMain nicht geöffnet"
        Exit Sub
    End If
    
    gobjAuftrag.Laden (varID)
    
    With Forms!frmSearchMain
            !txt0 = gobjAuftrag.ID
            !txt1 = gobjAuftrag.Titel
            !txt2 = gobjAuftrag.BWIAlias
            !txt3 = gobjAuftrag.ICDStatus
            !txt4 = gobjAuftrag.AuftragStatus
            !txt5 = gobjAuftrag.Parent
            !txt6 = gobjAuftrag.LeistungsbeschreibungLink
            !txt7 = gobjAuftrag.Mengengeruest
            !txt8 = gobjAuftrag.RechnungLink
            !txt9 = gobjAuftrag.Bemerkung
    End With
End Sub

Private Sub LiefergegenstandAnzeigen(ByVal varID As Variant)
     
     ' Überprüfen, ob das Formular frmSearchMain geöffnet ist
    If Not CurrentProject.AllForms("frmSearchMain").IsLoaded Then
        Debug.Print "clsMenueSearchMain: frmSeachMain nicht geöffnet, " _
            & "AuftragAnzeigen abbrechen"
        Exit Sub
    End If
    
    gobjLiefergegenstand.Laden (varID)
    
    With Forms!frmSearchMain
        !txt0 = gobjLiefergegenstand.bezeichnung
        !txt1 = gobjLiefergegenstand.seriennummer
        !txt2 = gobjLiefergegenstand.anzahl
        !txt3 = gobjLiefergegenstand.AftrID
        !txt4 = gobjLiefergegenstand.BWIKey
        !txt5 = gobjLiefergegenstand.LieferscheinLink
        !txt6 = gobjLiefergegenstand.zielAftrID
        !txt7 = gobjLiefergegenstand.zielLinkLieferschein
        !txt8 = gobjLiefergegenstand.LiefergegenstandLink
        !txt9 = gobjLiefergegenstand.Bemerkung
    End With
End Sub

' wertet das Attribut mstrModus für die
' Funktion 'Suchen' aus
'Public Sub Suchen(ByVal varSuchen As Variant)
Public Sub Suchen(ByVal varSuchen As Variant)
    Select Case mstrModus
        Case "AuftragSuchen"
            If Not gobjAuftrag Is Nothing Then
                gobjAuftrag.Suchen (varSuchen)
            End If
        Case "LiefergegenstandSuchen"
            If Not gobjLiefergegenstand Is Nothing Then
                gobjLiefergegenstand.Suchen (varSuchen)
            End If
    End Select
End Sub

Public Sub cmd2()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' öffnet das Menü MenueSmall im Modus "AngebotErfassen"
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd2: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "AngebotErfassen"
            gobjMenueSmall.Oeffnen
        
        Case "LiefergegenstandSuchen"
            ' öffnet das Menü MenueExtraMall im Modus "LiefergegenstandErstellenOeffnen"
            If gobjMenueExtraSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd2: " _
                    & "gobjMenueExtraSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueExtraSmall.Modus = "LiefergegenstandErstellen"
            gobjMenueExtraSmall.Oeffnen
    End Select
End Sub

Public Sub cmd3()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' öffnet das Menü MenueSmall im Modus "AuftragErteilen"
            
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd3: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "AuftragErteilen"
            gobjMenueSmall.Oeffnen
        Case "LiefergegenstandSuchen"
            ' öffnet das Menü MenueSmall im Modus "AuslieferungErfassen"
            
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd3: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "AuslieferungErfassen"
            gobjMenueSmall.Oeffnen
    End Select
End Sub

Public Sub cmd4()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' öffnet das Menü MenueSmall im Modus "LeistungAbnehmen"
            
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd4: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "LeistungAbnehmen"
            gobjMenueSmall.Oeffnen
        Case "LiefergegenstandSuchen"
            ' öffnet das Menü MenueExtraSmall im Modus "KostenErfassen"
            
            If gobjMenueExtraSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd4: " _
                    & "gobjMenueExtraSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueExtraSmall.Modus = "KostenErfassen"
            gobjMenueExtraSmall.Oeffnen
    End Select
    
End Sub

Public Sub cmd5()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' öffnet das Menü MenueSmall im Modus "RechnungErfassen"
            
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd5: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "RechnungErfassen"
            gobjMenueSmall.Oeffnen
        Case "LiefergegenstandSuchen"
            ' öffnet das Menü MenueSmall im Modus "UmlaufreserveZuweisen"
            
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd5: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            
            gobjMenueSmall.Modus = "UmlaufreserveZuweisen"
            gobjMenueSmall.Oeffnen
    End Select
    
End Sub

Public Sub cmd6()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' öffnet das Menü MenueSmall im Modus "AuftragStornieren"
            If gobjMenueExtraSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd6: " _
                    & "gobjMenueExtraSmall nicht instanziiert"
                Exit Sub
            End If
            gobjMenueExtraSmall.Modus = "AuftragStornieren"
            gobjMenueExtraSmall.Oeffnen
        Case "LiefergegenstandSuchen"
            If gobjMenueSmall Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd6: " _
                    & "gobjMenueSmall nicht instanziiert"
                Exit Sub
            End If
            gobjMenueSmall.Modus = "UmlaufreserveAuslieferungErfassen"
            gobjMenueSmall.Oeffnen
    End Select
End Sub

Public Sub cmd7()
    Select Case mstrModus
        ' öffnet das Menü MenueSmall im Modus "AuftragDetails"
        Case "AuftragSuchen"
             If gobjAuftrag Is Nothing Or gobjMenueSmall Is Nothing Then
                Exit Sub
            End If
            gobjMenueSmall.Modus = "AuftragDetails"
            gobjMenueSmall.Oeffnen
        ' öffne das Menü MenueSmall im Modus "Liefergegenstand erstellen"
        Case "LiefergegenstandSuchen"
            If gobjMenueSmall Is Nothing Or gobjMenueSmall Is Nothing Then
                Exit Sub
            End If
            gobjMenueSmall.Modus = "LiefergegenstandErstellen"
            gobjMenueSmall.Oeffnen
    End Select
End Sub

Public Sub cmd8()
    Select Case mstrModus
        ' Querverweis von Auftrag nach Liefergegenstand
        ' schließt frmSearchMain, öffnet frmSearchMain im
        ' Modus LiefergegenstandSuchen, führt eine Suche
        ' nach der aktuell ausgewählten AftrID aus
        Case "AuftragSuchen"
            If gobjAuftrag Is Nothing Then
                Debug.Print "clsMenueSearchMain.cmd8: gobjAuftrag " _
                    + "is nothing, clsMenueSearchMain.cmd8 abbrechen"
                Exit Sub
            Else
                If IsNull(gobjAuftrag.ID) Then
                    MsgBox "Kein Datensatz ausgewählt", vbCritical, "Fehler"
                    Exit Sub
                Else
                    LiefergegenstandSuchen
                    gobjLiefergegenstand.Suchen (gobjAuftrag.ID)
                End If
            End If
        Case "LiefergegenstandSuchen"
            If gobjLiefergegenstand Is Nothing Then
                Exit Sub
            Else
                If IsNull(gobjLiefergegenstand) Then
                    AuftragSuchen
                Else
                    AuftragSuchen
                    gobjAuftrag.Suchen (gobjLiefergegenstand.AftrID)
                End If
            End If
    End Select
End Sub

Public Sub cmd9()
    Select Case mstrModus
        Case "AuftragSuchen"
            ' AuftragSpeichern
        Case "LiefergegenstandSuchen"
            ' LiefergegenstandSpeichern
    End Select
End Sub

Public Sub cmd10()
    Select Case mstrModus
    End Select
End Sub

Public Sub cmd11()
    Select Case mstrModus
        Case "AuftragSuchen"
            AuftragSpeichern
        Case "LiefergegenstandSuchen"
            LiefergegenstandSpeichern
    End Select
End Sub

Private Sub AuftragSpeichern()
    Debug.Print "clsMenueSearchMain.AuftragSpeichern ausführen"
    If gobjAuftrag Is Nothing Then
        Debug.Print "clsMenueSearchMain.AuftragSpeichern: " _
            + "gobjAuftrag is nothing, " _
            + "clsMenueSearchMain.AuftragSpeichern abgebrochen"
        Exit Sub
    End If

    With Forms!frmSearchMain
        gobjAuftrag.Titel = !txt1
        gobjAuftrag.BWIAlias = !txt2
        gobjAuftrag.Parent = !txt5
        gobjAuftrag.LeistungsbeschreibungLink = !txt6
        gobjAuftrag.Mengengeruest = !txt7
        gobjAuftrag.RechnungLink = !txt8
        gobjAuftrag.Bemerkung = !txt9
    End With
    
    gobjAuftrag.Speichern
End Sub

Private Sub LiefergegenstandSpeichern()
    Debug.Print "clsMenueSearchMain.LiefergegenstandSpeichern ausführen"
    If gobjLiefergegenstand Is Nothing Then
        Debug.Print "clsMenueSearchMain.LiefergegenstandSpeichern: " _
            + "gobjLiefergegenstand is nothing, clsMenueSearchMain.LiefergegenstandSpeichern " _
            + "abgebrochen"
        Exit Sub
    End If
    
    With Forms!frmSearchMain
        gobjLiefergegenstand.bezeichnung = !txt0
        gobjLiefergegenstand.seriennummer = !txt1
        gobjLiefergegenstand.anzahl = !txt2
        gobjLiefergegenstand.AftrID = !txt3
        gobjLiefergegenstand.LieferscheinLink = !txt5
        gobjLiefergegenstand.zielAftrID = !txt6
        gobjLiefergegenstand.zielLinkLieferschein = !txt7
        gobjLiefergegenstand.LiefergegenstandLink = !txt8
        gobjLiefergegenstand.Bemerkung = !txt9
    End With
    
    gobjLiefergegenstand.Speichern
End Sub

' Prüft, ob der übergegebene Wert vom Typ Boolean ist
' und überführt ihn in diesen Typ
Private Function PruefeBoolean(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
    
        If mbolVerbatim = True Then
            Debug.Print "clsLiefergegenstand.PruefeBoolean: " _
                & "Property ist null oder leer"
        End If
        
        PruefeBoolean = varInput
        Exit Function
    End If
    
    PruefeBoolean = CBool(varInput)
End Function

Public Property Let Modus(ByVal strModus As String)
    mstrModus = strModus
End Property

Public Property Get Modus() As String
    Modus = mstrModus
End Property

Private Sub Class_Initialize()
    Debug.Print "Klasse clsMenueSearchMain wurde istanziiert."
End Sub

Private Sub Class_Terminate()
    Debug.Print "Klasse clsMenueSearchMain wurde terminiert."
End Sub

Public Property Get Verbatim() As Variant
    Verbatim = PruefeBoolean(mbolVerbatim)
End Property

Public Property Let Verbatim(ByVal varVerbatim As Variant)
    mbolVerbatim = PruefeBoolean(bolVerbatim)
End Property




