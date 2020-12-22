Attribute VB_Name = "basMain"
' basMain

Option Compare Database
Option Explicit

Public gobjAuftrag1 As clsAuftrag
Public gobjAuftrag As clsAuftragNew

Public gobjSearchMain As clsMenueSearchMain
Public gobjLiefergegenstand As clsLiefergegenstand

Public gobjMenueSmall As clsMenueSmall
Public gobjMenueExtraSmall As clsMenueExtraSmall


Public Sub AuftragSuchen()
    Debug.Print "basMain.AuftragSuchen ausführen"
    
    ' Objekt gobjAuftrag instanziieren
        If gobjAuftrag Is Nothing Then
            Debug.Print "basMain.AuftragSuchen: " _
                + "gobjAuftrag is nothing, " _
                + "gobjAuftrag instanziieren"
            Set gobjAuftrag = New clsAuftragNew
        End If
        
    ' Objekt mobjSearchMain instanziieren
        If gobjSearchMain Is Nothing Then
            Debug.Print "basMain.AuftragSuchen: " _
                + "gobjSearchMain is nothing, " _
                + "gobjAuftrag instanziieren"
            Set gobjSearchMain = New clsMenueSearchMain
        End If
        
    ' SearchMain in den Modus AuftragSuchen versetzen
        gobjSearchMain.Modus = "AuftragSuchen"
    
    ' Formular frmSearchMain öffnen
        gobjSearchMain.Oeffnen
        
    ' Formular MenueSmall instanziieren
        If gobjMenueSmall Is Nothing Then
            Debug.Print "basMain.AuftragSuchen: " _
                + "gobjMenueSmall is nothing, gobjMenueSmall " _
                + "instanziieren"
            Set gobjMenueSmall = New clsMenueSmall
        End If
        
    ' Formular ExtraSmall instanziieren
        Set gobjMenueExtraSmall = New clsMenueExtraSmall
End Sub

Public Sub LiefergegenstandSuchen()
    Debug.Print "basMain.LiefergegenstandSuchen ausführen"
    
    ' Objekt gobjAuftrag instanziieren
        If gobjLiefergegenstand Is Nothing Then
            Debug.Print "basMain.LiefergegenstandSuchen: " _
                + "gobjLiefergegegenstand is nothing, " _
                + "gobjAuftrag instanziieren"
            Set gobjLiefergegenstand = New clsLiefergegenstand
        End If
    ' Objekt mobjSearchMain instanziieren
        If gobjSearchMain Is Nothing Then
            Debug.Print "basMain.LiefergegenstandSuchen: " _
                + "gobjSearchMain is nothing, " _
                + "gobjSearchMain instanziieren"
            Set gobjSearchMain = New clsMenueSearchMain
        End If
        
    ' SearchMain in den Modus AuftragSuchen versetzen
        gobjSearchMain.Modus = "LiefergegenstandSuchen"
    
    ' Formular frmSearchMain öffnen
        gobjSearchMain.Oeffnen
        
    ' Formular MenueSmall instanziieren
        If gobjMenueSmall Is Nothing Then
            Debug.Print "basMain.LiefergegenstandSuchen: " _
                + "gobjMenueSmall is nothing, " _
                + "gobjMenueSmall instanziieren"
            Set gobjMenueSmall = New clsMenueSmall
        End If
        
    ' Formular ExtraSmall instanziieren
        If gobjMenueExtraSmall Is Nothing Then
            Debug.Print "basMain.LiefergegenstandSuchen: " _
                + "gobjMenueExtraSmall is nothing, " _
                + "gobjMenueExtraSmall instanziieren"
            Set gobjMenueExtraSmall = New clsMenueExtraSmall
        End If
End Sub

Public Sub AuftrageSchliessen()
    ' Objekt mobjSearchMain freigeben
        If Not gobjSearchMain Is Nothing Then
            Debug.Print "gobjSearchMain = nothing"
            gobjSearchMain = Nothing
        End If
    ' Objekt gobjMenueSmall freigeben
        If Not gobjMenueSmall Is Nothing Then
            Debug.Print "gobjMenueSmall = nothing"
            gobjMenueSmall = Nothing
        End If
    ' Objekt gobjMenueExtraSmall freigeben
        If Not gobjMenueExtraSmall Is Nothing Then
            Debug.Print "gobjMenueExtraSmall = Nothing"
            gobjMenueExtraSmall = Nothing
        End If
End Sub
