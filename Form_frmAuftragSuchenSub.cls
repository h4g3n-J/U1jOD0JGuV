VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAuftragSuchenSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form_frmAuftragSuchenSub
' Unterformular zum Formular gobjSearchMain
' stellt die Abfrage qryAuftragAuswahl grafisch
' dar und ermöglicht die Auswahl eines Datensatzes
' mittels Cursor

Option Compare Database
Option Explicit

Private Sub Form_Current()
    ' Prüfen, ob Objekt gobjAuftrag instanziiert wurde, wenn nicht dann nachholen
    If basMain.gobjAuftrag Is Nothing Then
        Debug.Print "frmAuftragSuchen.Form_Current: basMain.gobjAuftrag is nothing, " _
            + "gobjAuftrag instanziieren"
        Set gobjAuftrag = New clsAuftrag
    End If
    
    ' Prüfen, ob gobjSearchMain instanziiert wurde, wenn nicht dann abbrechen
    If basMain.gobjSearchMain Is Nothing Then
        Debug.Print "frmAuftragSuchen.Form_Current: " _
            + "Objekt basMain.gobjSearchMain nicht instanziiert, " _
            + "gobjSearchMain.DatensatzAnzeigen wird nicht ausgeführt"
        Exit Sub
    End If
    
    ' führt die Prozedur AuftragAnzeigen aus, um die Attribute des ausgewählten
    ' Datensatzes im Formular SearchMain anzuzeigen
    basMain.gobjSearchMain.DatensatzAnzeigen (Me.AftrID)
End Sub
