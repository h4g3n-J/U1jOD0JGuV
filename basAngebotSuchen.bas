Attribute VB_Name = "basAngebotSuchen"
Option Compare Database
Option Explicit

Public Sub BuildAngebotSuchen()
    
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngebotSuchen"
    End If
    
    ' define form name
    Dim strFormName As String
    strFormName = "frmAngebotSuchen"
    
    Dim strTempFormName As String

     ' clear form
    basSupport.ClearForm strFormName

    Dim frm As Form
    Set frm = CreateForm(Application.CurrentDb.Name, "frmSearchMain")
    ' frm.Repaint
    
    ' write temporary form name to strFormName
    strTempFormName = frm.Name
    
    ' set form caption
    frm.Caption = strFormName
    
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAngebotSuchen.BuildAngbotSuchen: " & strFormName & " erstellt"
    End If

End Sub
