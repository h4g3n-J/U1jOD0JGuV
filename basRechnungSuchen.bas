Attribute VB_Name = "basRechnungSuchen"
Option Compare Database
Option Explicit

Public Sub BuildRechnungSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchen.BuildRechnungSuchen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungSuchen"
    
    ' clear form
     basAuftragSuchen.ClearForm strFormName
     
     ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' set form caption
    objForm.Caption = strFormName
    
    ' declare command button
    Dim btnButton As CommandButton
    
    ' declare label
    Dim lblLabel As Label
    
    ' declare textbox
    Dim txtTextbox As TextBox
    
    ' declare subform
    Dim frmSubForm As SubForm
    
    ' declare grid variables
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intWidth As Integer
        Dim intHeight As Integer
        
        Dim intColumn As Integer
        Dim intRow As Integer
        Dim strParent As String
        
    ' wip starting here ---->
    
    ' <---- wip ending here
    
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchen.BuildAuftragSuchen: " & strFormName & " erstellt"
    End If

End Sub
