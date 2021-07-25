Attribute VB_Name = "basAuftragSuchen"
Option Compare Database
Option Explicit

Public Sub BuildAuftragSuchen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.BuildAuftragSuchen"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
    ' clear form
     basAuftragSuchenSub.ClearForm strFormName
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.ClearForm"
    End If
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
            End If
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            Exit For
        End If
    Next
    
End Sub

Private Sub TestClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.TestClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
    basAuftragSuchen.ClearForm strFormName
    
    Dim objForm As Object
    Set objForm = CreateForm
    
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " created."
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " exists."
            Exit For
        End If
    Next
    
    basAuftragSuchen.ClearForm strFormName
    
    Debug.Print "basAuftragSuchen.TestClearForm: ClearForm executed."
    
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            Debug.Print "basAuftragSuchen.TestClearForm: " & strFormName & " exists."
            Exit For
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchen.TestClearForm"
    End If
    
End Sub
