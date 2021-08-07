Attribute VB_Name = "basRechnungSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildRechnungSuchenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.BuildRechnungSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmRechnungSuchenSub"
    
    ' clear form
    ' basRechnungSuchenSub.ClearForm strFormName
    
    
End Sub

Private Sub TestBuildRechungSuchenSub()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestBuildRechnungSuchenSub"
    End If
    
    basRechnungSuchenSub.BuildRechnungSuchenSub
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestBuildRechnungSuchenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.ClearForm"
    End If
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        
        If objForm.Name = strFormName Then
        
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
            End If
            
            ' delete Form
            DoCmd.DeleteObject acForm, strFormName
            Exit For
            
        End If
        
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmRechnungSuchenSub"
    
    basRechnungSuchenSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basRechnungSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basRechnungSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub executed"
    End If
    
End Sub
