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
    basRechnungSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query qryRechnungSuchen
    Dim strQueryName As String
    strQueryName = "qryRechnungSuchen"
    basRechnungSuchenSub.BuildQuery strQueryName
    
    ' set recordset source
    objForm.RecordSource = strQueryName
    
    ' build information grid
    Dim aintInformationGrid() As Integer
        
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intColumnWidth As Integer
        Dim intRowHeight As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intColumn As Integer
        Dim intRow As Integer
        
            intNumberOfColumns = 11
            intNumberOfRows = 2
            intColumnWidth = 2500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.BuildAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub TestBuildRechungSuchenSub()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestBuildRechnungSuchenSub"
    End If
    
    basRechnungSuchenSub.BuildRechnungSuchenSub
    
    Dim strFormName As String
    strFormName = "fmrRechnungSuchenSub"
    
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
        
    Next
    
    If bolFormExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " detected", vbOKOnly, "basRechnungSuchenSub.TestBuildRechnungSuchenSub"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not detected", vbCritical, "basRechnungSuchenSub.TestBuildRechnungSuchenSub"
    End If
    
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
        Debug.Print "basRechnungSuchenSub.ClearForm executed"
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
        Debug.Print "basRechnungSuchenSub.TestClearForm executed"
    End If
    
End Sub

Private Sub BuildQuery(ByVal strQueryName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.BuildQuery"
    End If
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' delete query
    basRechnungSuchenSub.ClearQuery strQueryName
    
    ' declare query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    ' set query Name
    qdfQuery.Name = strQueryName
    
    ' set query SQL
    qdfQuery.SQL = " SELECT tblRechnung.*" & _
                        " FROM tblRechnung" & _
                        " ;"
                        
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With
    
ExitProc:
    qdfQuery.Close
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.BuildQuery executed"
    End If
    
End Sub

Private Sub TestBuildQuery()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.TestBuildQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryRechnungAuswahl"
    
    basRechnungSuchenSub.BuildQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basRechnungSuchenSub.TestBuildQuery"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basRechnungSuchenSub.TestBuildQuery"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestBuildQuery executed"
    End If
    
End Sub

Private Sub ClearQuery(ByVal strQueryName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.clearQuery"
    End If
    
    Dim objQuery As Object
    For Each objQuery In Application.CurrentData.AllQueries
        If objQuery.Name = strQueryName Then
            
            ' check if query is loaded
            If objQuery.IsLoaded Then
                DoCmd.Close acQuery, strQueryName, acSaveYes
            End If
                
            'delete query
            DoCmd.DeleteObject acQuery, strQueryName
            Exit For
        
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.clearQuery executed"
    End If
    
End Sub

Private Sub TestClearQuery()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestClearQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryRechnungAuswahl"
    
    basRechnungSuchenSub.ClearQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not deleted.", vbCritical, "basRechnungSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " was not detected", vbOKOnly, "basRechnungSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestClearQuery executed"
    End If
    
End Sub
