Attribute VB_Name = "basLeistungserfassungsblattSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildLeistungserfassungsblattSuchenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.BuildLeistungserfassungsblattSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLeistungserfasssungsblattSuchenSub"
    
    ' clear form
    basLeistungserfassungsblattSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query
    Dim strQueryName As String
    strQueryName = "qryLeistungserfassungsblattSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "Leistungserfassungsblatt"
    
    basLeistungserfassungsblattSuchenSub.SearchLeistungserfassungsblatt strQueryName, strQuerySource, strPrimaryKey
    
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
        
            intNumberOfColumns = 8
            intNumberOfRows = 2
            intColumnWidth = 2500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "RechnungNr"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "RechnungNr"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    ' start editing here --->
    
    ' <--- stop editing here
    
        
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    ' objForm.OnCurrent = "=selectLeistungserfassungsblatt()"
        
    ' set form properties
    objForm.AllowDatasheetView = True
    objForm.AllowFormView = False
    objForm.DefaultView = 2 '2 is for datasheet
    
    ' restore form size
    DoCmd.Restore
    
    ' close and rename form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.BuildAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.ClearForm"
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
        Debug.Print "basLeistungserfassungsblattSuchenSub.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattSuchenSub"
    
    basLeistungserfassungsblattSuchenSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basLeistungserfassungsblattSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basLeistungserfassungsblattSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestClearForm executed"
    End If
    
End Sub

Public Sub SearchLeistungserfassungsblatt(ByVal strQueryName As String, ByVal strQuerySource As String, ByVal strPrimaryKey As String, Optional varSearchTerm As Variant = Null)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.SearchLeistungserfassungsblatt"
    End If
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basLeistungserfassungsblattSuchenSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT " & strQuerySource & ".*" & _
                    " FROM " & strQuerySource & _
                    " WHERE " & strQuerySource & "." & strPrimaryKey & " LIKE '*" & strSearchTerm & "*'" & _
                    " ;"
    End With
    
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
        Debug.Print "basLeistungserfassungsblattSuchenSub.SearchLeistungserfassungsblatt executed"
    End If

End Sub

Private Sub TestSearchLeistungserfassungsblatt()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestSearchLeistungserfassungsblatt"
    End If
        
    ' build query qryRechnungSuchen
    Dim strQueryName As String
    strQueryName = "qryLeistungserfassungsblattSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblLeistungserfassungsblatt"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "Leistungserfassungsblatt"
    
    basLeistungserfassungsblattSuchenSub.SearchLeistungserfassungsblatt strQueryName, strQuerySource, strPrimaryKey
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basLeistungserfassungsblattSuchenSub.TestSearchLeistungserfassungsblatt"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basLeistungserfassungsblattSuchenSub.TestSearchLeistungserfassungsblatt"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestSearchLeistungserfassungsblatt executed"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.DeleteQuery"
    End If
    
    ' set dummy object
    Dim objDummy As Object
    ' search object list >>AllQueries<< for strQueryName
    For Each objDummy In Application.CurrentData.AllQueries
        If objDummy.Name = strQueryName Then
            
            ' check if query isloaded
            If objDummy.IsLoaded Then
                ' close query
                DoCmd.Close acQuery, strQueryName, acSaveYes
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basLeistungserfassungsblattSuchenSub.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
                End If
            End If
    
            ' delete query
            DoCmd.DeleteObject acQuery, strQueryName
            
            ' exit loop
            Exit For
        End If
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.DeleteQuery executed"
    End If
    
End Sub

Private Sub TestDeleteQuery()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestDeleteQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryLeistungserfassungsblattSuchen"
    
    ' delete query
    basLeistungserfassungsblattSuchenSub.DeleteQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not deleted.", vbCritical, "basLeistungserfassungsblattSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " was not detected", vbOKOnly, "basLeistungserfassungsblattSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestDeleteQuery executed"
    End If
    
End Sub
