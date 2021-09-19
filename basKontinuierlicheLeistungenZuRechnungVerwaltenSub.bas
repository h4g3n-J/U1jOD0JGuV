Attribute VB_Name = "basKontinuierlicheLeistungenZuRechnungVerwaltenSub"
Option Compare Database
Option Explicit

Public Sub BuildKontinuierlicheLeistungenZuRechnungVerwaltenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.BuildKontinuierlicheLeistungenZuRechnungVerwaltenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmKontinuierlicheLeistungenZuRechnungVerwaltenSub"
    
    ' clear form
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenZuRechnungVerwalten"
    
    Dim strQuerySource As String
    strQuerySource = "tblKontinuierlicheLeistungenZuRechnungVerwalten"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "ID"
    
    ' build query qryKontinuierlicheLeistungenZuRechnungVerwalten
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.SearchKontinuierlicheLeistungenZuRechnung
    
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
        
            intNumberOfColumns = 4
            intNumberOfRows = 2
            intColumnWidth = 1500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "ID"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "ID"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "RefKLID"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "KL ID"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "RefRechnungNr"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "RefRechnungNr"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "Bemerkung"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "Bemerkung"
            .Left = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    objForm.OnCurrent = "=selectKontinuierlicheLeistungenZuRechnung()"
        
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
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.BuildKontinuierlicheLeistungenZuRechnungVerwaltenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.ClearForm"
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
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmKontinuierlicheLeistungenZuRechnungVerwaltenSub"
    
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestClearForm executed"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.DeleteQuery"
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
                    Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
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
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.DeleteQuery executed"
    End If
    
End Sub

Private Sub TestDeleteQuery()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestDeleteQuery"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenZuRechnungVerwalten"
    
    ' delete query
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.DeleteQuery strQueryName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not deleted.", vbCritical, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " was not detected", vbOKOnly, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestDeleteQuery executed"
    End If
    
End Sub

Public Sub SearchKontinuierlicheLeistungenZuRechnung(Optional varSearchTerm As Variant = Null)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.SearchKontinuierlicheLeistungenZuRechnung"
    End If
    
    Dim strDomainName As String
    strDomainName = "tblKontinuierlicheLeistungenZuRechnung"
    
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenZuRechnungVerwalten"
    
    Dim strSearchField01 As String
    strSearchField01 = "RefKLID"
    
    Dim strSearchField02 As String
    strSearchField02 = "RefRechnungNr"
    
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
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
        
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT " & strDomainName & ".*" & _
                    " FROM " & strDomainName & _
                    " WHERE (" & strDomainName & "." & strSearchField01 & " Like '*" & varSearchTerm & "*') OR (" & strDomainName & "." & strSearchField02 & " Like '*" & varSearchTerm & "*')" & _
                ";"
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
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.SearchKontinuierlicheLeistungenZuRechnung executed"
    End If

End Sub

Private Sub TestSearchKontinuierlicheLeistungenZuRechnung()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestSearchKontinuierlicheLeistungenZuRechnung"
    End If
    
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenZuRechnungVerwalten"
    
    ' build query qryKontinuierlicheLeistungenZuRechnungVerwalten
    basKontinuierlicheLeistungenZuRechnungVerwaltenSub.SearchKontinuierlicheLeistungenZuRechnung
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestSearchKontinuierlicheLeistungenZuRechnung"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestSearchKontinuierlicheLeistungenZuRechnung"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestSearchKontinuierlicheLeistungenZuRechnung executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid"
    End If
    
    Const cintHorizontalSpacing As Integer = 60
    Const cintVerticalSpacing As Integer = 60
    
    Dim aintGrid() As Integer
    ReDim aintGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    Dim intColumn As Integer
    Dim intRow As Integer
    
    For intColumn = 0 To intNumberOfColumns - 1
        For intRow = 0 To intNumberOfRows - 1
            ' left
            aintGrid(intColumn, intRow, 0) = intLeft + intColumn * (intColumnWidth + cintHorizontalSpacing)
            ' top
            aintGrid(intColumn, intRow, 1) = intTop + intRow * (intRowHeight + cintVerticalSpacing)
            ' width
            aintGrid(intColumn, intRow, 2) = intColumnWidth
            ' height
            aintGrid(intColumn, intRow, 3) = intRowHeight
        Next
    Next
    
    CalculateGrid = aintGrid
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
' Error Code 1: returned horizontal value does not match the expected value
' Error Code 2: returned vertical value does not match the expected value
' Error Code 3: returned horizontal and vertical values do not match the expected values
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid"
    End If
    
    Dim intNumberOfRows As Integer
    Dim intNumberOfColumns As Integer
    Dim intRowHeight As Integer
    Dim intColumnWidth As Integer
    Dim intLeft As Integer
    Dim intTop As Integer
    
    intLeft = 50
    intTop = 50
    intNumberOfColumns = 2
    intNumberOfRows = 3
    intRowHeight = 100
    intColumnWidth = 50
    
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

    Const cintHorizontalSpacing As Integer = 60
    Const cintVerticalSpacing As Integer = 60
    
    Dim intErrorState As Integer
    intErrorState = 0
    
    Dim intBottom As Integer
    Dim intRight As Integer
    
    intBottom = intTop + (intNumberOfRows - 1) * (intRowHeight + cintVerticalSpacing)
    intRight = intLeft + (intNumberOfColumns - 1) * (intColumnWidth + cintHorizontalSpacing)
    
    If intRight <> aintInformationGrid(intNumberOfColumns - 1, 0, 0) Then
        intErrorState = intErrorState + 1
    End If
    
    If intBottom <> aintInformationGrid(0, intNumberOfRows - 1, 1) Then
        intErrorState = intErrorState + 2
    End If
    
    Select Case intErrorState
        Case 0
            MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid: Test passed", vbOKOnly, "Test Result"
        Case 1
            MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid: Test failed, Error Code 1", vbCritical, "Test Result"
        Case 2
            MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid: Test failed, Error Code 2", vbCritical, "Test Result"
        Case 3
            MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid: Test failed: Error Code 3", vbCritical, "Test Result"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Const cintHorizontalSpacing As Integer = 60
    Dim intLeftExpected As Integer
    intLeftExpected = cintLeft + (cintTestColumn - 1) * (cintHorizontalSpacing + cintColumnWidth)
    
    ' test run
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    Dim intLeftResult As Integer
    intLeftResult = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetLeft: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Const cintVerticalSpacing As Integer = 60
    Dim intTopExpected As Integer
    intTopExpected = cintTop + (cintTestRow - 1) * (cintVerticalSpacing + cintRowHeight)
    
    ' test run
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    Dim intTopResult As Integer
    intTopResult = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basKontinuierlicheLeistungenZuRechnungVerwaltenSub.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.TestGetWidth executed"
    End If
    
End Sub

Public Function selectKontinuierlicheLeistungenZuRechnung()
    ' Error Code 1: Form does not exist
    ' Error Code 2: Parent Form is not loaded

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenZuRechnungVerwaltenSub.selectKontinuierlicheLeistungenZuRechnung"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmKontinuierlicheLeistungenZuRechnung"
    
    ' check if frmAuftragSuchen exists (Error Code: 1)
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If Not bolFormExists Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.selectKontinuierlicheLeistungenZuRechnung aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmKontinuierlicheLeistungenZuRechnung not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.selectKontinuierlicheLeistungenZuRechnung aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare primary key
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "ID"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Auftrag
    Dim EinzelauftragZuRechnung As clsEinzelauftragZuRechnung
    Set EinzelauftragZuRechnung = New clsEinzelauftragZuRechnung
    
    ' select recordset
    EinzelauftragZuRechnung.SelectRecordset varRecordsetName
    
    ' show recordset
    ' referes to the textboxes in basEinzelauftragZuRechnungVerwalten
    ' Forms.Item(strFormName).Controls.Item("insert_textboxName_here") = CallByName(insert_Object_Name, "insert_Attribute_Name_here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(EinzelauftragZuRechnung, "ID", VbGet)
    Forms.Item(strFormName).Controls.Item("cbo01") = CallByName(EinzelauftragZuRechnung, "RefKLID", VbGet)
    Forms.Item(strFormName).Controls.Item("cbo02") = CallByName(EinzelauftragZuRechnung, "RefRechnungNr", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(EinzelauftragZuRechnung, "Bemerkung", VbGet)
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenZuRechnungVerwaltenSub.selectKontinuierlicheLeistungenZuRechnung executed"
    End If
    
End Function

