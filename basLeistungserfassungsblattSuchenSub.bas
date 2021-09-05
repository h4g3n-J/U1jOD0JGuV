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
        
            intNumberOfColumns = 5
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
            .ControlSource = "Leistungserfassungsblatt"
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
            .Caption = "Leistungserfassungsblatt"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "RechnungNr"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "RechnungNr"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "Bemerkung"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "Bemerkung"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "BelegID"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "BelegID"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "Brutto"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "Brutto"
            .Left = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
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

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.CalculateGrid"
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
        Debug.Print "basLeistungserfassungsblattSuchenSub.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestCalculateGrid"
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
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "basLeistungserfassungsblattSuchenSub.TestCalculateGrid: Procedure successful", vbOKOnly, "Test Result"
        Case 1
            MsgBox "Failure: horizontal value is wrong", vbCritical, "basLeistungserfassungsblattSuchenSub.TestCalculateGrid"
        Case 2
            MsgBox "Failure: vertical value is wrong", vbCritical, "basLeistungserfassungsblattSuchenSub.TestCalculateGrid"
        Case 3
            MsgBox "Failure: horizontal and vertical values are wrong", vbCritical, "basLeistungserfassungsblattSuchenSub.TestCalculateGrid"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basLeistungserfassungsblattSuchenSub.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetLeft: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.GetTop: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basLeistungserfassungsblattSuchenSub.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetTop: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetTop executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetWidth: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basLeistungserfassungsblattSuchenSub.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetWidth: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetWidth executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetHeight: column 0 is not available"
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basLeistungserfassungsblattSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basLeistungserfassungsblattSuchenSub.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetHeight: Test failed. Error Code: 1", vbCritical
    Else
        MsgBox "basLeistungserfassungsblattSuchenSub.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestGetHeight executed"
    End If
    
End Sub

Public Function selectLeistungserfassungsblatt()
    ' Error Code 1: Form does not exist
    ' Error Code 2: Parent Form is not loaded

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.selectLeistungserfassungsblatt"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmLeistungserfassungsblattSuchen"
    
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
        Debug.Print "basLeistungserfassungsblattSuchenSub.selectLeistungserfassungsblatt aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmLeistungserfassungsblattSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.selectLeistungserfassungsblatt aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare primary key
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "Leistungserfassungsblatt"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Auftrag
    Dim Leistungserfassungsblatt As clsLeistungserfassungsblatt
    Set Leistungserfassungsblatt = New clsLeistungserfassungsblatt
    
    ' select recordset
    Leistungserfassungsblatt.SelectRecordset varRecordsetName
    
    ' show recordset
    ' referes to the textboxes in basLeistungserfassungsblattSuchen
    ' Forms.Item(strFormName).Controls.Item("insert_textboxName_here") = CallByName(insert_Object_Name, "insert_Attribute_Name_here", VbGet)
    ' start editing here -->
    ' Forms.Item(strFormName).Controls.Item("txt00") = CallByName(Leistungserfassungsblatt, "RechnungNr", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt01") = CallByName(Leistungserfassungsblatt, "Bemerkung", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt02") = CallByName(Leistungserfassungsblatt, "RechnungLink", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt03") = CallByName(Leistungserfassungsblatt, "TechnischRichtigDatum", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt04") = CallByName(Leistungserfassungsblatt, "IstTeilrechnung", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt05") = CallByName(Leistungserfassungsblatt, "IstSchlussrechnung", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt06") = CallByName(Leistungserfassungsblatt, "KalkulationLNWLink", VbGet)
    ' Forms.Item(strFormName).Controls.Item("txt07") = CallByName(Leistungserfassungsblatt, "RechnungBrutto", VbGet)
    ' <-- stop editing here
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.selectLeistungserfassungsblatt executed"
    End If
    
End Function
