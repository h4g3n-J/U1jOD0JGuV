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
        
            intNumberOfColumns = 8
            intNumberOfRows = 2
            intColumnWidth = 2500
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
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
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "Bemerkung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Bemerkung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "RechnungLink"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "RechnungLink"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "TechnischRichtigDatum"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "TechnischRichtigDatum"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "IstTeilrechnung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "IstTeilrechnung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "IstSchlussrechnung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl05
    intColumn = 6
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "IstSchlussrechnung"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 7
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .ControlSource = "KalkulationLNWLink"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl06
    intColumn = 7
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "KalkulationLNWLink"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 8
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .ControlSource = "RechnungBrutto"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl07
    intColumn = 8
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "RechnungBrutto"
            .Left = basRechnungSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basRechnungSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basRechnungSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basRechnungSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    objForm.OnCurrent = "=selectRechnung()"
        
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
        Debug.Print "basRechnungSuchenSub.BuildAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub TestBuildRechungSuchenSub()
    ' Error Code 1: Form was not detected
    ' Error Code 2: Detected controls do not match expected textboxes
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestBuildRechnungSuchenSub"
    End If
    
    ' execute Procedure
    basRechnungSuchenSub.BuildRechnungSuchenSub
    
    Dim strFormName As String
    strFormName = "frmRechnungSuchenSub"
    
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    Dim strErrorMessage As String
    
    ' test if form exists
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If bolFormExists = False Then
        bolErrorState = True
        strErrorMessage = "Error code: 1"
        GoTo ExitProc
    Else
        bolErrorState = False
    End If
    
    ' count controls
    Const cintNumberOfControlsExpected As Integer = 16
    
        ' open form in datasheet view (acFormDS)
        DoCmd.OpenForm strFormName, acFormDS, , , acFormReadOnly, acWindowNormal
        
    Dim intNumberOfControls As Integer
    intNumberOfControls = Forms.Item(strFormName).Controls.Count
    
    If intNumberOfControls <> cintNumberOfControlsExpected Then
        bolErrorState = True
        strErrorMessage = strErrorMessage & vbCrLf & "Error code: 2"
    End If
    
    ' close form
    DoCmd.Close acForm, strFormName, acSaveNo
    
ExitProc:
    ' test result
    If bolErrorState Then
        MsgBox "basRechnungSuchenSub.BuildRechnungSuchenSub: Test failed." & vbCrLf & strErrorMessage, vbCritical, "Test Result"
    Else
        MsgBox "basRechnungSuchenSub.BuildRechnungSuchenSub: Test passed.", vbOKOnly, "Test Result"
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
        Debug.Print "execute basRechnungSuchenSub.TestBuildQuery"
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

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.CalculateGrid"
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
        Debug.Print "basRechnungSuchenSub.CalculateGrid executed"
    End If
    
End Function

Private Sub TestCalculateGrid()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestCalculateGrid"
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
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)

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
            MsgBox "Procedure successful", vbOKOnly, "basRechnungSuchenSub.TestCalculateGrid"
        Case 1
            MsgBox "Failure: horizontal value is wrong", vbCritical, "basRechnungSuchenSub.TestCalculateGrid"
        Case 2
            MsgBox "Failure: vertical value is wrong", vbCritical, "basRechnungSuchenSub.TestCalculateGrid"
        Case 3
            MsgBox "Failure: horizontal and vertical values are wrong", vbCritical, "basRechnungSuchenSub.TestCalculateGrid"
    End Select
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestCalculateGrid executed"
    End If
    
End Sub

Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.GetLeft"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basRechnungSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetLeft"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intLeftResult = basRechnungSuchenSub.GetLeft(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intLeftResult <> intLeftExpected Then
        MsgBox "basRechnungSuchenSub.TestGetLeft: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungSuchenSub.TestGetLeft: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetLeft executed"
    End If
    
End Sub

Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.GetTop"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungSuchenSub.GetTop: column 0 is not available"
        MsgBox "basRechnungSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetTop"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
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
    intTopResult = basRechnungSuchenSub.GetTop(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intTopResult <> intTopExpected Then
        MsgBox "basRechnungSuchenSub.TestGetTop: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungSuchenSub.TestGetTop: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetTop executed"
    End If
    
End Sub

Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetHeight"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungSuchenSub.TestGetHeight: column 0 is not available"
        MsgBox "basRechnungSuchenSub.TestGetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetHeight"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intHeightExpected As Integer
    intHeightExpected = cintRowHeight
    
    ' test run
    Dim intHeightResult As Integer
    intHeightResult = basRechnungSuchenSub.GetHeight(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intHeightResult <> intHeightExpected Then
        MsgBox "basRechnungSuchenSub.TestGetHeight: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungSuchenSub.TestGetHeight: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetHeight executed"
    End If
    
End Sub

Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetWidth"
    End If

    If intColumn = 0 Then
        Debug.Print "basRechnungSuchenSub.TestGetWidth: column 0 is not available"
        MsgBox "basRechnungSuchenSub.TestGetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()
    ' Error code1: returned value mismatches expected velue

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.TestGetWidth"
    End If
    
    Const cintNumberOfColumns As Integer = 3
    Const cintNumberOfRows As Integer = 2
    Const cintRowHeight As Integer = 100
    Const cintColumnWidth As Integer = 50
    Const cintLeft As Integer = 50
    Const cintTop As Integer = 50
        
    Dim aintInformationGrid() As Integer
    ReDim aintInformationGrid(cintNumberOfColumns - 1, cintNumberOfRows - 1, 3)
    
    aintInformationGrid = basRechnungSuchenSub.CalculateGrid(cintNumberOfColumns, cintNumberOfRows, cintLeft, cintTop, cintColumnWidth, cintRowHeight)
    
    ' set test parameters
    Const cintTestColumn As Integer = 2
    Const cintTestRow As Integer = 2
    
    ' set anticipated result
    Dim intWidthExpected As Integer
    intWidthExpected = cintColumnWidth
    
    ' test run
    Dim intWidthResult As Integer
    intWidthResult = basRechnungSuchenSub.GetWidth(aintInformationGrid, cintTestColumn, cintTestRow)
    
    If intWidthResult <> intWidthExpected Then
        MsgBox "basRechnungSuchenSub.TestGetWidth: Test missed. Error Code: 1", vbCritical
    Else
        MsgBox "basRechnungSuchenSub.TestGetWidth: Test passed.", vbOKOnly, "Test Result"
    End If

    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.TestGetWidth executed"
    End If
    
End Sub

Public Function selectRechnung()
    ' Error Code 1: Form does not exist
    ' Error Code 2: Parent Form is not loaded

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basRechnungSuchenSub.selectRechnung"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmRechnungSuchen"
    
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
        Debug.Print "basRechnungSuchenSub.selectRechnung aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmRechnungSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basRechnungSuchenSub.selectRechnung aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "RechnungNr"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Auftrag
    Dim Rechnung As clsRechnung
    Set Rechnung = New clsRechnung
    
    ' select recordset
    Rechnung.SelectRecordset varRecordsetName
    
    ' show recordset
    ' referes to the textboxes in basRechnungSuchen
    ' Forms.Item(strFormName).Controls.Item("insert_textboxName_here") = CallByName(insert_Object_Name, "insert_Attribute_Name_here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(Rechnung, "RechnungNr", VbGet)
    Forms.Item(strFormName).Controls.Item("txt01") = CallByName(Rechnung, "Bemerkung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt02") = CallByName(Rechnung, "RechnungLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(Rechnung, "TechnischRichtigDatum", VbGet)
    Forms.Item(strFormName).Controls.Item("txt04") = CallByName(Rechnung, "IstTeilrechnung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt05") = CallByName(Rechnung, "IstSchlussrechnung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt06") = CallByName(Rechnung, "KalkulationLNWLink", VbGet)
    Forms.Item(strFormName).Controls.Item("txt07") = CallByName(Rechnung, "RechnungBrutto", VbGet)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basRechnungSuchenSub.selectRechnung executed"
    End If
    
End Function
