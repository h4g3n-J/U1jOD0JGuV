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
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "Bemerkung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl01"
            .Caption = "Bemerkung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "RechnungLink"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl02"
            .Caption = "RechnungLink"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "TechnischRichtigDatum"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl03"
            .Caption = "TechnischRichtigDatum"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "IstTeilrechnung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl04"
            .Caption = "IstTeilrechnung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "IstSchlussrechnung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl05
    intColumn = 6
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl05"
            .Caption = "IstSchlussrechnung"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt06
    intColumn = 7
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt06"
            .ControlSource = "KalkulationLNWLink"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl06
    intColumn = 7
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl06"
            .Caption = "KalkulationLNWLink"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt07
    intColumn = 8
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt07"
            .ControlSource = "RechnungBrutto"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl07
    intColumn = 8
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl07"
            .Caption = "RechnungBrutto"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    
    
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
