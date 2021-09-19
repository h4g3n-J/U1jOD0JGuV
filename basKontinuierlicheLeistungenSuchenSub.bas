Attribute VB_Name = "basKontinuierlicheLeistungenSuchenSub"
Option Compare Database
Option Explicit

' build form
Public Sub BuildKontinuierlicheLeistungenSuchenSub()
    
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenSuchenSub.BuildKontinuierlicheLeistungenSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmKontinuierlicheLeistungenSuchenSub"
    
    ' clear form
    basKontinuierlicheLeistungenSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create query qryKontinuierlicheLeistungenAuswahl
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenAuswahl"
    basKontinuierlicheLeistungenSuchenSub.BuildqryKontinuierlicheLeistungenAuswahl strQueryName
    
    ' set recordsetSource
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
        
            intNumberOfColumns = 3
            intNumberOfRows = 2
            intColumnWidth = 1600
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basKontinuierlicheLeistungenSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "KLID"
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "KLID"
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "AngebotBrutto"
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl01
    intColumn = 2
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "AngebotBrutto"
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "Bemerkung"
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basKontinuierlicheLeistungenSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basKontinuierlicheLeistungenSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basKontinuierlicheLeistungenSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basKontinuierlicheLeistungenSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
        
    ' set OnCurrent methode
    objForm.OnCurrent = "=SelectKontinuierlicheLeistungen()"
    
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
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.BuildKontinuierlicheLeistungenSuchenSub executed"
    End If
        
End Sub

' load recordset to destination form
Public Function SelectKontinuierlicheLeistungen()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenSuchenSub.SelectKontinuierlicheLeistungen"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmKontinuierlicheLeistungenSuchen"
    
    ' check if frmKontinuierlicheLeistungenSuchen exists (Error Code: 1)
    Dim bolFormExists As Boolean
    bolFormExists = False
    
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolFormExists = True
        End If
    Next
    
    If Not bolFormExists Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.SelectKontinuierlicheLeistungen aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmKontinuierlicheLeistungenSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.SelectKontinuierlicheLeistungen aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "KLID"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class KontinuierlicheLeistungen
    Dim KontinuierlicheLeistungen As clsKontinuierlicheLeistungen
    Set KontinuierlicheLeistungen = New clsKontinuierlicheLeistungen
    
    ' select recordset
    KontinuierlicheLeistungen.SelectRecordset varRecordsetName
    
    ' show recordset
    ' Forms.Item(strFormName).Controls.Item("insert textboxName here") = CallByName(Auftrag, "insert Attribute Name here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(KontinuierlicheLeistungen, "KLID", VbGet)
    Forms.Item(strFormName).Controls.Item("txt02") = CallByName(KontinuierlicheLeistungen, "AngebotBrutto", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(KontinuierlicheLeistungen, "Bemerkung", VbGet)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.SelectKontinuierlicheLeistungen executed"
    End If
    
ExitProc:
    Set KontinuierlicheLeistungen = Nothing
End Function

' returns array
' (column, row, property)
' properties: 0 - Left, 1 - Top, 2 - Width, 3 - Height
' calculates left, top, width and height parameters
Private Function CalculateInformationGrid(ByVal intNumberOfColumns As Integer, ByRef aintColumnWidth() As Integer, ByVal intNumberOfRows As Integer, Optional ByVal intLeft As Integer = 10000, Optional ByVal intTop As Integer = 2430)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.CalculateTableSetting ausfuehren"
    End If
    
    intNumberOfColumns = intNumberOfColumns - 1
    intNumberOfRows = intNumberOfRows - 1
    
    ' column dimension
    Const cintHorizontalSpacing As Integer = 60
            
    ' row dimension
    Dim intRowHeight As Integer
    intRowHeight = 330
    
    Const cintVerticalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    Dim aintGridSettings() As Integer
    ReDim aintGridSettings(intNumberOfColumns, intNumberOfRows, cintNumberOfProperties)
    
    ' compute cell position properties
    Dim inti As Integer
    Dim intj As Integer
    For inti = 0 To intNumberOfColumns
        ' For intr = 0 To cintNumberOfRows
        For intj = 0 To intNumberOfRows
            ' set column left
            aintGridSettings(inti, intj, 0) = intLeft + inti * (aintColumnWidth(inti) + cintHorizontalSpacing)
            ' set row top
            aintGridSettings(inti, intj, 1) = intTop + intj * (intRowHeight + cintVerticalSpacing)
            ' set column width
            aintGridSettings(inti, intj, 2) = aintColumnWidth(inti)
            ' set row height
            aintGridSettings(inti, intj, 3) = intRowHeight
        Next
    Next

    CalculateInformationGrid = aintGridSettings
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.CalculateInformationGrid ausgefuehrt"
    End If

End Function

Private Sub ClearForm(ByVal strFormName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenSuchenSub.ClearForm"
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
        Debug.Print "basKontinuierlicheLeistungenSuchenSub executed"
    End If
    
End Sub

Public Sub SearchKontinuierlicheLeistungen(Optional varSearchTerm As Variant)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenSuchenSub.SearchKontinuierlicheLeistungen"
    End If
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' define query name
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenAuswahl"
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basKontinuierlicheLeistungenSuchenSub.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = " SELECT qryKontinuierlicheLeistungen.*" & _
                    " FROM qryKontinuierlicheLeistungen" & _
                    " WHERE qryKontinuierlicheLeistungen.KLID LIKE '*" & strSearchTerm & "*'" & _
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
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.SearchKontinuierlicheLeistungen executed"
    End If

End Sub

Private Sub TestSearchKontinuierlicheLeistungen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLeistungserfassungsblattSuchenSub.TestSearchKontinuierlicheLeistungen"
    End If
        
    ' build query qryRechnungSuchen
    Dim strQueryName As String
    strQueryName = "qryKontinuierlicheLeistungenSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblKontinuierlicheLeistungen"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "KLID"
    
    basLeistungserfassungsblattSuchenSub.SearchKontinuierlicheLeistungen strQueryName, strQuerySource, strPrimaryKey
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentData.AllQueries
        If objForm.Name = strQueryName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Procedure successful: " & vbCr & vbCr & strQueryName & " detected", vbOKOnly, "basLeistungserfassungsblattSuchenSub.TestSearchKontinuierlicheLeistungen"
    Else
        MsgBox "Failure: " & vbCr & vbCr & strQueryName & " was not detected", vbCritical, "basLeistungserfassungsblattSuchenSub.TestSearchKontinuierlicheLeistungen"
    End If
        
    ' event message
    If gconVerbatim Then
        Debug.Print "basLeistungserfassungsblattSuchenSub.TestSearchKontinuierlicheLeistungen executed"
    End If
    
End Sub

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Long, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basKontinuierlicheLeistungenSuchenSub.CalculateGrid"
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
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.CalculateGrid executed"
    End If
    
End Function

Private Function TestCalculateGrid()

    Dim aintInformationGrid() As Integer
        
        Dim intNumberOfColumns As Integer
        Dim intNumberOfRows As Integer
        Dim intColumnWidth As Integer
        Dim intRowHeight As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim intColumn As Integer
        Dim intRow As Integer
        
            intNumberOfColumns = 15
            intNumberOfRows = 2
            intColumnWidth = 1600
            intRowHeight = 330
            intLeft = 50
            intTop = 50
    
    ReDim aintInformationGrid(intNumberOfColumns - 1, intNumberOfRows - 1)
    
    aintInformationGrid = basKontinuierlicheLeistungenSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = False
    
    If Not bolOutput Then
        TestCalculateGrid = aintInformationGrid
        Exit Function
    End If
    
    For intColumn = 0 To UBound(aintInformationGrid, 1)
        For intRow = 0 To UBound(aintInformationGrid, 2)
            Debug.Print "column " & intColumn & ", row " & intRow & ", left: " & aintInformationGrid(intColumn, intRow, 0)
            Debug.Print "column " & intColumn & ", row " & intRow & ", top: " & aintInformationGrid(intColumn, intRow, 1)
            Debug.Print "column " & intColumn & ", row " & intRow & ", width: " & aintInformationGrid(intColumn, intRow, 2)
            Debug.Print "column " & intColumn & ", row " & intRow & ", height: " & aintInformationGrid(intColumn, intRow, 3)
        Next
    Next
    
    TestCalculateGrid = aintInformationGrid
    
End Function

' get left from grid
Private Function GetLeft(alngGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Long
    
    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = alngGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetLeft executed"
    End If
    
End Function

Private Sub TestGetLeft()

    Dim aintGrid() As Integer
    aintGrid = basKontinuierlicheLeistungenSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If Not bolOutput Then
        Exit Sub
    End If
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Left: " & basKontinuierlicheLeistungenSuchenSub.GetLeft(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
End Sub

' get top from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetTop: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetTop executed"
    End If
    
End Function

Private Sub TestGetTop()

    Dim aintGrid() As Integer
    aintGrid = basKontinuierlicheLeistungenSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Top: " & basKontinuierlicheLeistungenSuchenSub.GetTop(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub

' get width from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetWidth: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenSuchenSub.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetWidth executed"
    End If
    
End Function

Private Sub TestGetWidth()

    Dim aintGrid() As Integer
    aintGrid = basKontinuierlicheLeistungenSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Width: " & basKontinuierlicheLeistungenSuchenSub.GetWidth(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub

' get height from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetHeight: column 0 is not available"
        MsgBox "basKontinuierlicheLeistungenSuchenSub.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basKontinuierlicheLeistungenSuchenSub.GetHeight executed"
    End If
    
End Function

Private Sub TestGetHeight()

    Dim aintGrid() As Integer
    aintGrid = basKontinuierlicheLeistungenSuchenSub.TestCalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Height: " & basKontinuierlicheLeistungenSuchenSub.GetHeight(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If

End Sub



