Attribute VB_Name = "basAuftragSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildAuftragSuchenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.BuildAuftragSuchenSub "
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmAuftragSuchenSub"
    
    ' clear form
    basAuftragSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' create query qryAuftragSuchen
    Dim strQueryName As String
    strQueryName = "qryAuftragSuchen"
    basAuftragSuchenSub.BuildQuery strQueryName
    
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
    
    aintInformationGrid = basAuftragSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
            
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "AftrID"
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
            .Caption = "AftrID"
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
            .ControlSource = "AftrTitel"
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
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt01")
        With lblLabel
            .Name = "lbl01"
            .Caption = "AftrTitel"
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
            .ControlSource = "StatusKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
            .IsHyperlink = False
        End With
        
    'lbl02
    intColumn = 3
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt02")
        With lblLabel
            .Name = "lbl02"
            .Caption = "StatusKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "OwnerKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
            .IsHyperlink = False
        End With
        
    'lbl03
    intColumn = 4
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt03")
        With lblLabel
            .Name = "lbl03"
            .Caption = "OwnerKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "PrioritaetKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
            .IsHyperlink = False
        End With
        
    'lbl04
    intColumn = 5
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt04")
        With lblLabel
            .Name = "lbl04"
            .Caption = "PrioritaetKey"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = False
        End With
        
    'txt05
    intColumn = 6
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt05"
            .ControlSource = "ParentKey"
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
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt05")
        With lblLabel
            .Name = "lbl05"
            .Caption = "ParentKey"
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
            .ControlSource = "Bemerkung"
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
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt06")
        With lblLabel
            .Name = "lbl06"
            .Caption = "Bemerkung"
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
            .ControlSource = "BeginnSoll"
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
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt07")
        With lblLabel
            .Name = "lbl07"
            .Caption = "BeginnSoll"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt08
    intColumn = 9
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt08"
            .ControlSource = "EndeSoll"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl08
    intColumn = 9
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt08")
        With lblLabel
            .Name = "lbl08"
            .Caption = "EndeSoll"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt09
    intColumn = 10
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt09"
            .ControlSource = "Erstellt"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl09
    intColumn = 10
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt09")
        With lblLabel
            .Name = "lbl09"
            .Caption = "Erstellt"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt10
    intColumn = 11
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt10"
            .ControlSource = "kunde"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
        
    'lbl10
    intColumn = 11
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt10")
        With lblLabel
            .Name = "lbl10"
            .Caption = "Kunde"
            .Left = basAuftragSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basAuftragSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basAuftragSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basAuftragSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    ' column added? -> update intNumberOfColumns
        
    ' set oncurrent methode
    objForm.OnCurrent = "=selectAuftrag()"
    
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
        Debug.Print "basAuftragSuchenSub.BuildAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.ClearForm"
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
        Debug.Print "basAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub BuildQuery(ByVal strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.BuildQuery"
    End If
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' delete query
    basAuftragSuchenSub.ClearQuery strQueryName
    
    ' declare query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    ' set query Name
    qdfQuery.Name = strQueryName
    
    ' set query SQL
    qdfQuery.SQL = " SELECT tblAuftrag.*" & _
                        " FROM tblAuftrag" & _
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
        Debug.Print "basAuftragSuchenSub.BuildQuery executed"
    End If
    
End Sub

Private Sub ClearQuery(ByVal strQueryName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.clearQuery"
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
        Debug.Print "basAuftragSuchenSub.clearQuery executed"
    End If
    
End Sub

Public Function selectAuftrag()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.SelectAuftrag"
    End If
    
    ' declare form name
    Dim strFormName As String
    strFormName = "frmAuftragSuchen"
    
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
        Debug.Print "basAuftragSuchenSub.selectAuftrag aborted, Error Code: 1"
        Exit Function
    End If
    
    ' if frmAuftragSuchen not isloaded go to exit (Error Code: 2)
    If Not Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
        Debug.Print "basAuftragSuchenSub.selectAuftrag aborted, Error Code: 2"
        Exit Function
    End If
    
    ' declare control object name
    Dim strControlObjectName As String
    strControlObjectName = "frbSubForm"
    
    ' declare reference attribute
    Dim strReferenceAttributeName As String
    strReferenceAttributeName = "AftrID"
    
    ' set recordset origin
    Dim varRecordsetName As Variant
    varRecordsetName = Forms.Item(strFormName).Controls(strControlObjectName).Controls(strReferenceAttributeName)
    
    ' initiate class Auftrag
    Dim Auftrag As clsAuftrag
    Set Auftrag = New clsAuftrag
    
    ' select recordset
    Auftrag.SelectRecordset varRecordsetName
    
    ' show recordset
    ' Forms.Item(strFormName).Controls.Item("insert textboxName here") = CallByName(Auftrag, "insert Attribute Name here", VbGet)
    Forms.Item(strFormName).Controls.Item("txt00") = CallByName(Auftrag, "AftrID", VbGet)
    Forms.Item(strFormName).Controls.Item("txt01") = CallByName(Auftrag, "AftrTitel", VbGet)
    Forms.Item(strFormName).Controls.Item("txt02") = CallByName(Auftrag, "StatusKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt03") = CallByName(Auftrag, "OwnerKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt04") = CallByName(Auftrag, "PrioritaetKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt05") = CallByName(Auftrag, "ParentKey", VbGet)
    Forms.Item(strFormName).Controls.Item("txt06") = CallByName(Auftrag, "Bemerkung", VbGet)
    Forms.Item(strFormName).Controls.Item("txt07") = CallByName(Auftrag, "BeginnSoll", VbGet)
    Forms.Item(strFormName).Controls.Item("txt08") = CallByName(Auftrag, "EndeSoll", VbGet)
    Forms.Item(strFormName).Controls.Item("txt09") = CallByName(Auftrag, "Erstellt", VbGet)
    Forms.Item(strFormName).Controls.Item("txt10") = CallByName(Auftrag, "kunde", VbGet)
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.SelectAuftrag executed"
    End If
    
End Function

Private Function CalculateGrid(ByVal intNumberOfColumns As Integer, ByVal intNumberOfRows As Integer, ByVal intLeft As Integer, ByVal intTop As Integer, ByVal intColumnWidth As Integer, ByVal intRowHeight As Integer)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basAuftragSuchenSub.CalculateGrid"
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
        Debug.Print "basAuftragSuchenSub.CalculateGrid executed"
    End If
    
End Function

' get left from grid
Private Function GetLeft(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchenSub.GetLeft: column 0 is not available"
        MsgBox "basAuftragSuchenSub.GetLeft: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetLeft = aintGrid(intColumn - 1, intRow - 1, 0)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.GetLeft executed"
    End If
    
End Function

' get top from grid
Private Function GetTop(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchenSub.GetTop: column 0 is not available"
        MsgBox "basAuftragSuchenSub.GetTop: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetTop = aintGrid(intColumn - 1, intRow - 1, 1)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.GetTop executed"
    End If
    
End Function

' get width from grid
Private Function GetWidth(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchenSub.GetWidth: column 0 is not available"
        MsgBox "basAuftragSuchenSub.GetWidth: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetWidth = aintGrid(intColumn - 1, intRow - 1, 2)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.GetWidth executed"
    End If
    
End Function

' get height from grid
Private Function GetHeight(aintGrid As Variant, ByVal intColumn As Integer, ByVal intRow As Integer) As Integer
    
    If intColumn = 0 Then
        Debug.Print "basAuftragSuchenSub.GetHeight: column 0 is not available"
        MsgBox "basAuftragSuchenSub.GetHeight: column 0 is not available. Please choose a higher value", vbCritical, "Error"
        Exit Function
    End If
    
    GetHeight = aintGrid(intColumn - 1, intRow - 1, 3)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basAuftragSuchenSub.GetHeight executed"
    End If
    
End Function

