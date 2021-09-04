Attribute VB_Name = "basLeistungserfassungsblattSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildRechnungSuchenSub()

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


