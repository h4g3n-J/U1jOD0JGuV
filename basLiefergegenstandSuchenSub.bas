Attribute VB_Name = "basLiefergegenstandSuchenSub"
Option Compare Database
Option Explicit

Public Sub BuildLiefergegenstandSuchenSub()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandSuchenSub.BuildLiefergegenstandSuchenSub"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmLiefergegenstandSuchenSub"
    
    ' clear form
    basLiefergegenstandSuchenSub.ClearForm strFormName
    
    ' declare form
    Dim objForm As Form
    
    ' create form
    Set objForm = CreateForm
    
    ' declare temporary form name
    Dim strTempFormName As String
    strTempFormName = objForm.Name
    
    ' build query
    Dim strQueryName As String
    strQueryName = "qryLiefergegenstandSuchen"
    
    Dim strQuerySource As String
    strQuerySource = "tblLiefergegenstand"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "LiefergegenstandID"
    
    basLiefergegenstandSuchenSub.SearchLiefergegenstand strQueryName, strQuerySource, strPrimaryKey
    
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
    
    aintInformationGrid = basLiefergegenstandSuchenSub.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeight)
    
    Dim lblLabel As Label
    Dim txtTextbox As TextBox
    
    ' create textbox before label, so label can be associated
    'txt00
    intColumn = 1
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt00"
            .ControlSource = "LiefergegenstandID"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
            .IsHyperlink = False
        End With
    
    'lbl00
    intColumn = 1
    intRow = 1
    Set lblLabel = CreateControl(strTempFormName, acLabel, acDetail, "txt00")
        With lblLabel
            .Name = "lbl00"
            .Caption = "LiefergegenstandID"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt01
    intColumn = 2
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt01"
            .ControlSource = "RechnungNr"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt02
    intColumn = 3
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt02"
            .ControlSource = "Bemerkung"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt03
    intColumn = 4
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt03"
            .ControlSource = "BelegID"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
        
    'txt04
    intColumn = 5
    intRow = 2
    Set txtTextbox = CreateControl(strTempFormName, acTextBox, acDetail)
        With txtTextbox
            .Name = "txt04"
            .ControlSource = "Brutto"
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
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
            .Left = basLiefergegenstandSuchenSub.GetLeft(aintInformationGrid, intColumn, intRow)
            .Top = basLiefergegenstandSuchenSub.GetTop(aintInformationGrid, intColumn, intRow)
            .Width = basLiefergegenstandSuchenSub.GetWidth(aintInformationGrid, intColumn, intRow)
            .Height = basLiefergegenstandSuchenSub.GetHeight(aintInformationGrid, intColumn, intRow)
            .Visible = True
        End With
    ' column added? -> update intNumberOfColumns
    
    ' set oncurrent methode
    objForm.OnCurrent = "=selectLiefergegenstand()"
        
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
        Debug.Print "basLiefergegenstandSuchenSub.BuildAuftragSuchenSub executed"
    End If
    
End Sub

Private Sub ClearForm(ByVal strFormName As String)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandSuchenSub.ClearForm"
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
        Debug.Print "basLiefergegenstandSuchenSub.ClearForm executed"
    End If
    
End Sub

Private Sub TestClearForm()

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basLiefergegenstandSuchenSub.ClearForm"
    End If
    
    Dim strFormName As String
    strFormName = "frmLiefergegenstandSuchenSub"
    
    basLiefergegenstandSuchenSub.ClearForm strFormName
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
        
    Dim objForm As Object
    For Each objForm In Application.CurrentProject.AllForms
        If objForm.Name = strFormName Then
            bolObjectExists = True
        End If
    Next
    
    If bolObjectExists Then
        MsgBox "Failure: " & vbCr & vbCr & strFormName & " was not deleted.", vbCritical, "basLiefergegenstandSuchenSub.TestClearForm"
    Else
        MsgBox "Procedure successful: " & vbCr & vbCr & strFormName & " was not detected", vbOKOnly, "basLiefergegenstandSuchenSub.TestClearForm"
    End If
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basLiefergegenstandSuchenSub.TestClearForm executed"
    End If
    
End Sub
