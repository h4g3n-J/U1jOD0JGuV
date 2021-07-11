Attribute VB_Name = "basCreateOffer"
Option Compare Database
Option Explicit

Public Sub BuildCreateOffer()
    
    'command message
    If gconVerbatim Then
        Debug.Print "basCreateOffer.BuildCreateOffer ausfuehren"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "frmCreateOffer"
    
    ' declare temporary form name
    Dim strTempFormName As String

    ' clear form
    basCreateOffer.ClearForm strFormName

    ' declare form
    Dim frm As Form
    Set frm = CreateForm
    
    ' write temporary form name to strFormName
    strTempFormName = frm.Name
    
    ' set form caption
    frm.Caption = strFormName
    
    ' create information grid
        ' set top left position
        Dim intLeft As Integer
        intLeft = 10000
        
        ' set top position
        Dim intTop As Integer
        intTop = 2430
        
        ' set column width
        Dim intColumnWidth(1) As Integer
        intColumnWidth(0) = 2540
        intColumnWidth(1) = 3120
    
        ' set number of rows
        Dim intNumberOfRows As Integer
        intNumberOfRows = 6
        
        Dim aintInformationGrid() As Integer
        ' aintInformationGrid = basAngebotSuchen.CalculateInformationGrid(2, intColumnWidth, intNumberOfRows, intLeft, intTop)
    
        ' create textboxes
        ' basAngebotSuchen.CreateTextbox strTempFormName, aintInformationGrid, intNumberOfRows
        
        ' create labels
        ' basAngebotSuchen.CreateLabel strTempFormName, aintInformationGrid, intNumberOfRows
    
        ' create captions
        ' Dim astrCaptionSettings() As String
        ' astrCaptionSettings = basAngebotSuchen.CaptionAndValueSettings(intNumberOfRows) ' get caption settings
        ' basAngebotSuchen.SetLabelCaption strTempFormName, astrCaptionSettings, intNumberOfRows ' set caption
    
    ' create command buttons
    ' basAngebotSuchen.CreateCommandButton strTempFormName, aintInformationGrid
    
    ' create subform
    ' basAngebotSuchen.CreateSubForm strTempFormName, aintInformationGrid
        
    ' close form
    DoCmd.Close acForm, strTempFormName, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strTempFormName
    
    'event message
    If gconVerbatim Then
        Debug.Print "basCreateOffer.BuildForm ausfuehren"
    End If
    
End Sub

' delete form
' 1. check if form exists
' 2. close if form is loaded
' 3. delete form
Private Sub ClearForm(ByVal strFormName As String)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basCreateOffer.ClearForm ausfuehren"
    End If
    
    Dim objDummy As Object
    For Each objDummy In Application.CurrentProject.AllForms
        If objDummy.Name = strFormName Then
            
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
                
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basCreateOffer.ClearForm: " & strFormName & " ist geoeffnet, Formular schlieﬂen"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "basCreateOffer.ClearForm: " & strFormName & " existiert bereits, Formular loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
    
    'event message
    If gconVerbatim Then
        Debug.Print "basCreateOffer.ClearForm ausgefuehrt"
    End If
        
End Sub
