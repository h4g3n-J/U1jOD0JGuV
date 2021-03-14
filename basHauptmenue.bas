Attribute VB_Name = "basHauptmenue"
' basHautpmenue

Option Compare Database
Option Explicit

' returns command button settings
Private Function CommandButtonsSettings()
    
    If gconVerbatim Then
        Debug.Print "basHauptmenue.CommandButtonsSettings ausfuehren"
    End If
    
    Dim avarCommandButtonsSettings(7, 3) As Variant

    avarCommandButtonsSettings(0, 0) = "name"
        avarCommandButtonsSettings(1, 0) = "Visible"
        avarCommandButtonsSettings(2, 0) = "Left"
        avarCommandButtonsSettings(3, 0) = "Top"
        avarCommandButtonsSettings(4, 0) = "Width"
        avarCommandButtonsSettings(5, 0) = "Height"
        avarCommandButtonsSettings(6, 0) = "Caption"
        avarCommandButtonsSettings(7, 0) = "OnClick"
    avarCommandButtonsSettings(0, 1) = "cmd0"
        avarCommandButtonsSettings(1, 1) = True
        avarCommandButtonsSettings(2, 1) = 100
        avarCommandButtonsSettings(3, 1) = 100
        avarCommandButtonsSettings(4, 1) = 1701
        avarCommandButtonsSettings(5, 1) = 283
        avarCommandButtonsSettings(6, 1) = "Ticket suchen"
        avarCommandButtonsSettings(7, 1) = "=OpenFrmAuftragSuchen()"
    avarCommandButtonsSettings(0, 2) = "cmd1"
        avarCommandButtonsSettings(1, 2) = True
        avarCommandButtonsSettings(2, 2) = 100
        avarCommandButtonsSettings(3, 2) = 600
        avarCommandButtonsSettings(4, 2) = 1701
        avarCommandButtonsSettings(5, 2) = 283 * 2
        avarCommandButtonsSettings(6, 2) = "Angebot " & vbCrLf & "suchen"
        avarCommandButtonsSettings(7, 2) = "=OpenFrmAngebotSuchen()"
    avarCommandButtonsSettings(0, 3) = "cmd2"
        avarCommandButtonsSettings(1, 3) = False
        avarCommandButtonsSettings(2, 3) = 100
        avarCommandButtonsSettings(3, 3) = 1383
        avarCommandButtonsSettings(4, 3) = 1701
        avarCommandButtonsSettings(5, 3) = 283
        avarCommandButtonsSettings(6, 3) = "Build Application"
        avarCommandButtonsSettings(7, 3) = "=BuildApplication()"
        
    CommandButtonsSettings = avarCommandButtonsSettings
End Function

Public Sub BuildFormHauptmenue()

    If gconVerbatim Then
        Debug.Print "basHautpmenue.BuildFormHauptmenue ausfuehren"
    End If
    
    ' define form name
    Dim strFormName As String
    strFormName = "frmHauptmenue"
    
    ' clear form
    basSupport.ClearForm strFormName
    
    ' create form
    Dim frmHauptmenue As Form
    Set frmHauptmenue = CreateForm
    
    ' save temporary form name in strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = frmHauptmenue.Name
    
    ' get Commandbuttons settings
    Dim avarCommandButtonsSettings As Variant
    avarCommandButtonsSettings = basHauptmenue.CommandButtonsSettings
    
    ' create Commandbutton
        Dim inti As Integer
        For inti = LBound(avarCommandButtonsSettings, 2) + 1 To UBound(avarCommandButtonsSettings, 2)
            Dim CmdButton As CommandButton
            Set CmdButton = CreateControl(strFormNameTemp, acCommandButton, acDetail)
            CmdButton.Name = avarCommandButtonsSettings(0, inti)
            CmdButton.Visible = avarCommandButtonsSettings(1, inti)
            CmdButton.Left = avarCommandButtonsSettings(2, inti)
            CmdButton.Top = avarCommandButtonsSettings(3, inti)
            CmdButton.Width = avarCommandButtonsSettings(4, inti)
            CmdButton.Height = avarCommandButtonsSettings(5, inti)
            
            ' handle visible = False
            If avarCommandButtonsSettings(1, inti) Then
                CmdButton.Caption = avarCommandButtonsSettings(6, inti)
                CmdButton.OnClick = avarCommandButtonsSettings(7, inti)
            End If
            
            Set CmdButton = Nothing
        Next
    
    ' set form caption
    frmHauptmenue.Caption = strFormName
    
    ' close and save form
    DoCmd.Close acForm, frmHauptmenue.Name, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basMain.FormularErstellen: " & strFormName & " erstellt"
    End If
    
End Sub

Public Function OpenFrmAuftragSuchen()
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHautpmenue.OpenFormAuftragSuchen ausfuehren"
    End If
    
    basSearchMain.OpenFormSearchMain "AuftragSuchen"
End Function

Public Function OpenFrmAngebotSuchen()
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHautpmenue.OpenFormAngebotSuchen ausfuehren"
    End If

    DoCmd.OpenForm "frmAngebotSuchen", acNormal
End Function


