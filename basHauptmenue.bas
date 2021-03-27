Attribute VB_Name = "basHauptmenue"
' basHautpmenue

Option Compare Database
Option Explicit

' returns command button settings
Private Function CommandButtonSettings()
    
    If gconVerbatim Then
        Debug.Print "basHauptmenue.CommandButtonSettings ausfuehren"
    End If
    
    Dim varSettings(3, 7) As Variant

    varSettings(0, 0) = "name"
        varSettings(0, 1) = "Visible"
        varSettings(0, 2) = "Left"
        varSettings(0, 3) = "Top"
        varSettings(0, 4) = "Width"
        varSettings(0, 5) = "Height"
        varSettings(0, 6) = "Caption"
        varSettings(0, 7) = "OnClick"
    varSettings(1, 0) = "cmd0"
        varSettings(1, 1) = True
        varSettings(1, 2) = 100
        varSettings(1, 3) = 100
        varSettings(1, 4) = 1701
        varSettings(1, 5) = 283
        varSettings(1, 6) = "Ticket suchen"
        varSettings(1, 7) = "=OpenFrmAuftragSuchen()"
    varSettings(2, 0) = "cmd1"
        varSettings(2, 1) = True
        varSettings(2, 2) = 100
        varSettings(2, 3) = 600
        varSettings(2, 4) = 1701
        varSettings(2, 5) = 283 * 2
        varSettings(2, 6) = "Angebot " & vbCrLf & "suchen"
        varSettings(2, 7) = "=OpenFrmAngebotSuchen()"
    varSettings(3, 0) = "cmd2"
        varSettings(3, 1) = True
        varSettings(3, 2) = 100
        varSettings(3, 3) = 1383
        varSettings(3, 4) = 1701
        varSettings(3, 5) = 283
        varSettings(3, 6) = "Build Application"
        varSettings(3, 7) = "=BuildApplication()"
        
    CommandButtonSettings = varSettings
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
    Dim avarCommandButtonSettings As Variant
    avarCommandButtonSettings = basHauptmenue.CommandButtonSettings
    
    ' create Commandbutton
        Dim inti As Integer
        For inti = LBound(avarCommandButtonSettings, 1) + 1 To UBound(avarCommandButtonSettings, 1)
            Dim CmdButton As CommandButton
            Set CmdButton = CreateControl(strFormNameTemp, acCommandButton, acDetail)
            CmdButton.Name = avarCommandButtonSettings(inti, 0)
            CmdButton.Visible = avarCommandButtonSettings(inti, 1)
            CmdButton.Left = avarCommandButtonSettings(inti, 2)
            CmdButton.Top = avarCommandButtonSettings(inti, 3)
            CmdButton.Width = avarCommandButtonSettings(inti, 4)
            CmdButton.Height = avarCommandButtonSettings(inti, 5)
            
            ' handle visible = False
            If avarCommandButtonSettings(1, inti) Then
                CmdButton.Caption = avarCommandButtonSettings(inti, 6)
                CmdButton.OnClick = avarCommandButtonSettings(inti, 7)
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


