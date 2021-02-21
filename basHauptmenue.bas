Attribute VB_Name = "basHauptmenue"
' basHautpmenue

Option Compare Database
Option Explicit

Private avarCommandButtonsLayout(7, 1) As Variant

Private Sub LayoutConfig()
    ' 0 = object name
    ' 1 = visible
    ' 2 = position left
    ' 3 = position top
    ' 4 = width
    ' 5 = heigth
    ' 6 = caption
    ' 7 = OnClick
    
    avarCommandButtonsLayout(0, 0) = "cmd0"
        avarCommandButtonsLayout(1, 0) = True
        avarCommandButtonsLayout(2, 0) = 100
        avarCommandButtonsLayout(3, 0) = 100
        avarCommandButtonsLayout(4, 0) = 1701
        avarCommandButtonsLayout(5, 0) = 283
        avarCommandButtonsLayout(6, 0) = "Ticket suchen"
        avarCommandButtonsLayout(7, 0) = "=OpenFormAuftragSuchen()"
    avarCommandButtonsLayout(0, 1) = "cmd1"
        avarCommandButtonsLayout(1, 1) = True
        avarCommandButtonsLayout(2, 1) = True
        avarCommandButtonsLayout(3, 1) = True
        avarCommandButtonsLayout(4, 1) = True
        avarCommandButtonsLayout(5, 1) = True
        avarCommandButtonsLayout(6, 1) = "Liefergegenstand suchen"
        avarCommandButtonsLayout(7, 1) = ""
    ' avarCommandButtonsLayout(0, 2) = "cmdEinzelauftrag"
        ' avarCommandButtonsLayout(1, 2) = "Einzelaufträge"
        ' avarCommandButtonsLayout(2, 2) = False
    ' avarCommandButtonsLayout(0, 3) = "cmdKontinuierlicheLeistungen"
        ' avarCommandButtonsLayout(1, 3) = "Liefergegenstand suchen"
        ' avarCommandButtonsLayout(2, 3) = False
    ' avarCommandButtonsLayout(0, 4) = "cmdForecast"
        ' avarCommandButtonsLayout(1, 4) = "Forecast"
        ' avarCommandButtonsLayout(2, 4) = False
End Sub

Private Sub CreateFormHautpmenue()

    If gconVerbatim Then
        Debug.Print "basHautpmenue.CreateFormHauptmenue ausfuehren"
    End If
    
    ' define form name
    Dim strFormName As String
    strFormName = "frmHauptmenue"
    
    ' clear form
    basSupport.ClearForm strFormName
    
    Dim frmHauptmenue As Form
    Set frmHauptmenue = CreateForm
    
    ' save temporary form name in strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = frmHauptmenue.Name
    
    ' set layout
    basHauptmenue.LayoutConfig
        
    ' create Commandbutton
    basHauptmenue.CreateCommandButton strFormNameTemp
    
    ' set form caption
    frmHauptmenue.Caption = strFormName
    
    ' close and save form
    DoCmd.Close acForm, frmHauptmenue.Name, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basMain.FormularErstelle: " & strFormName & " erstellt"
    End If
    
End Sub

' Private Sub CreateCommandButton(ByVal aControls As Variant)
' Private Function CreateCommandButton(ByVal strFormName)
Private Sub CreateCommandButton(ByVal strFormName As String)
' Private Sub CreateCommandButton(ByVal strFormName, strButtonName, strCaption, strOnClick As String, intPositionLeft, intPositionTop, intSizeWidth, intSizeHeigth As Integer)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basHautpmenue.CreateCommandButton ausfuehren"
    End If
    
    ' create Commandbutton
        Dim inti As Integer
        For inti = LBound(avarCommandButtonsLayout, 2) To UBound(avarCommandButtonsLayout, 2) - 1
            Dim CmdButton As CommandButton
            Set CmdButton = CreateControl(strFormName, acCommandButton, acDetail)
            
            ' set Name
                CmdButton.Name = avarCommandButtonsLayout(0, inti)
                
            ' set position left
                CmdButton.Left = avarCommandButtonsLayout(2, inti)
                
            ' set position top
            CmdButton.Top = avarCommandButtonsLayout(3, inti)
            
            ' set width
            CmdButton.Width = avarCommandButtonsLayout(4, inti)
            
            ' set heigth
            CmdButton.Height = avarCommandButtonsLayout(5, inti)
            
            ' set visible
            CmdButton.Visible = avarCommandButtonsLayout(1, inti)
            
            ' handler in case visible = false
            If avarCommandButtonsLayout(1, inti) Then
                ' set Caption
                CmdButton.Caption = avarCommandButtonsLayout(6, inti)
                ' set OnClick
                CmdButton.OnClick = avarCommandButtonsLayout(7, inti)
            End If
            
            Set CmdButton = Nothing
        Next
End Sub

Public Function OpenFormAuftragSuchen()
    
    If gconVerbatim Then
        Debug.Print "basHautpmenue.OpenFormAuftragSuchen ausfuehren"
    End If
    
    basSearchMain.OpenFormAuftrag
End Function


