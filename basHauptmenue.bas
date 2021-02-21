Attribute VB_Name = "basHauptmenue"
' basHautpmenue

Option Compare Database
Option Explicit

Private avarLayout(2, 4) As Variant

Private Sub LayoutConfig()
    ' 0 = object name
    ' 1 = object caption
    ' 2 = object visible
    avarLayout(0, 0) = "cmd0"
        avarLayout(1, 0) = "Ticket suchen"
        avarLayout(2, 0) = True
    avarLayout(0, 1) = "cmd1"
        avarLayout(1, 1) = "Liefergegenstand suchen"
        avarLayout(2, 1) = True
    avarLayout(0, 2) = "cmdEinzelauftrag"
        avarLayout(1, 2) = "Einzelaufträge"
        avarLayout(2, 2) = False
    avarLayout(0, 3) = "cmdKontinuierlicheLeistungen"
        avarLayout(1, 3) = "Liefergegenstand suchen"
        avarLayout(2, 3) = False
    avarLayout(0, 4) = "cmdForecast"
        avarLayout(1, 4) = "Forecast"
        avarLayout(2, 4) = False
End Sub

' open frmSearchMain and set textboxes and labels
' Public Sub OpenFormHauptmenue()
Public Function OpenFormHauptmenue()
    
    If gconVerbatim = True Then
        Debug.Print "basHauptmenue.OpenFormHauptmenue ausfuehren"
    End If
    
    ' Set Form name
    Dim strFormName As String
    strFormName = "frmHauptmenue"
    
    ' initialize LayoutConfig
    LayoutConfig
    
    ' set labels and textboxes
    Dim inti As Integer
        
    ' set command buttons
    For inti = LBound(avarLayout, 2) To UBound(avarLayout, 2)
        ' set caption
        Forms.Item(strFormName).Controls.Item(avarLayout(0, inti)).Caption = avarLayout(1, inti)
        ' set visibility
        Forms.Item(strFormName).Controls.Item(avarLayout(0, inti)).Visible = avarLayout(2, inti)
    Next
    
End Function

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
    
    ' create Commandbutton
    basHauptmenue.CreateCommandButton (strFormNameTemp)
    
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
    
    ' Dim aControls As Variant
       
    ' define button name
    Dim strButtonName As String
    strButtonName = "cmd0"
    
    ' define Postion Left
    Dim intPositionLeft As Integer
    intPositionLeft = 100
    
    ' define Position Top
    Dim intPositionTop As Integer
    intPositionTop = 100
    
    ' define Size Width
    Dim intSizeWidth As Integer
    intSizeWidth = 100
    
    ' definde Size Heigth
    Dim intSizeHeigth As Integer
    intSizeHeigth = 200
    
    ' define Caption
    Dim strCaption As String
    strCaption = "Auftrag bearbeiten"
    
    ' define OnClick behaviour
    Dim strOnClick As String
    strOnClick = "=OpenFormAuftragSuchen()"
    
    ' create Commandbutton
    Dim CmdButton As CommandButton
    ' Set CmdButton = CreateControl(strFormName, acCommandButton, acDetail, , , intPositionLeft, intPositionTop, intSizeWidth, intSizeHeigth)
    Set CmdButton = CreateControl(strFormName, acCommandButton, acDetail)
    
        ' set Name
        CmdButton.Name = strButtonName
        
        ' set Caption
        CmdButton.Caption = strCaption
    
        ' set OnClick
        CmdButton.OnClick = strOnClick
    
End Sub

Public Function OpenFormAuftragSuchen()
    
    If gconVerbatim Then
        Debug.Print "basHautpmenue.OpenFormAuftragSuchen ausfuehren"
    End If
    
    basSearchMain.OpenFormAuftrag
End Function


