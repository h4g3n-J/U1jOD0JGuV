Attribute VB_Name = "basTest"
' basTest

Option Compare Database
Option Explicit

Private Sub TestAuftragHinzufuegen()
    Dim TestKlasse As clsAuftrag
    Set TestKlasse = New clsAuftrag
    
    TestKlasse.AddRecordset
    
    ' Clean up
    Set TestKlasse = Nothing
End Sub

Private Sub TestAngebotHinzufuegen()
    Dim TestKlasse As clsAngebot
    Set TestKlasse = New clsAngebot
    
    TestKlasse.AddRecordset
    
    ' Clean up
    Set TestKlasse = Nothing
End Sub

Private Sub TestRecordsetExists()
    Debug.Print "TestRecordsetExists: " _
        & basSupport.RecordsetExists("tblAuftrag", "AftrID", "345")
End Sub

Private Sub TestForEach()
    Dim astrTable(1) As String
    Dim varTableName As Variant
    
    astrTable(0) = "tblAngebot"
    astrTable(1) = "tblAuftrag"
    
    For Each varTableName In astrTable
        Debug.Print varTableName
    Next
End Sub

Private Sub TestForNext()
    Dim astrTable(1, 2) As String
    Dim lngi As Long
    Dim lngj As Long
    
    Dim bolTest As Boolean
    bolTest = True
    
    astrTable(0, 0) = "test 0, 0"
    astrTable(0, 1) = "test 0, 1"
    astrTable(0, 2) = "test 0, 2"
    astrTable(1, 0) = "test 1, 0"
    astrTable(1, 1) = "test 1, 1"
    astrTable(1, 2) = "test 1, 2"
    
    astrTable(0, 0) = "test 0, 0"
    astrTable(0, 1) = "test 0, 1"
    astrTable(0, 2) = "test 0, 2"
    astrTable(1, 0) = "test 1, 0"
    astrTable(1, 1) = CStr(bolTest)
    astrTable(1, 2) = "test 1, 2"
    
    For lngi = LBound(astrTable, 1) To UBound(astrTable, 1)
        For lngj = LBound(astrTable, 2) To UBound(astrTable, 2)
            Debug.Print astrTable(lngi, lngj)
        Next
    Next
    
    Debug.Print "CBool(astrTable(1, 1)): "; CBool(astrTable(1, 1))
End Sub

Private Sub TestRechnungHinzufuegen()
    Dim TestKlasse As clsRechnung
    Set TestKlasse = New clsRechnung
    
    TestKlasse.AddRecordset
    
    ' Clean up
    Set TestKlasse = Nothing
End Sub

Private Sub testLeistungserfassungsblattHinzufuegen()
    Dim TestKlasse As clsLeistungserfassungungsblatt
    Set TestKlasse = New clsLeistungserfassungungsblatt
    
    TestKlasse.AddRecordset
    
    ' Clean up
    Set TestKlasse = Nothing
End Sub

Private Sub testDatensatzLaden()
    Dim TestKlasse As clsAuftrag
    Set TestKlasse = New clsAuftrag
    
    TestKlasse.SelectRecordset "BCH25900", True
    
    Set TestKlasse = Nothing
End Sub

Private Sub TestCollection()
    Dim intInteger As Integer
    intInteger = 100
    
    Dim strString As String
    strString = "Hallo"
    
    Dim colTestCollection As Collection
    Set colTestCollection = New Collection
    
    With colTestCollection
        .Add intInteger
        .Add strString
    End With
    
    Dim inti As Integer
    For inti = 1 To 2
        Debug.Print "TestCollection: colTestCollection.Item(" & inti & ") = " & colTestCollection.Item(inti)
    Next
    
    Dim varEintrag As Variant
    For Each varEintrag In colTestCollection
        Debug.Print "TestCollection: colTestCollection.Item = " & varEintrag
    Next
    
    Debug.Print "TestCollection: colTestCollection.Item(1) =" & colTestCollection.Item(1)
    Debug.Print "TestCollection: colTestCollection.Item(2) =" & colTestCollection.Item(2)
    ' die Werte der Collection sind schreibgeschützt
    ' colTestCollection.Item(2) = "Welt"
    Debug.Print "TestCollection: colTestCollection.Item(2) =" & colTestCollection.Item(2)
End Sub

Private Sub TestArray()
    Dim intInteger As Integer
    intInteger = 100
    
    Dim strString As String
    strString = "Hallo"
    
    Dim varArray(1) As Variant
        
    varArray(0) = intInteger
    varArray(1) = strString
    
    Dim inti As Integer
    For inti = 0 To 1
        Debug.Print "TestCollection: varArray(" & inti & ") = " & varArray(inti)
    Next
    
    varArray(1) = "Welt"
    
    Debug.Print "TestCollection: varArray(1) = " & varArray(1)
    Debug.Print "TestCollection: strString = " & strString
    
    ' die Werte der Collection sind schreibgeschützt
    ' colTestCollection.Item(2) = "Welt"
End Sub

Private Sub TestFindItemInArray()
    Dim varArray(1, 1) As Variant
    
    varArray(0, 0) = "nill"
    varArray(0, 1) = "one"
    varArray(1, 0) = 1
    varArray(1, 1) = 2
    
    Debug.Print "basTest.TestFindItemInArray: " & basSupport.FindItemInArray(varArray, "one", True)
End Sub

Private Sub TestFindItemArrayInProperty()
    Dim TestKlasse As clsAuftrag
    Set TestKlasse = New clsAuftrag
    
    TestKlasse.AftrID = "1"
    TestKlasse.Erstellt = #1/1/1900 11:59:59 PM#
    Debug.Print "basTest.TestFindItemArrayInProperty: TestKlasse.AftrID = " & TestKlasse.AftrID
    Debug.Print "basTest.TestFindItemArrayInProperty: TestKlasse.Erstellt = " & TestKlasse.Erstellt
End Sub

Public Sub TestformularErstellen()
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basMain.FomularErstellen ausfuehren"
    End If
    
    ' set form name
    Dim strFormName As String
    strFormName = "Testformular"
    
    ' initiate Formular
    Dim objForm As Form
    
    ' create Formular
    Set objForm = CreateForm
    
    ' create command button
    Dim CmdButton As CommandButton
    Set CmdButton = CreateControl(objForm.Name, acCommandButton, acDetail, , , 100, 100)
    
        ' set commandbutton caption
        CmdButton.Caption = "Auftrag Suchen oeffnen"
        
        ' set onclick behaviour
        CmdButton.OnClick = "=AuftragBearbeitenOeffnen()"
        
    ' save temporary form name in variable strFormNameTemp
    Dim strFormNameTemp As String
    strFormNameTemp = objForm.Name
        
    ' close if form is loaded
    ' delete if form already exists
    DeleteForm strFormName
    
    ' set objForm.Caption
    objForm.Caption = strFormName
    
    ' close and save form
    DoCmd.Close acForm, objForm.Name, acSaveYes
    
    ' rename form
    DoCmd.Rename strFormName, acForm, strFormNameTemp
    
    If gconVerbatim Then
        Debug.Print "basMain.FormularErstellen: " & strFormName & " erstellt"
    End If
End Sub

Public Function AuftragBearbeitenOeffnen()
    DoCmd.OpenForm "frmSearchMain"
End Function

Private Sub DeleteForm(ByVal strFormName As String)
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basMain.DeleteForm ausfuehren"
    End If
    
    ' check if form already exists
    ' check if formular is loaded
    ' close loaded form
    ' delete loaded form
    Dim objDummy As Object
    For Each objDummy In Application.CurrentProject.AllForms
        If objDummy.Name = strFormName Then
            
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basMain.FormularErstellen: " & strFormName & " ist geoeffnet, Formular schließen"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "basMain.FomularErstellen: " & strFormName & " existiert bereits, Formular loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
    
End Sub

Sub TestAllforms()

Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

For Each obj In dbs.AllForms
    Debug.Print obj.Name
Next obj

End Sub

' only in access project, not available in access datebase
Sub TestAllFunctions()

Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentData

For Each obj In dbs.AllFunctions
    Debug.Print obj.Name
Next obj

End Sub

Sub TestAllMacros()

Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

For Each obj In dbs.AllMacros
    Debug.Print obj.Name
Next obj

End Sub

Sub TestAllModules()

Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject

For Each obj In dbs.AllModules
    Debug.Print obj.Name
Next obj

End Sub

Sub TestBuildQryAngebotAuswahl()
    basBuild.BuildQryAngebotAuswahl ("REA_001_001_2020")
    
    MsgBox "TestBuildQryAngebotAuswahl() part 1 executed", vbOKOnly
    
    basBuild.BuildQryAngebotAuswahl
End Sub

Sub testLifecycleBarSingle()
    Dim intNumberOfColumns As Integer
        intNumberOfColumns = 1
        
        Dim intColumnWidth(0) As Integer
        intColumnWidth(0) = 30
        
        Dim intLeft As Integer
        intLeft = 100
        
        Dim intTop As Integer
        intTop = 2430
        
        Dim intRowHeight As Integer
        intRowHeight = 330
        
        Dim aintPositions() As Integer
        aintPositions = basSupport.CalculateLifecycleBar(intNumberOfColumns, intColumnWidth, intLeft, intTop, intRowHeight)
        
        ' event message
        If gconVerbatim Then
            Debug.Print "basTest.CreateCommandButtonSingle executed"
        End If
End Sub

Sub testLifecycleBarPair()
    Dim intNumberOfColumns As Integer
        intNumberOfColumns = 3
        
        Dim intColumnWidth(2) As Integer
        intColumnWidth(0) = 2730
        intColumnWidth(1) = 2730
        intColumnWidth(2) = 2730
        
        Dim intLeft As Integer
        intLeft = 510
        
        Dim intTop As Integer
        intTop = 1700
        
        Dim intRowHeight As Integer
        intRowHeight = 330
        
        Dim aintPositions() As Integer
        aintPositions = basSupport.CalculateLifecycleBar(intNumberOfColumns, intColumnWidth, intLeft, intTop, intRowHeight)
        
        ' event message
        If gconVerbatim Then
            Debug.Print "basTest.CreateCommandButton executed"
        End If
End Sub

Private Sub TestBuildCreateOffer()

    ' command message
    If gconVerbatim Then
        Debug.Print "basTest.TestBuildCreateOffer ausfuehren"
    End If
    
    basCreateOffer.BuildCreateOffer
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBuildCreateOffer executed"
    End If

End Sub

Private Sub TestAngebotSuchenSub_ClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basTest.TestBuildAngebotSuchenSub"
    End If
    
    ' procedure set to private now
    basAngebotSuchenSub.ClearForm "frmAngebotSuchenSub"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBuildAngebotSuchenSub executed"
    End If
    
End Sub


Private Sub TestBuildAngebotSuchenSub()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basTest.TestBuildAngebotSuchenSub"
    End If
    
    basAngebotSuchenSub.BuildAngebotSuchenSub
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBuildAngebotSuchenSub executed"
    End If
    
End Sub

Private Sub TestBuildAngebotSuchen()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basTest.TestBuildAngebotSuchen"
    End If
    
    ' basAngebotSuchenSub.BuildAngebotSuchenSub
    basAngebotSuchen.BuildAngebotSuchen
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBuildAngebotSuchen executed"
    End If
    
End Sub

Private Sub TestBasAngebotSuchen_ClearForm()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basTest.TestBasAngebotSuchen_ClearForm"
    End If
    
    ' basAngebotSuchenSub.BuildAngebotSuchenSub
    basAngebotSuchen.ClearForm "frmAngebotSuchen"
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBasAngebotSuchen_ClearForm executed"
    End If
    
End Sub

Private Sub TestBasAngebotSuchen_CalculateLifecycleGrid()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basTest.TestBasAngebotSuchen_CalculateLifecycleGrid"
    End If
    
    Dim aintArray() As Integer
    aintArray = basAngebotSuchen.CalculateLifecycleGrid
    
    Dim inti As Integer
    For inti = 0 To UBound(aintArray, 1)
        
        Dim intj As Integer
        For intj = 0 To UBound(aintArray, 2)
            Debug.Print "CalculateLifecycleGrid(" & inti & " ," & intj & ") = " & aintArray(inti, intj)
        Next
        
    Next
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basTest.TestBasAngebotSuchen_CalculateLifecycleGrid executed"
    End If
    
End Sub

Private Sub TestBasAngebotSuchen_GetLeft()

    Dim aintArray() As Integer
    aintArray = basAngebotSuchen.CalculateLifecycleGrid
     
    Dim intColumn As Integer
    intColumn = 1
    
    Debug.Print "Left (column: " & intColumn & "): " & basAngebotSuchen.GetLeft(aintArray, intColumn)
    
End Sub

Private Sub TestBasAngebotSuchen_GetTop()

    Dim aintArray() As Integer
    aintArray = basAngebotSuchen.CalculateLifecycleGrid
     
    Dim intColumn As Integer
    intColumn = 1
        
    Debug.Print "Top (column: " & intColumn & "): " & basAngebotSuchen.GetTop(aintArray, intColumn)
    
End Sub

Private Sub TestBasAngebotSuchen_GetWidth()

    Dim aintArray() As Integer
    aintArray = basAngebotSuchen.CalculateLifecycleGrid
     
    Dim intColumn As Integer
    intColumn = 1
        
    Debug.Print "Top (column: " & intColumn & "): " & basAngebotSuchen.GetWidth(aintArray, intColumn)
    
End Sub

Private Sub TestBasAngebotSuchen_GetHeight()

    Dim aintArray() As Integer
    aintArray = basAngebotSuchen.CalculateLifecycleGrid
     
    Dim intColumn As Integer
    intColumn = 1
        
    Debug.Print "Top (column: " & intColumn & "): " & basAngebotSuchen.GetHeight(aintArray, intColumn)
    
End Sub

Private Function TestBasAngebotSuchen_CalculateGrid()
    
    Dim aintGrid() As Integer
    Dim intNumberOfColumns As Integer
    Dim intNumberOfRows As Integer
    Dim intLeft As Integer
    Dim intTop As Integer
    Dim intColumnWidth As Integer
    Dim intRowHeigth As Integer
    
    intNumberOfColumns = 2
    intNumberOfRows = 2
    intLeft = 1
    intTop = 1
    intColumnWidth = 3120
    intRowHeigth = 330
    
    ReDim aintGrid(intNumberOfColumns - 1, intNumberOfRows - 1, 3)
    
    aintGrid = basAngebotSuchen.CalculateGrid(intNumberOfColumns, intNumberOfRows, intLeft, intTop, intColumnWidth, intRowHeigth)
    
    Dim bolOutput As Boolean
    bolOutput = False
    
    ' toggle output
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To intNumberOfColumns - 1
            For intj = 0 To intNumberOfRows - 1
                Debug.Print "Column " & inti & " , Row " & intj & " , Left: " & aintGrid(inti, intj, 0)
                Debug.Print "Column " & inti & " , Row " & intj & " , Top: " & aintGrid(inti, intj, 1)
                Debug.Print "Column " & inti & " , Row " & intj & " , Width: " & aintGrid(inti, intj, 2)
                Debug.Print "Column " & inti & " , Row " & intj & " , Height: " & aintGrid(inti, intj, 3)
            Next
        Next
    
    End If
    
    TestBasAngebotSuchen_CalculateGrid = aintGrid
    
End Function

Private Sub TestBasAngebotSuchen_GetLeftPlus()

    Dim aintGrid() As Integer
    aintGrid = basTest.TestBasAngebotSuchen_CalculateGrid
     
    ' toggle output
    Dim bolOutput As Boolean
    bolOutput = True
    
    If bolOutput Then
    
        Dim inti As Integer
        Dim intj As Integer
        
        For inti = 0 To UBound(aintGrid, 1)
            For intj = 0 To UBound(aintGrid, 2)
                Debug.Print "Column " & inti + 1 & " , Row " & intj + 1 & " , Left: " & basAngebotSuchen.GetLeftPlus(aintGrid, inti + 1, intj + 1)
            Next
        Next
    
    End If
    
End Sub
