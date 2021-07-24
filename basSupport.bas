Attribute VB_Name = "basSupport"
' basSupport

Option Compare Database
Option Explicit

' enables to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Public Function CheckDataType(ByVal varInput, varMode As Variant) As Variant
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.CheckDataType ausfuehren"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "basSupport.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
    
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "basSupport.CheckDataType: varInput ist null"
        End If
        
        ' return input
        CheckDataType = varInput
        Exit Function
    End If
    
    ' declare output
    Dim varOutput As Variant
    
    Select Case varMode
        Case "string"
            varOutput = (varInput)
        Case "link"
            ' check if varInput is already in link format (#...#),
            ' if not convert to link format
            If Left(varInput, 1) = "#" And Right(varInput, 1) = "#" Then
                varOutput = CStr(varInput)
            Else
                varOutput = "#" + CStr(varInput) + "#"
            End If
        Case "date"
            varOutput = CDate(varInput)
        Case "currency"
            varOutput = CCur(varInput)
        Case "integer"
            varOutput = CInt(varInput)
        Case "boolean"
            varOutput = CBool(varInput)
    End Select
    
    CheckDataType = varOutput
End Function

' checks if a specific table or query exists
' returns true if positive
' strModus feasible values: "table", "query"
Public Function ObjectExists(ByVal strObjectName, strModus As String) As Boolean
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim RecordSet As Object
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.ObjectExists ausfuehren"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "basSupport.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
    End If
    
    Select Case strModus
        Case "table"
            For Each RecordSet In dbsCurrentDB.TableDefs
                If RecordSet.Name = strObjectName Then
                    bolObjectExists = True
                End If
            Next RecordSet
        Case "query"
            For Each RecordSet In dbsCurrentDB.QueryDefs
                If RecordSet.Name = strObjectName Then
                    bolObjectExists = True
                End If
            Next RecordSet
    End Select
    
    ObjectExists = bolObjectExists
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Function

' checks if mandatory field is filled
' returns FALSE if so
Public Function PflichtfeldIstLeer(ByVal varInput As Variant) As Boolean
    Dim bolStatus As Boolean
    bolStatus = True
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.PflichtfeldIstLeer ausfuehren, varInput = " & varInput
    End If
    
    If Not (IsEmpty(varInput)) And varInput <> "" Then
        bolStatus = False
    End If
    
    PflichtfeldIstLeer = bolStatus
End Function

' checks if recordset exists
' returns TRUE if so
Public Function RecordsetExists(ByVal varTblName As Variant, ByVal varFieldName As Variant, ByVal varRecordsetName As Variant) As Boolean
    Dim bolStatus As Boolean
    bolStatus = False
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.RecordsetExists ausfuehren"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "basSupport.RecordsetExists: varTblName = " & varTblName & " , varFieldName = " & " , varRecordsetName = " & varRecordsetName
    End If
    
    varTblName = CStr(varTblName)
    varFieldName = CStr(varFieldName)
    varRecordsetName = CStr(varRecordsetName)
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(varTblName, dbOpenDynaset)
    
    If DCount(varFieldName, varTblName, varFieldName & " Like '" & varRecordsetName & "'") > 0 Then
        bolStatus = True
    End If
    
    RecordsetExists = bolStatus
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Function

' creates a new recordset in a m:n relationship
' checks table existence, input, reference recordset
' returns name of the new recordset
Public Function AddRecordsetMN(ByVal strTableAName, strTableAKeyColumn, strTableAArtifact, strTableAInputDialogMessage, _
    strTableAInputDialogTitle, strTableBName, strTableBKeyColumn, strTableBArtifact, strTableBInputDialogMessage, _
        strTableBInputDialogTitle, strTableAssistanceName As String, gconVerbatim As Boolean) As String
        
    MsgBox ("basSupport.AddRecordsetMN aufgerufen - aufrufende Prozedur überprüfen"), vbOKOnly
        
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.AddRecordsetMN ausfuehren"
    End If
    
    Dim astrConfig(6, 2) As String
    
        ' table A - the table on the m-side
        ' table name
            astrConfig(0, 0) = strTableAName
        ' key column name
            astrConfig(1, 0) = strTableAKeyColumn
        ' artifact name
            astrConfig(2, 0) = strTableAArtifact
        ' input dialog message
            astrConfig(3, 0) = strTableAInputDialogMessage
        ' input dialog title
            astrConfig(4, 0) = strTableAInputDialogTitle
        ' target state for RecordExists - supposed to be true
            astrConfig(5, 0) = CStr(True)
        ' recordset name - supposed to be empty
            astrConfig(6, 0) = ""
    
        ' table B - the table on the n-side
        ' table name
            astrConfig(0, 1) = strTableBName
        ' key column name
            astrConfig(1, 1) = strTableBKeyColumn
        ' artifact name
            astrConfig(2, 1) = strTableBArtifact
        ' input dialog message
            astrConfig(3, 1) = strTableBInputDialogMessage
        ' input dialog title
            astrConfig(4, 1) = strTableBInputDialogTitle
        ' target state for RecordExists - supposed to be false
            astrConfig(5, 1) = CStr(False)
        ' recordset name
            astrConfig(6, 1) = ""
    
        ' assistance table - the m:n-table
        ' table name
            astrConfig(0, 2) = strTableAssistanceName
            
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    
    Dim inti As Integer
    Dim intj As Integer
    
    If gconVerbatim = True Then
        For inti = LBound(astrConfig, 2) To UBound(astrConfig, 2)
            For intj = LBound(astrConfig, 1) To UBound(astrConfig, 1)
                Debug.Print "astrConfig(" & intj & ", " & inti & ") = " & astrConfig(intj, inti)
            Next
        Next
    End If

    ' check if tables exist
    Dim intError As Integer
    intError = 0
    
    For inti = LBound(astrConfig, 2) To UBound(astrConfig, 2)
        ' If basSupport.TabelleExistiert(astrConfig(0, inti)) = False Then
        If basSupport.ObjectExists(astrConfig(0, inti), "table", False) = False Then
            Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(0, inti) _
                & " existiert nicht."
            intError = intError + 1
        Else:
            If gconVerbatim = True Then
            Debug.Print "basSupport.AddRecordsetMN: " + astrConfig(0, inti) _
                + " existiert"
            End If
        End If
    Next
    
    If intError > 0 Then
        Debug.Print "basSupport.AddRecodset: Prozedur abgebrochen"
        GoTo ExitProc
    End If
    
    'ask for the names of the recordsets
    Dim lngi As Long
    For lngi = LBound(astrConfig, 2) To UBound(astrConfig, 2) - 1
        astrConfig(6, lngi) = InputBox(astrConfig(3, lngi), astrConfig(4, lngi))
        
            ' check if inputbox is empty, if true then messagebox + exit procedure
            If basSupport.PflichtfeldIstLeer(astrConfig(6, lngi)) = True Then
                Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(2, lngi) & " ist leer, " _
                    & "Prozedur abgebrochen"
                MsgBox astrConfig(2, lngi) & " ist leer. Prozedur wird abgebrochen.", vbCritical, "Fehler"
                GoTo ExitProc
            End If
        
            ' return input
            If gconVerbatim = True Then
                Debug.Print "basSupport.AddRecordsetMN: " & astrConfig(2, lngi) _
                    & " = " & astrConfig(6, lngi)
            End If
            
            ' check if recordset exists
            If CStr(basSupport.RecordsetExists(astrConfig(0, lngi), astrConfig(1, lngi), _
                astrConfig(6, lngi))) <> astrConfig(5, lngi) Then
                ' error message: messagebox + exit procedure
                    If astrConfig(5, lngi) = CStr(False) Then
                        Debug.Print "basSupport.AddRecordsetMN: '" & astrConfig(6, lngi) _
                            & "' existiert bereits. Prozedur abgebrochen."
                        MsgBox "'" & astrConfig(6, lngi) & "' existiert bereits.", _
                            vbCritical, "Doppelter Eintrag"
                    Else:
                        Debug.Print "basSupport.AddRecordsetMN: '" & astrConfig(6, lngi) _
                            & "' existiert nicht. Prozedur abgebrochen."
                        MsgBox "'" & astrConfig(6, lngi) & "' existiert nicht.", _
                            vbCritical, "Referenzdantensatz fehlt"
                    End If
                GoTo ExitProc
            End If
    Next
        
    ' create recordset in assistance table
    Set rstRecordset = dbsCurrentDB.OpenRecordset(astrConfig(0, 2), dbOpenDynaset)
    
        rstRecordset.AddNew
            rstRecordset.Fields(astrConfig(1, 0)) = astrConfig(6, 0)
            rstRecordset.Fields(astrConfig(1, 1)) = astrConfig(6, 1)
        rstRecordset.Update
        
        ' close rstRecordset
        rstRecordset.Close
        Set rstRecordset = Nothing
        
    ' create recordset in table B
    Set rstRecordset = dbsCurrentDB.OpenRecordset(astrConfig(0, 1), dbOpenDynaset)
    
        rstRecordset.AddNew
            rstRecordset.Fields(astrConfig(1, 1)) = astrConfig(6, 1)
        rstRecordset.Update
        
    ' confirmation message
    MsgBox astrConfig(6, 1) & " erzeugt.", vbOKOnly, "Datensatz erstellen"

        ' return table B recordset name
        AddRecordsetMN = astrConfig(6, 1)
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Function

' search varWanted in two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Public Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.FindItemInArray ausfuehren"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "basSupport.FindItemInArray: varWanted = " & varWanted
    End If
    
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "basSupport.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' array style B: (intField, intIndex)
        Case "B"
            Do While avarArray(intField, intIndex) <> varWanted
                If intIndex = UBound(avarArray, 2) Then
                    Debug.Print "basSupport.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "basSupport.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex
End Function

' delete form
' 1. check if form exists
' 2. close if form is loaded
' 3. delete form
Public Sub ClearForm(ByVal strFormName As String)
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basSupport.ClearForm ausfuehren"
    End If
    
    MsgBox ("basSupport.ClearForm aufgerufen - aufrufende Prozedur überprüfen"), vbOKOnly
    
    Dim objDummy As Object
    For Each objDummy In Application.CurrentProject.AllForms
        If objDummy.Name = strFormName Then
            
            ' check if form is loaded
            If Application.CurrentProject.AllForms.Item(strFormName).IsLoaded Then
                
                ' close form
                DoCmd.Close acForm, strFormName, acSaveYes
                
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basSupport.ClearForm: " & strFormName & " ist geoeffnet, Formular schließen"
                End If
            End If
            
            ' delete form
            DoCmd.DeleteObject acForm, strFormName
            
            ' verbatim message
            If gconVerbatim = True Then
                Debug.Print "BasSupport.ClearForm: " & strFormName & " existiert bereits, Formular loeschen"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
End Sub

' intNumberOfColumns: defines the number of columns
' aintColumnWidth: array, defines the width of each column
' intLeft: top left position
' intTop: top position
' intRowHeight: row height
' returns array: (i, 0) Left, (i, 1) Top, (i, 2) Width, (i, 3) Height
Public Function CalculateLifecycleBar(ByVal intNumberOfColumns As Integer, ByRef aintColumnWidth() As Integer, Optional ByVal intLeft As Integer = 100, Optional ByVal intTop As Integer = 2430, Optional ByVal intRowHeight = 330)

    ' command message
    If gconVerbatim Then
        Debug.Print "execute basSupport.CalculateLifecycleBar"
    End If
    
    ' set column spacing
    Const cintHorizontalSpacing As Integer = 60
    
    Const cintNumberOfProperties = 3
    
    Dim aintBarSettings() As Integer
    ReDim aintBarSettings(intNumberOfColumns, cintNumberOfProperties)
    
    ' compute cell position properties
    Dim inti As Integer
    intNumberOfColumns = intNumberOfColumns - 1 ' adjust for counting
    For inti = 0 To intNumberOfColumns
            ' set column left
            aintBarSettings(inti, 0) = intLeft + inti * (aintColumnWidth(inti) + cintHorizontalSpacing)
            ' set row top
            aintBarSettings(inti, 1) = intTop
            ' set column width
            aintBarSettings(inti, 2) = aintColumnWidth(inti)
            ' set row height
            aintBarSettings(inti, 3) = intRowHeight
    Next

    CalculateLifecycleBar = aintBarSettings
    
End Function

' get transform row and column information into position using aintTableSetting
Public Function PositionObjectInTable(ByVal objObject As Object, aintTableSetting() As Integer, intColumn As Integer, intRow As Integer) As Object
    
    ' verbatim message
    If gconVerbatim Then
        Debug.Print "basSupport.PositionObjectInTable ausfuehren"
    End If
    
    MsgBox "called basSupport.PositionObjectInTable - replace with local method", vbOKOnly
    
    If Not (TypeOf objObject Is TextBox Or TypeOf objObject Is Label Or TypeOf objObject Is CommandButton) Then
        Debug.Print "basAngebotSuchen.TextboxPosition: falscher Objekttyp uebergeben, Funktion abgebrochen"
        Exit Function
    End If
    
    objObject.Left = aintTableSetting(intColumn, intRow, 0)
    objObject.Top = aintTableSetting(intColumn, intRow, 1)
    objObject.Width = aintTableSetting(intColumn, intRow, 2)
    objObject.Height = aintTableSetting(intColumn, intRow, 3)
    
    Set PositionObjectInTable = objObject
End Function
