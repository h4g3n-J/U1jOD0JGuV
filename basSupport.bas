Attribute VB_Name = "basSupport"
' basSupport

Option Compare Database
Option Explicit

' enables to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
Public Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    Dim varOutput As Variant
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.CheckDataType ausfuehren, varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
        ' error message
        Debug.Print "basSupport.CheckDataType: varInput ist null"
        CheckDataType = varInput
        Exit Function
    End If
    
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
        Debug.Print "basSupport.ObjectExists ausfuehren, strObjectName = " & strObjectName & " , strModus = " & strModus
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
        Debug.Print "basSupport.RecordsetExists ausfuehren, varTblName = " & varTblName & " , varFieldName = " & " , varRecordsetName = " & varRecordsetName
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
        strTableBInputDialogTitle, strTableAssistanceName As String, bolVerbatim As Boolean) As String
        
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
    
    If bolVerbatim = True Then
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
            If bolVerbatim = True Then
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
            If bolVerbatim = True Then
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

' creates a new recordset in a table with parent character
' it checks if the referenced table exists, the input box is empty
' or the recordset already exists
' returns the name of the created recordset
Public Sub AddRecordsetParent(ByVal _
    strTableName, strKeyColumn, strRecordsetName, strArtifact, strDialogMessage, strDialogTitle As String)
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.AddRecordsetParent ausfuehren"
    End If

    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb

    Dim rstRecordset As DAO.RecordSet

    ' Dim strRecordsetName As String

    ' debug message: return input
    If bolVerbatim = True Then
        Debug.Print "strTableName: " & strTableName & vbCrLf _
            & "strKeyColumn: " & strKeyColumn & vbCrLf _
            & "strArtifact: " & strArtifact & vbCrLf _
            & "strDialogMessage: " & strDialogMessage & vbCrLf _
            & "strDialogTitle: " & strDialogTitle & vbCrLf _
            & "bolVerbatim: " & CStr(bolVerbatim)
    End If

    ' check if table exists
    ' If basSupport.TabelleExistiert(strTableName) = False Then
    If basSupport.ObjectExists(strTableName, "table") = False Then
        Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    Else:
        If bolVerbatim = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert."
        End If
    End If

    ' ask for the recordset name
    ' strRecordsetName = InputBox(strDialogMessage, strDialogTitle)

        ' check if recordset name is empty, if true then messagebox + exit procedure
        If basSupport.PflichtfeldIstLeer(strRecordsetName) = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strArtifact & " ist leer. Prozedur abgebrochen."
            MsgBox strArtifact & " ist leer. Prozedur wird abgebrochen.", vbCritical, "Fehler"
            GoTo ExitProc
        End If

        ' debug message: return input
        If gconVerbatim = True Then
            Debug.Print "basSupport.AddrecordsetParent: RecordsetName = " & strRecordsetName
        End If

        ' check if recordset exists
        If basSupport.RecordsetExists(strTableName, strKeyColumn, strRecordsetName) = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strRecordsetName & " existiert bereits. Prozedur abgebrochen."
            MsgBox strRecordsetName & " existiert bereits. Prozedur abgebrochen.", vbCritical, "Doppelter Eintrag"
            GoTo ExitProc
        End If

        ' create recordset
        Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)

            rstRecordset.AddNew
                rstRecordset.Fields(strKeyColumn) = strRecordsetName
            rstRecordset.Update
            
        ' confirmation message
        MsgBox strRecordsetName & " erzeugt.", vbOKOnly, "Datensatz erstellen"

        ' return value (deactivated)
        ' AddRecordsetParent = strRecordsetName

    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Public Function FindItemArray(ByVal avarArray, varItem As Variant) As Variant
    
    Dim intLoop As Integer
    intLoop = LBound(avarArray, 2)
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.FindItemArray ausfuehren, varItem = " & varItem
    End If
    
    Do While avarArray(0, intLoop) <> varItem
        If intLoop = UBound(avarArray, 2) Then
            Debug.Print "basSupport.FindItemArray: '" & varItem & "' im übergebenen Array nicht gefunden"
            FindItemArray = Null
            Exit Function
        Else
            intLoop = intLoop + 1
        End If
    Loop
    
    If gconVerbatim = True Then
        Debug.Print "basSupport.FindItemArray: intLoop = " & intLoop
    End If
    
    FindItemArray = intLoop
End Function
