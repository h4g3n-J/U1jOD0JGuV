Attribute VB_Name = "basSupport"
' basSupport

Option Compare Database
Option Explicit

' Die Pruefe-Prozeduren sollen ermöglichen, dass null Werte in den
' Recordset geschrieben und von dort ausgelesen als solche ausgelesen
' werden können, gleichzeitig sollen sie verhindern, dass falsche
' Datentypen eingegeben werden können

' Prüft, ob der übergebene Wert vom Typ String ist
' und überführt ihn in diesen Typ
Public Function PruefeString(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeString = varInput
        Exit Function
    End If
    
    PruefeString = CStr(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ String ist
' und überführt ihn in diesen Typ
' nach dem Speichern muss das Formular akutalisiert werden,
' um den Link nutzen zu können
Public Function PruefeLink(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeLink = varInput
        Exit Function
    End If
    
    ' Prüfen, ob varInput bereits im Link-Format (#...#) vorliegt,
    ' wenn ja, dann nicht mit # einschließen -> verhindert ungültige
    ' Pfade (##...##)
    If Left(varInput, 1) = "#" And Right(varInput, 1) = "#" Then
        PruefeLink = CStr(varInput)
        Exit Function
    End If
        
    PruefeLink = "#" + CStr(varInput) + "#"
    
End Function

' Prüft, ob der übergebene Wert vom Typ Date ist
' und überführt ihn in diesen Typ
Public Function PruefeDatum(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeDatum = varInput
        Exit Function
    End If
    
    PruefeDatum = CDate(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Currency ist
' und überführt ihn in diesen Typ
Public Function PruefeWaehrung(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeWaehrung = varInput
        Exit Function
    End If
    
    PruefeWaehrung = CCur(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Integer ist
' und überführt ihn in diesen Typ
Public Function PruefeInteger(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "Property ist null oder leer"
        PruefeWahrung = varInput
        Exit Function
    End If
    
    PruefeInteger = CInt(varInput)
End Function

' Prüft, ob der übergebene Wert vom Typ Integer ist
' und überführt ihn in diesen Typ
Public Function PruefeBoolean(ByVal varInput As Variant) As Variant
    If IsNull(varInput) Or IsEmpty(varInput) Then
        ' Debug.Print "barSupport.PruefeBoolean: Property ist null oder Leer"
        PruefeBoolean = varInput
        Exit Function
    End If

    PruefeBoolean = CBool(varInput)
End Function

' checks if a specific table or query exists
' returns true if positive
' strModus feasible values: "table", "query"
Public Function ObjectExists(ByVal strObjectName, strModus As String, Optional ByVal bolVerbatim As Boolean = False) As Boolean
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim RecordSet As Object
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
    
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
    
    If bolVerbatim = True Then
        Debug.Print "basSupport.ObjectExists: " & strObjectName & " existiert."
    End If
    
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
' returns the name of the new recordset
Public Function AddRecordsetParent(ByVal _
    strTableName, strKeyColumn, strArtifact, strDialogMessage, strDialogTitle As String, _
    bolVerbatim As Boolean) As String

    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb

    Dim rstRecordset As DAO.RecordSet

    Dim strRecordsetName As String

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
    If basSupport.ObjectExists(strTableName, "table", False) = False Then
        Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        GoTo ExitProc
    Else:
        If bolVerbatim = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strTableName & " existiert."
        End If
    End If

    ' ask for the name of the recordset
    strRecordsetName = InputBox(strDialogMessage, strDialogTitle)

        ' check if inputbox is empty, if true then messagebox + exit procedure
        If basSupport.PflichtfeldIstLeer(strRecordsetName) = True Then
            Debug.Print "basSupport.AddrecordsetParent: " & strArtifact & " ist leer. Prozedur abgebrochen."
            MsgBox strArtifact & " ist leer. Prozedur wird abgebrochen.", vbCritical, "Fehler"
            GoTo ExitProc
        End If

        ' debug message: return input
        If bolVerbatim = True Then
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

        AddRecordsetParent = strRecordsetName

    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Function
