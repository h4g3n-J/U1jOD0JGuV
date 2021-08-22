Attribute VB_Name = "basSupport"
' basSupport

Option Compare Database
Option Explicit

' checks if mandatory field is filled
' returns FALSE if so
Public Function PflichtfeldIstLeer(ByVal varInput As Variant) As Boolean
    Dim bolStatus As Boolean
    bolStatus = True
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "basSupport.PflichtfeldIstLeer ausfuehren, varInput = " & varInput
    End If
    
    MsgBox "evoked basSupport.PflichtfeldIstLeer", vbCritical, "Warning"
    
    If Not (IsEmpty(varInput)) And varInput <> "" Then
        bolStatus = False
    End If
    
    PflichtfeldIstLeer = bolStatus
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
    
    MsgBox "evoked basSupport.AddRecordsetMN", vbCritical, "Warning"
    
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
