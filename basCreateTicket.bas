Attribute VB_Name = "basCreateTicket"
Option Compare Database
Option Explicit

Sub BuildCreateTicket()
    
    If gconVerbatim Then
        Debug.Print "execute basCreateOffer.CreateFormCreateOffer"
    End If
    
End Sub

' creates a new recordset in a m:n relationship
' checks table existence, input, reference recordset
' returns name of the new recordset
Public Sub CreateTicket(ByVal strOfferName As Object)
                
    ' command message
    If gconVerbatim = True Then
        Debug.Print "basSupport.AddRecordsetMN ausfuehren"
    End If
    
    Dim astrConfig(6, 2) As String
    
        ' table A - the table on the m-side
        ' table name
            Dim strTableName As String
            strTableName = "tblAuftrag"
            astrConfig(0, 0) = strTableName
        ' key column name
            Dim strTableAKeyColumn As String
            strTableAKeyColumn = "AftrID"
            astrConfig(1, 0) = strTableAKeyColumn
        ' artifact name
            Dim strOfferName As String
            astrConfig(2, 0) = strTableAArtifact
        ' ' input dialog message
        '     astrConfig(3, 0) = strTableAInputDialogMessage
        ' ' input dialog title
        '     astrConfig(4, 0) = strTableAInputDialogTitle
        ' ' target state for RecordExists - supposed to be true
        '     astrConfig(5, 0) = CStr(True)
        ' ' recordset name - supposed to be empty
        '     astrConfig(6, 0) = ""

    ' initialize database and recordset
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    Dim rstRecordset As DAO.RecordSet
    
    Dim inti As Integer
    Dim intj As Integer
    
    ' check if tables exist

    if not basSupport.ObjectExists(strTableName, "table")
        Debug.Print strTableName & " existiert nicht. Prozedur abgebrochen"
        GoTo ExitProc
    End If

    Dim intError As Integer
    intError = 0
      
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
End Sub

