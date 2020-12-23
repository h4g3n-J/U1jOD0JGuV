Attribute VB_Name = "basSupport"
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

' Prueft ob Abfrage existiert
Public Function AbfrageExistiert(ByVal strAbfrageName As String) As Boolean
    
    If dbsCurrentDb = Nothing Then
        Dim dbsCurrentDb As DAO.Database
        Set dbsCurrentDb = CurrentDb
    End If
    
    Dim RecordSet As Object
    
    ' Default: False, wird zu True, wenn die gesuchte Abfrage
    ' gefunden wurde
    Dim bolQueryExists As Boolean
    bolQueryExists = False
    
    For Each RecordSet In dbsCurrentDb.QueryDefs
        If RecordSet.Name = strAbfrageName Then
            bolQueryExists = True
        End If
    Next RecordSet
        
    AbfrageExistiert = bolQueryExists
    
ExitProc:
    dbsCurrentDb.Close
    Set dbsCurrentDb = Nothing
End Function

' Prueft ob Tabelle existiert
Public Function TabelleExistiert(ByVal strTabelleName As String) As Boolean
    Dim dbsCurrentDb As DAO.Database
    Set dbsCurrentDb = CurrentDb
    
    Dim RecordSet As Object
    
    ' Default: False, wird zu True, wenn die gesuchte Abfrage
    ' gefunden wurde
    Dim bolTableExists As Boolean
    bolTableExists = False
    
    For Each RecordSet In dbsCurrentDb.TableDefs
        If RecordSet.Name = strTabelleName Then
            bolTableExists = True
        End If
    Next RecordSet
    
    TabelleExistiert = bolTableExists
    
ExitProc:
    dbsCurrentDb.Close
    Set dbsCurrentDb = Nothing
End Function

' Prueft ob Pflichtfeld befuellt wurde
Public Function PflichtfeldIstLeer(ByVal varInput As Variant) As Boolean
    Dim bolStatus As Boolean
    bolStatus = True
    
    If Not (IsEmpty(varInput)) And varInput <> "" Then
        bolStatus = False
    End If
    
    PflichtfeldIstLeer = bolStatus
End Function

' Prueft ob Datensatz existiert
Public Function RecordsetExists(ByVal varTblName As Variant, ByVal varFieldName As Variant, ByVal varRecordsetName As Variant) As Boolean
    Dim bolStatus As Boolean
    bolStatus = False
    
    varTblName = CStr(varTblName)
    varFieldName = CStr(varFieldName)
    varRecordsetName = CStr(varRecordsetName)
    
    Dim dbsCurrentDb As DAO.Database
    Set dbsCurrentDb = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDb.OpenRecordset(varTblName, dbOpenDynaset)
    
    If DCount(varFieldName, varTblName, varFieldName & " Like '" & varRecordsetName & "'") > 0 Then
        bolStatus = True
    End If
    
    RecordsetExists = bolStatus
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDb.Close
    Set dbsCurrentDb = Nothing
End Function

