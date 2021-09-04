Attribute VB_Name = "basTestLeistungserfassungsblatt"
Option Compare Database
Option Explicit

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub Leistungserfassungsblatt()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Leistungserfassungsblatt"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    ' check returned varType, feasible values are:
        ' vbEmpty           0   Empty (uninitialized)
        ' vbNull            1   Null (no valid data)
        ' vbInteger         2   Integer
        ' vbLong            3   Long integer
        ' vbSingle          4   Single-precision floating-point number
        ' vbDouble          5   Double-precision floating-point number
        ' vbCurrency        6   Currency value
        ' vbDate            7   Date value
        ' vbString          8   String
        ' vbObject          9   Object
        ' vbError           10  Error value
        ' vbBoolean         11  Boolean value
        ' vbVariant         12  Variant (used only with arrays of variants)
        ' vbDataObject      13  A data access object
        ' vbDecimal         14  Decimal value
        ' vbByte            17  Byte value
        ' vbLongLong        20  LongLong integer (valid on 64-bit platforms only)
        ' vbUserDefinedType 36  Variants that contain user-defined types
        ' vbArray           8192    Array (always added to another constant when returned by this function)
    Dim intVarType As Integer
    intVarType = 8
    
    rstRecordset.Leistungserfassungsblatt = strTestValue
    
    If Not rstRecordset.Leistungserfassungsblatt = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Leistungserfassungsblatt failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Leistungserfassungsblatt) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Leistungserfassungsblatt: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.RechnunugNr executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub RechnungNr()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.RechnungNr"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    ' check returned varType, feasible values are:
        ' vbEmpty           0   Empty (uninitialized)
        ' vbNull            1   Null (no valid data)
        ' vbInteger         2   Integer
        ' vbLong            3   Long integer
        ' vbSingle          4   Single-precision floating-point number
        ' vbDouble          5   Double-precision floating-point number
        ' vbCurrency        6   Currency value
        ' vbDate            7   Date value
        ' vbString          8   String
        ' vbObject          9   Object
        ' vbError           10  Error value
        ' vbBoolean         11  Boolean value
        ' vbVariant         12  Variant (used only with arrays of variants)
        ' vbDataObject      13  A data access object
        ' vbDecimal         14  Decimal value
        ' vbByte            17  Byte value
        ' vbLongLong        20  LongLong integer (valid on 64-bit platforms only)
        ' vbUserDefinedType 36  Variants that contain user-defined types
        ' vbArray           8192    Array (always added to another constant when returned by this function)
    Dim intVarType As Integer
    intVarType = 8
    
    rstRecordset.RechnungNr = strTestValue
    
    If Not rstRecordset.RechnungNr = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.RechnungNr failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.RechnungNr) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.RechnungNr: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.RechnunugNr executed"
    
End Sub
' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub Bemerkung()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Bemerkung"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    ' check returned varType, feasible values are:
        ' vbEmpty           0   Empty (uninitialized)
        ' vbNull            1   Null (no valid data)
        ' vbInteger         2   Integer
        ' vbLong            3   Long integer
        ' vbSingle          4   Single-precision floating-point number
        ' vbDouble          5   Double-precision floating-point number
        ' vbCurrency        6   Currency value
        ' vbDate            7   Date value
        ' vbString          8   String
        ' vbObject          9   Object
        ' vbError           10  Error value
        ' vbBoolean         11  Boolean value
        ' vbVariant         12  Variant (used only with arrays of variants)
        ' vbDataObject      13  A data access object
        ' vbDecimal         14  Decimal value
        ' vbByte            17  Byte value
        ' vbLongLong        20  LongLong integer (valid on 64-bit platforms only)
        ' vbUserDefinedType 36  Variants that contain user-defined types
        ' vbArray           8192    Array (always added to another constant when returned by this function)
    Dim intVarType As Integer
    intVarType = 8
    
    rstRecordset.Bemerkung = strTestValue
    
    If Not rstRecordset.Bemerkung = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Bemerkung) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Bemerkung: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.Bemerkung executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub BelegID()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.BelegID"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim strTestValue As String
    strTestValue = "Test"
    
    ' check returned varType, feasible values are:
        ' vbEmpty           0   Empty (uninitialized)
        ' vbNull            1   Null (no valid data)
        ' vbInteger         2   Integer
        ' vbLong            3   Long integer
        ' vbSingle          4   Single-precision floating-point number
        ' vbDouble          5   Double-precision floating-point number
        ' vbCurrency        6   Currency value
        ' vbDate            7   Date value
        ' vbString          8   String
        ' vbObject          9   Object
        ' vbError           10  Error value
        ' vbBoolean         11  Boolean value
        ' vbVariant         12  Variant (used only with arrays of variants)
        ' vbDataObject      13  A data access object
        ' vbDecimal         14  Decimal value
        ' vbByte            17  Byte value
        ' vbLongLong        20  LongLong integer (valid on 64-bit platforms only)
        ' vbUserDefinedType 36  Variants that contain user-defined types
        ' vbArray           8192    Array (always added to another constant when returned by this function)
    Dim intVarType As Integer
    intVarType = 8
    
    rstRecordset.BelegID = strTestValue
    
    If Not rstRecordset.BelegID = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.BelegID failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BelegID) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.BelegID: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.BelegID executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub Brutto()

    ' command message
    Debug.Print "execute basTestLeistungserfassungsblatt.Brutto"
    
    Dim rstRecordset As clsLeistungserfassungsblatt
    Set rstRecordset = New clsLeistungserfassungsblatt
    
    Dim curTestValue As Currency
    curTestValue = 123.45
    
    ' check returned varType, feasible values are:
        ' vbEmpty           0   Empty (uninitialized)
        ' vbNull            1   Null (no valid data)
        ' vbInteger         2   Integer
        ' vbLong            3   Long integer
        ' vbSingle          4   Single-precision floating-point number
        ' vbDouble          5   Double-precision floating-point number
        ' vbCurrency        6   Currency value
        ' vbDate            7   Date value
        ' vbString          8   String
        ' vbObject          9   Object
        ' vbError           10  Error value
        ' vbBoolean         11  Boolean value
        ' vbVariant         12  Variant (used only with arrays of variants)
        ' vbDataObject      13  A data access object
        ' vbDecimal         14  Decimal value
        ' vbByte            17  Byte value
        ' vbLongLong        20  LongLong integer (valid on 64-bit platforms only)
        ' vbUserDefinedType 36  Variants that contain user-defined types
        ' vbArray           8192    Array (always added to another constant when returned by this function)
    Dim intVarType As Integer
    intVarType = 6
        
    rstRecordset.Brutto = curTestValue
    
    If Not rstRecordset.Brutto = curTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Brutto) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Brutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.Brutto: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.Brutto executed"
    
End Sub
