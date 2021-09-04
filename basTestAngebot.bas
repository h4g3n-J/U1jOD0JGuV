Attribute VB_Name = "basTestAngebot"
Option Compare Database
Option Explicit

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub BWIKey()

    ' command message
    Debug.Print "execute basTestAngebot.BWIKey"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
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
    
    rstRecordset.BWIKey = strTestValue
    
    If Not rstRecordset.BWIKey = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.BWIKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BWIKey) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.BWIKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.BWIKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.BWIKey executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub EAkurzKey()

    ' command message
    Debug.Print "execute basTestAngebot.EAkurzKey"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
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
    
    rstRecordset.EAkurzKey = strTestValue
    
    If Not rstRecordset.EAkurzKey = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.EAkurzKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.EAkurzKey) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.EAkurzKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.EAkurzKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.EAkurzKey executed"
    
End Sub
