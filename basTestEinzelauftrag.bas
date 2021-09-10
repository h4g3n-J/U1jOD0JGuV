Attribute VB_Name = "basTestEinzelauftrag"
Option Compare Database
Option Explicit

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub EAkurzKey()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.EAkurzKey"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
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
        MsgBox "basTestEinzelauftrag.EAkurzKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.EAkurzKey) = intVarType Then
        MsgBox "basTestEinzelauftrag.EAkurzKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.EAkurzKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.EAkurzKey executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub MengengeruestLink()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.MengengeruestLink"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim strTestValue As String
    strTestValue = "#Test#"
     
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
    
    rstRecordset.MengengeruestLink = strTestValue
    
    If Not rstRecordset.MengengeruestLink = strTestValue Then
        MsgBox "basTestEinzelauftrag.MengengeruestLink failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.MengengeruestLink) = intVarType Then
        MsgBox "basTestEinzelauftrag.MengengeruestLink failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.MengengeruestLink: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.MengengeruestLink executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub LeistungsbeschreibungLink()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.LeistungsbeschreibungLink"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim strTestValue As String
    strTestValue = "#Test#"
     
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
    
    rstRecordset.LeistungsbeschreibungLink = strTestValue
    
    If Not rstRecordset.LeistungsbeschreibungLink = strTestValue Then
        MsgBox "basTestEinzelauftrag.LeistungsbeschreibungLink failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.LeistungsbeschreibungLink) = intVarType Then
        MsgBox "basTestEinzelauftrag.LeistungsbeschreibungLink failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.LeistungsbeschreibungLink: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.LeistungsbeschreibungLink executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub Bemerkung()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.Bemerkung"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
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
        MsgBox "basTestEinzelauftrag.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Bemerkung) = intVarType Then
        MsgBox "basTestEinzelauftrag.Bemerkung failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.Bemerkung: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.Bemerkung executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub BeauftragtDatum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.BeauftragtDatum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.BeauftragtDatum = datTestValue
    
    If Not rstRecordset.BeauftragtDatum = datTestValue Then
        MsgBox "basTestEinzelauftrag.BeauftragtDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BeauftragtDatum) = intVarType Then
        MsgBox "basTestEinzelauftrag.BeauftragtDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.BeauftragtDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.BeauftragtDatum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AbgebrochenDatum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AbgebrochenDatum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.AbgebrochenDatum = datTestValue
    
    If Not rstRecordset.AbgebrochenDatum = datTestValue Then
        MsgBox "basTestEinzelauftrag.AbgebrochenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgebrochenDatum) = intVarType Then
        MsgBox "basTestEinzelauftrag.AbgebrochenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AbgebrochenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AbgebrochenDatum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub MitzeichnungI21Datum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.MitzeichnungI21Datum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.MitzeichnungI21Datum = datTestValue
    
    If Not rstRecordset.MitzeichnungI21Datum = datTestValue Then
        MsgBox "basTestEinzelauftrag.MitzeichnungI21Datum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.MitzeichnungI21Datum) = intVarType Then
        MsgBox "basTestEinzelauftrag.MitzeichnungI21Datum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.MitzeichnungI21Datum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.MitzeichnungI21Datum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub MitzeichnungI25Datum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.MitzeichnungI25Datum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.MitzeichnungI25Datum = datTestValue
    
    If Not rstRecordset.MitzeichnungI25Datum = datTestValue Then
        MsgBox "basTestEinzelauftrag.MitzeichnungI25Datum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.MitzeichnungI25Datum) = intVarType Then
        MsgBox "basTestEinzelauftrag.MitzeichnungI25Datum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.MitzeichnungI25Datum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.MitzeichnungI25Datum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AngebotDatum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AngebotDatum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.AngebotDatum = datTestValue
    
    If Not rstRecordset.AngebotDatum = datTestValue Then
        MsgBox "basTestEinzelauftrag.AngebotDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AngebotDatum) = intVarType Then
        MsgBox "basTestEinzelauftrag.AngebotDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AngebotDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AngebotDatum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AbgenommenDatum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AbgenommenDatum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.AbgenommenDatum = datTestValue
    
    If Not rstRecordset.AbgenommenDatum = datTestValue Then
        MsgBox "basTestEinzelauftrag.AbgenommenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgenommenDatum) = intVarType Then
        MsgBox "basTestEinzelauftrag.AbgenommenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AbgenommenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AbgenommenDatum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub StorniertDatum()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.StorniertDatum"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.StorniertDatum = datTestValue
    
    If Not rstRecordset.StorniertDatum = datTestValue Then
        MsgBox "basTestEinzelauftrag.StorniertDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.StorniertDatum) = intVarType Then
        MsgBox "basTestEinzelauftrag.StorniertDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.StorniertDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.StorniertDatum executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AngebotBrutto()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AngebotBrutto"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim curTestValue As Date
    curTestValue = 12.34
     
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
    
    rstRecordset.AngebotBrutto = curTestValue
    
    If Not rstRecordset.AngebotBrutto = curTestValue Then
        MsgBox "basTestEinzelauftrag.AngebotBrutto failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AngebotBrutto) = intVarType Then
        MsgBox "basTestEinzelauftrag.AngebotBrutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AngebotBrutto: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AngebotBrutto executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub BWIKey()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.BWIKey"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim strTestValue As String
    strTestValue = "test"
     
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
        MsgBox "basTestEinzelauftrag.BWIKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BWIKey) = intVarType Then
        MsgBox "basTestEinzelauftrag.BWIKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.BWIKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.BWIKey executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AftrBeginn()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AftrBeginn"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.AftrBeginn = datTestValue
    
    If Not rstRecordset.AftrBeginn = datTestValue Then
        MsgBox "basTestEinzelauftrag.AftrBeginn failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AftrBeginn) = intVarType Then
        MsgBox "basTestEinzelauftrag.AftrBeginn failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AftrBeginn: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AftrBeginn executed"
    
End Sub

' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AftrEnde()

    ' command message
    Debug.Print "execute basTestEinzelauftrag.AftrEnde"
    
    Dim rstRecordset As clsEinzelauftrag
    Set rstRecordset = New clsEinzelauftrag
    
    Dim datTestValue As Date
    datTestValue = "07.09.2021"
     
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
    intVarType = 7
    
    rstRecordset.AftrEnde = datTestValue
    
    If Not rstRecordset.AftrEnde = datTestValue Then
        MsgBox "basTestEinzelauftrag.AftrEnde failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AftrEnde) = intVarType Then
        MsgBox "basTestEinzelauftrag.AftrEnde failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestEinzelauftrag.AftrEnde: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestEinzelauftrag.AftrEnde executed"
    
End Sub
