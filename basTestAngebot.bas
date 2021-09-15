Attribute VB_Name = "basTestAngebot"
Option Compare Database
Option Explicit

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
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
        MsgBox "basTestAngebot.BWIKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BWIKey) = intVarType Then
        MsgBox "basTestAngebot.BWIKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.BWIKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.BWIKey executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
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
        MsgBox "basTestAngebot.EAkurzKey failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.EAkurzKey) = intVarType Then
        MsgBox "basTestAngebot.EAkurzKey failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.EAkurzKey: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.EAkurzKey executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub LeistungsbeschreibungLink()

    ' command message
    Debug.Print "execute basTestAngebot.LeistungsbeschreibungLink"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
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
        MsgBox "basTestAngebot.LeistungsbeschreibungLink failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.LeistungsbeschreibungLink) = intVarType Then
        MsgBox "basTestAngebot.LeistungsbeschreibungLink failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.LeistungsbeschreibungLink: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.LeistungsbeschreibungLink executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub Bemerkung()

    ' command message
    Debug.Print "execute basTestAngebot.Bemerkung"
    
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
    
    rstRecordset.Bemerkung = strTestValue
    
    If Not rstRecordset.Bemerkung = strTestValue Then
        MsgBox "basTestAngebot.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Bemerkung) = intVarType Then
        MsgBox "basTestAngebot.Bemerkung failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.Bemerkung: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.Bemerkung executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub BeauftragtDatum()

    ' command message
    Debug.Print "execute basTestAngebot.BeauftragtDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.BeauftragtDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BeauftragtDatum) = intVarType Then
        MsgBox "basTestAngebot.BeauftragtDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.BeauftragtDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.BeauftragtDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AbgebrochenDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AbgebrochenDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.AbgebrochenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgebrochenDatum) = intVarType Then
        MsgBox "basTestAngebot.AbgebrochenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AbgebrochenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AbgebrochenDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AngebotDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AngebotDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.AngebotDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AngebotDatum) = intVarType Then
        MsgBox "basTestAngebot.AngebotDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AngebotDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AngebotDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AbgenommenDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AbgenommenDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.AbgenommenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgenommenDatum) = intVarType Then
        MsgBox "basTestAngebot.AbgenommenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AbgenommenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AbgenommenDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AftrBeginn()

    ' command message
    Debug.Print "execute basTestAngebot.AftrBeginn"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.AftrBeginn failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AftrBeginn) = intVarType Then
        MsgBox "basTestAngebot.AftrBeginn failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AftrBeginn: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AftrBeginn executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AftrEnde()

    ' command message
    Debug.Print "execute basTestAngebot.AftrEnde"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.AftrEnde failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AftrEnde) = intVarType Then
        MsgBox "basTestAngebot.AftrEnde failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AftrEnde: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AftrEnde executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub StorniertDatum()

    ' command message
    Debug.Print "execute basTestAngebot.StorniertDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim datTestValue As Date
    datTestValue = "04.09.2021"
     
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
        MsgBox "basTestAngebot.StorniertDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.StorniertDatum) = intVarType Then
        MsgBox "basTestAngebot.StorniertDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.StorniertDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.StorniertDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expectation
Private Sub AngebotBrutto()

    ' command message
    Debug.Print "execute basTestAngebot.AngebotBrutto"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim curTestValue As Currency
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
        MsgBox "basTestAngebot.AngebotBrutto failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AngebotBrutto) = intVarType Then
        MsgBox "basTestAngebot.AngebotBrutto failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestAngebot.AngebotBrutto: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestAngebot.AngebotBrutto executed"
    
End Sub
