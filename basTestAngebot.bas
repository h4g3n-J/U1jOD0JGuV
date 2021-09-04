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

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
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
        MsgBox "basTestLeistungserfassungsblatt.LeistungsbeschreibungLink failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.LeistungsbeschreibungLink) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.LeistungsbeschreibungLink failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.LeistungsbeschreibungLink: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.LeistungsbeschreibungLink executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
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
        MsgBox "basTestLeistungserfassungsblatt.Bemerkung failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.Bemerkung) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.Bemerkung failed, Error Code: 2", vbCritical, "Test Result"
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
Private Sub BeauftragtDatum()

    ' command message
    Debug.Print "execute basTestAngebot.BeauftragtDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim strTestValue As String
    strTestValue = "04.09.2021"
     
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
    
    rstRecordset.BeauftragtDatum = strTestValue
    
    If Not rstRecordset.BeauftragtDatum = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.BeauftragtDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.BeauftragtDatum) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.BeauftragtDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.BeauftragtDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.BeauftragtDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub AbgebrochenDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AbgebrochenDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim strTestValue As String
    strTestValue = "04.09.2021"
     
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
    
    rstRecordset.AbgebrochenDatum = strTestValue
    
    If Not rstRecordset.AbgebrochenDatum = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.AbgebrochenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgebrochenDatum) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.AbgebrochenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.AbgebrochenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.AbgebrochenDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub AngebotDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AngebotDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim strTestValue As String
    strTestValue = "04.09.2021"
     
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
    
    rstRecordset.AngebotDatum = strTestValue
    
    If Not rstRecordset.AngebotDatum = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.AngebotDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AngebotDatum) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.AngebotDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.AngebotDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.AngebotDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub AbgenommenDatum()

    ' command message
    Debug.Print "execute basTestAngebot.AbgenommenDatum"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim strTestValue As String
    strTestValue = "04.09.2021"
     
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
    
    rstRecordset.AbgenommenDatum = strTestValue
    
    If Not rstRecordset.AbgenommenDatum = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.AbgenommenDatum failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AbgenommenDatum) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.AbgenommenDatum failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.AbgenommenDatum: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.AbgenommenDatum executed"
    
End Sub

' checks property procedures
' Error Code: 1 - returned value does not match the input value
' Error Code: 2 - returned data type does not match the expection
Private Sub AftrBeginn()

    ' command message
    Debug.Print "execute basTestAngebot.AftrBeginn"
    
    Dim rstRecordset As clsAngebot
    Set rstRecordset = New clsAngebot
    
    Dim strTestValue As String
    strTestValue = "04.09.2021"
     
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
    
    rstRecordset.AftrBeginn = strTestValue
    
    If Not rstRecordset.AftrBeginn = strTestValue Then
        MsgBox "basTestLeistungserfassungsblatt.AftrBeginn failed, Error Code: 1", vbCritical, "Test Result"
        Exit Sub
    End If
    
    If Not VarType(rstRecordset.AftrBeginn) = intVarType Then
        MsgBox "basTestLeistungserfassungsblatt.AftrBeginn failed, Error Code: 2", vbCritical, "Test Result"
        Exit Sub
    End If
    
    MsgBox "basTestLeistungserfassungsblatt.AftrBeginn: Procedure successful", vbOKOnly, "Test Result"
    
ExitProc:
    Set rstRecordset = Nothing
    
    ' event message
    Debug.Print "basTestLeistungserfassungsblatt.AftrBeginn executed"
    
End Sub
