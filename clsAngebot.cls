VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAngebot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsAngebot

Option Compare Database
Option Explicit

' dient als Zwischenspeicher für die ID des
' mit dem Cursor ausgewählten Datensatzes im Unterformular
' wird mit der Prozedur Laden geschrieben und von der
' Prozedur Speichern abgerufen
Private mvarClipboard As Variant

' container for selected recordset, is also relevant for properties procedures
' lookup DefaultProperties for array composition
Private mvarRecordsetClipboard As Variant

' verbatim setting
Private mbolVerbatim As Boolean
    
Private Function DefaultProperties()

     mbolVerbatim = False
     
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.DefaultProperties"
    End If
    
    ' set array size according to the number of attributes
    Dim intNumberOfAttributes As Integer
    intNumberOfAttributes = 13
    
    Dim varDefaultProperties As Variant
    ReDim varDefaultProperties(intNumberOfAttributes, 2)
    varDefaultProperties(0, 0) = "Attribut Name"
        varDefaultProperties(0, 1) = "data type"
        varDefaultProperties(0, 2) = "default value"
    varDefaultProperties(1, 0) = "BWIKey"
        varDefaultProperties(1, 1) = "string"
        varDefaultProperties(1, 2) = Null
    varDefaultProperties(2, 0) = "MengengeruestLink"
        varDefaultProperties(2, 1) = "link"
        varDefaultProperties(2, 2) = Null
    varDefaultProperties(3, 0) = "LeistungsbeschreibungLink"
        varDefaultProperties(3, 1) = "link"
        varDefaultProperties(3, 2) = Null
    varDefaultProperties(4, 0) = "Bemerkung"
        varDefaultProperties(4, 1) = "string"
        varDefaultProperties(4, 2) = Null
    varDefaultProperties(5, 0) = "BeauftragtDatum"
        varDefaultProperties(5, 1) = "date"
        varDefaultProperties(5, 2) = Null
    varDefaultProperties(6, 0) = "AbgebrochenDatum"
        varDefaultProperties(6, 1) = "date"
        varDefaultProperties(6, 2) = Null
    varDefaultProperties(7, 0) = "AngebotDatum"
        varDefaultProperties(7, 1) = "date"
        varDefaultProperties(7, 2) = Null
    varDefaultProperties(8, 0) = "AbgenommenDatum"
        varDefaultProperties(8, 1) = "date"
        varDefaultProperties(8, 2) = Null
    varDefaultProperties(9, 0) = "AftrBeginn"
        varDefaultProperties(9, 1) = "date"
        varDefaultProperties(9, 2) = Null
    varDefaultProperties(10, 0) = "AftrEnde"
        varDefaultProperties(10, 1) = "date"
        varDefaultProperties(10, 2) = Null
    varDefaultProperties(11, 0) = "StorniertDatum"
        varDefaultProperties(11, 1) = "date"
        varDefaultProperties(11, 2) = Null
    varDefaultProperties(12, 0) = "AngebotBrutto"
        varDefaultProperties(12, 1) = "currency"
        varDefaultProperties(12, 2) = Null
    ' Attribute added or removed? -> dont't miss to adjust intNumberOfAttributes (above)

    DefaultProperties = varDefaultProperties
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.DefaultProperties executed"
    End If
    
End Function
    
Private Sub Class_Initialize()

    If gconVerbatim Then
        Debug.Print "Klasse clsAngebot instanziiert."
    End If
    
    ' load default value to recordset
    mvarRecordsetClipboard = DefaultProperties
    
End Sub

Private Sub Class_Terminate()
    If gconVerbatim Then
        Debug.Print "Klasse clsAngebot terminiert."
    End If
End Sub

Public Property Let BWIKey(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "BWIKey"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let BWIKey"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let BWIKey executed"
    End If
    
End Property

Public Property Get BWIKey() As Variant

    Dim strPropertyName As String
    strPropertyName = "BWIKey"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get BWIKey"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    BWIKey = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get BWIKey executed"
    End If
    
End Property

Public Property Let MengengeruestLink(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "MengengeruestLink"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let MengengeruestLink"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let MengengeruestLink executed"
    End If
    
End Property

Public Property Get MengengeruestLink() As Variant

    Dim strPropertyName As String
    strPropertyName = "MengengeruestLink"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get MengengeruestLink"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    MengengeruestLink = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get MengengeruestLink executed"
    End If
    
End Property

Public Property Let LeistungsbeschreibungLink(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "LeistungsbeschreibungLink"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let LeistungsbeschreibungLink"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let LeistungsbeschreibungLink executed"
    End If
    
End Property

Public Property Get LeistungsbeschreibungLink() As Variant

    Dim strPropertyName As String
    strPropertyName = "LeistungsbeschreibungLink"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get LeistungsbeschreibungLink"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    LeistungsbeschreibungLink = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get LeistungsbeschreibungLink executed"
    End If
    
End Property

Public Property Let Bemerkung(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let Bemerkung"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let Bemerkung executed"
    End If
    
End Property

Public Property Get Bemerkung() As Variant

    Dim strPropertyName As String
    strPropertyName = "Bemerkung"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get Bemerkung"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    Bemerkung = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get Bemerkung executed"
    End If
    
End Property

Public Property Let BeauftragtDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "BeauftragtDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let BeauftragtDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let BeauftragtDatum executed"
    End If
    
End Property

Public Property Get BeauftragtDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "BeauftragtDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get BeauftragtDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    BeauftragtDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get BeauftragtDatum executed"
    End If
    
End Property

Public Property Let AbgebrochenDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AbgebrochenDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AbgebrochenDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AbgebrochenDatum executed"
    End If
    
End Property

Public Property Get AbgebrochenDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AbgebrochenDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AbgebrochenDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AbgebrochenDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AbgebrochenDatum executed"
    End If
    
End Property

Public Property Let AngebotDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AngebotDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AngebotDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AngebotDatum executed"
    End If
    
End Property

Public Property Get AngebotDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AngebotDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AngebotDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AngebotDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AngebotDatum executed"
    End If
    
End Property

Public Property Let AbgenommenDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AbgenommenDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AbgenommenDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AbgenommenDatum executed"
    End If
    
End Property

Public Property Get AbgenommenDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "AbgenommenDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AbgenommenDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AbgenommenDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AbgenommenDatum executed"
    End If
    
End Property

Public Property Let AftrBeginn(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AftrBeginn"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AftrBeginn"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AftrBeginn executed"
    End If
    
End Property

Public Property Get AftrBeginn() As Variant

    Dim strPropertyName As String
    strPropertyName = "AftrBeginn"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AftrBeginn"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AftrBeginn = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AftrBeginn executed"
    End If
    
End Property

Public Property Let AftrEnde(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AftrEnde"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AftrEnde"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AftrEnde executed"
    End If
    
End Property

Public Property Get AftrEnde() As Variant

    Dim strPropertyName As String
    strPropertyName = "AftrEnde"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AftrEnde"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AftrEnde = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AftrEnde executed"
    End If
    
End Property

Public Property Let StorniertDatum(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "StorniertDatum"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let StorniertDatum"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let StorniertDatum executed"
    End If
    
End Property

Public Property Get StorniertDatum() As Variant

    Dim strPropertyName As String
    strPropertyName = "StorniertDatum"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get StorniertDatum"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    StorniertDatum = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get StorniertDatum executed"
    End If
    
End Property

Public Property Let AngebotBrutto(ByVal varPropertyInput As Variant)

    Dim strPropertyName As String
    strPropertyName = "AngebotBrutto"
    
    'command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Let AngebotBrutto"
    End If
    
    ' scan property array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' transform varPropertyInput into target data type (mvarRecordsetClipboard(intItemPosition, 1)) and assign to recordset (mvarRecordsetClipboard(intItemPosition, 2))
    mvarRecordsetClipboard(intItemPosition, 2) = CheckDataType(varPropertyInput, mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Let AngebotBrutto executed"
    End If
    
End Property

Public Property Get AngebotBrutto() As Variant

    Dim strPropertyName As String
    strPropertyName = "AngebotBrutto"

    ' command message
    If mbolVerbatim = True Then
        Debug.Print "execute clsAngebot Get AngebotBrutto"
    End If
    
    ' scan attribut array (mvarRecordsetClipboard) for attribut name
    Dim intItemPosition As Integer
    intItemPosition = FindItemInArray(mvarRecordsetClipboard, strPropertyName)
    
    ' get value (mvarRecordsetClipboard( intItemPosition, 2)), transform into target data type (mvarRecordsetClipboard(intItemPosition, 1))
    AngebotBrutto = CheckDataType(mvarRecordsetClipboard(intItemPosition, 2), mvarRecordsetClipboard(intItemPosition, 1))
    
    ' event message
    If mbolVerbatim = True Then
        Debug.Print "clsAngebot Get AngebotBrutto executed"
    End If
    
End Property

' if no recordset is selected, set to default
' check if query exists
' check if recordset exists
' assign recordset to mvarRecordsetClipboard
Public Sub SelectRecordset(ByVal varRecordsetName As Variant)
    
    ' set error state
    Dim bolErrorState As Boolean
    bolErrorState = False
    
    ' name of the table or query that contains the
    ' sought-after recordset
    Dim strTableName As String
    strTableName = "tblAngebot"

    ' the type of the containing table,
    ' feasible values are: table, query
    Dim strObjectType As String
    strObjectType = "table"

    ' name of the column that contains the recordset
    Dim strKeyColumn As String
    strKeyColumn = "BWIKey"

    ' define database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    ' define recordset
    Dim rstRecordset As DAO.RecordSet
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' error handler if no recordset is selected
    ' end exitProc
    If IsNull(varRecordsetName) Then
        ' verbatim message
        If bolLocalVerbatim = True Then
            Debug.Print "clsAngebot.SelectRecordset: varRecordsetName ist leer, " _
                & "Defaultwert anzeigen"
        End If
        ' default Value
        mvarRecordsetClipboard = DefaultProperties
        Exit Sub
    End If
        
    ' check if target table exists, when negative _
    ' return error + exitProc
    If ObjectExists(strTableName, strObjectType) = False Then
        Debug.Print "clsAngebot.SelectRecordset: " & strTableName & " existiert nicht. Prozedur abgebrochen."
        Exit Sub
    End If
    
    ' check if recordset exists
    ' likely obsolete, since the recordset is chosen from a recordset selection
    If RecordsetExists(strTableName, strKeyColumn, varRecordsetName) <> True Then
        Debug.Print "clsAngebot.SelectRecordset: " & varRecordsetName & " existiert nicht. Prozedur abgebrochen."
        MsgBox varRecordsetName & " existiert nicht. Prozedur abgebrochen.", vbCritical, "Datensatz nicht gefunden"
        Exit Sub
    End If

    ' select recordset from table or query
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    rstRecordset.FindFirst (strKeyColumn & " = '" & CStr(varRecordsetName) & "'")

    ' assign recordset to properties
    ' mvarRecordsetClipboard(inti, 2): property value
    ' mvarRecordsetClipboard(inti, 0): property name
    Dim inti As Integer
    For inti = LBound(mvarRecordsetClipboard, 1) + 1 To UBound(mvarRecordsetClipboard, 1)
        mvarRecordsetClipboard(inti, 2) = rstRecordset.Fields(mvarRecordsetClipboard(inti, 0))
    Next
    
    ' verbatim message: return data content from array
    ' mvarRecordsetClipboard(inti, 2): property value
    ' mvarRecordsetClipboard(inti, 0): property name
    If bolLocalVerbatim = True Then
        For inti = LBound(mvarRecordsetClipboard, 1) To UBound(mvarRecordsetClipboard, 1)
            Debug.Print "clsAngebot.SelectRecordset: " & mvarRecordsetClipboard(inti, 0) & vbTab & mvarRecordsetClipboard(inti, 2)
        Next
    End If
    
    ' clean up
    rstRecordset.Close
    Set rstRecordset = Nothing

ExitProc:
        dbsCurrentDB.Close
        Set dbsCurrentDB = Nothing
End Sub

' checks if a specific table or query exists
' returns true if positive
' strModus feasible values: "table", "query"
Private Function ObjectExists(ByVal strObjectName, strModus As String) As Boolean
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim RecordSet As Object
    
    Dim bolObjectExists As Boolean
    bolObjectExists = False
    
    ' verbatim message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.ObjectExists"
    End If
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.ObjectExists: strObjectName = " & strObjectName & " , strModus = " & strModus
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

Private Function RecordsetExists(ByVal varTblName As Variant, ByVal varFieldName As Variant, ByVal varRecordsetName As Variant) As Boolean

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.RecordsetExists"
    End If

    Dim bolStatus As Boolean
    bolStatus = False
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.RecordsetExists: varTblName = " & varTblName & " , varFieldName = " & " , varRecordsetName = " & varRecordsetName
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
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.RecordsetExists executed"
    End If
    
End Function

' search varWanted in a two dimensional array
' array style A: (intIndex, intField)
' array style B: (intField, intIndex)
' return intIndex
Private Function FindItemInArray(ByVal avarArray, varWanted As Variant, Optional ByVal intField As Integer = 0, Optional ByVal strArrayStyle As String = "A") As Variant
    
    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.FindItemInArray"
    End If
   
    Dim intIndex As Integer
    intIndex = LBound(avarArray, 2)
    
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.FindItemInArray: varWanted = " & varWanted
    End If
    
    Select Case strArrayStyle
        ' array style A: (intIndex, intField)
        Case "A"
            ' scan array until match
            Do While avarArray(intIndex, intField) <> varWanted
                If intIndex = UBound(avarArray, 1) Then
                    Debug.Print "clsAngebot.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' array style B: (intField, intIndex)
        Case "B"
            Do While avarArray(intField, intIndex) <> varWanted
                If intIndex = UBound(avarArray, 2) Then
                    Debug.Print "clsAngebot.FindItemInArray: '" & varWanted & "' im übergebenen Array nicht gefunden"
                    FindItemInArray = Null
                    Exit Function
                Else
                    intIndex = intIndex + 1
                End If
            Loop
        ' input error handler
        Case Else
            Debug.Print "clsAngebot.FindItemInArray: arrayStyle ungültig, Wertevorrat beachten"
            Exit Function
        End Select
    
    ' return intIntex
    FindItemInArray = intIndex

    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.FindItemInArray executed"
    End If
    
End Function

' allows to write null in recordsets
' varMode feasible values: string, link, date, currency, integer, boolean
' transform and return Input corresponding to selected Mode
Private Function CheckDataType(ByVal varInput, varMode As Variant) As Variant

    ' command message
    If gconVerbatim = True Then
        Debug.Print "execute clsAngebot.CheckDataType"
    End If
    
    ' local verbatim setting
    Dim bolLocalVerbatim As Boolean
    bolLocalVerbatim = False
    
    ' local verbatim message
    If bolLocalVerbatim Then
        Debug.Print "clsAngebot.CheckDataType: varInput = " & varInput & " , varMode = " & varMode
    End If
    
    If IsNull(varInput) Then
        ' local verbatim
        If bolLocalVerbatim Then
            Debug.Print "clsAngebot.CheckDataType: varInput ist null"
        End If
        
        ' return input
        CheckDataType = varInput
        
        ' event message
        Debug.Print "clsRechnung.CheckDatyType executed"
        Exit Function
    End If
    
    ' declare output
    Dim varOutput As Variant
    
    Select Case varMode
        Case "string"
            varOutput = (varInput)
        Case "link"
            ' check if varInput already is a link (#...#),
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
    
    ' event message
    If gconVerbatim = True Then
        Debug.Print "clsAngebot.CheckDataType execute"
    End If

End Function

Public Sub CreateRecordset()
    ' Error Code 1: "BWIKey" input is missing
    ' Error Code 2: "BWIKey" is taken, choose another value
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAngebot.CreateRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAngebot"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "BWIKey"
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        MsgBox "clsAngebot.CreateRecordset: mandatory value is missing, Error Code 1", vbCritical, "Error"
        Debug.Print "Error: clsAngebot.CreateRecordset aborted, Error Code 1"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    ' check if primary key value is taken
    If DCount(strPrimaryKey, strTableName, strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'") > 0 Then
        MsgBox "clsAngebot.CreateRecordset: recordset of that name is taken, Error Code 2", vbCritical, "Error"
        Debug.Print "Error: clsAngebot.CreateRecordset aborted, Error Code 2"
        rstRecordset.Close
        Set rstRecordset = Nothing
        Exit Sub
    End If
    
    rstRecordset.AddNew
        rstRecordset!BWIKey = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BWIKey"), intValuePosition)
        rstRecordset!MengengeruestLink = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "MengengeruestLink"), intValuePosition)
        rstRecordset!LeistungsbeschreibungLink = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "LeistungsbeschreibungLink"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BeauftragtDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BeauftragtDatum"), intValuePosition)
        rstRecordset!AbgebrochenDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AbgebrochenDatum"), intValuePosition)
        rstRecordset!AngebotDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AngebotDatum"), intValuePosition)
        rstRecordset!AbgenommenDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AbgenommenDatum"), intValuePosition)
        rstRecordset!AftrBeginn = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrBeginn"), intValuePosition)
        rstRecordset!AftrEnde = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrEnde"), intValuePosition)
        rstRecordset!StorniertDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "StorniertDatum"), intValuePosition)
        rstRecordset!AngebotBrutto = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AngebotBrutto"), intValuePosition)
        ' class changed? Do not miss to update the class properties!
    rstRecordset.Update
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAngebot.CreateRecordset executed"
    End If
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub

Public Sub DeleteRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAngebot.DeleteRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAngebot"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "BWIKey"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsAngebot.DeleteRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    ' delete recordset
    rstRecordset.Delete
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAngebot.DeleteRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    
End Sub

Public Sub SaveRecordset()
    ' Error Code 1: no recordset was supplied

    ' command message
    If gconVerbatim Then
        Debug.Print "execute clsAngebot.SaveRecordset"
    End If

    Dim strTableName As String
    strTableName = "tblAngebot"
    
    Dim strPrimaryKey As String
    strPrimaryKey = "BWIKey"
    
    Dim intValuePosition As Integer
    intValuePosition = 2
    
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
    
    Dim rstRecordset As DAO.RecordSet
    Set rstRecordset = dbsCurrentDB.OpenRecordset(strTableName, dbOpenDynaset)
    
    ' check primary key value
    If IsNull(mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition)) Then
        Debug.Print "Error: clsAngebot.SaveRecordset aborted, Error Code 1"
        GoTo ExitProc
    End If
        
    ' load recordset
    rstRecordset.FindFirst (strPrimaryKey & " Like '" & mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, strPrimaryKey), intValuePosition) & "'")
    
    Dim strRecordsetName As String
    strRecordsetName = rstRecordset.Name
    
    rstRecordset.Edit
        rstRecordset!MengengeruestLink = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "MengengeruestLink"), intValuePosition)
        rstRecordset!LeistungsbeschreibungLink = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "LeistungsbeschreibungLink"), intValuePosition)
        rstRecordset!Bemerkung = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "Bemerkung"), intValuePosition)
        rstRecordset!BeauftragtDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "BeauftragtDatum"), intValuePosition)
        rstRecordset!AbgebrochenDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AbgebrochenDatum"), intValuePosition)
        rstRecordset!AngebotDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AngebotDatum"), intValuePosition)
        rstRecordset!AbgenommenDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AbgenommenDatum"), intValuePosition)
        rstRecordset!AftrBeginn = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrBeginn"), intValuePosition)
        rstRecordset!AftrEnde = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AftrEnde"), intValuePosition)
        rstRecordset!StorniertDatum = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "StorniertDatum"), intValuePosition)
        rstRecordset!AngebotBrutto = mvarRecordsetClipboard(FindItemInArray(mvarRecordsetClipboard, "AngebotBrutto"), intValuePosition)
    rstRecordset.Update
    
    ' event message
    If gconVerbatim Then
        Debug.Print "clsAngebot.SaveRecordset executed"
    End If
    
ExitProc:
    rstRecordset.Close
    Set rstRecordset = Nothing
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
End Sub
